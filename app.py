from flask import Flask, request, render_template_string, send_file
from typing import Optional, Union, List, Dict
from difflib import SequenceMatcher
import pandas as pd
import requests
import tempfile
import html
import json
import re
import os

app = Flask(__name__)

# ===== 登入保護 =====
from flask import Response

USERNAME = "admin"
PASSWORD = "123456"

def check_auth(username, password):
    return username == USERNAME and password == PASSWORD

def authenticate():
    return Response(
        "需要登入", 401,
        {"WWW-Authenticate": 'Basic realm="Login Required"'}
    )

@app.before_request
def require_auth():
    protected_paths = ["/", "/export"]

    if request.path in protected_paths:
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()

# ========= 可調整參數 =========
REQUEST_TIMEOUT = 25
DEFAULT_MAX_RESULTS = 20
EXPORT_FILENAME = "政府標案查詢結果_V4.xlsx"

MAX_VARIANTS = 5
MAX_API_CANDIDATES = 30

# 先用目前可見的 openfun API；若未來端點改動，再補 fallback
SEARCH_API_ENDPOINTS = [
    "https://pcc-api.openfun.app/api/searchbytitle",
    "https://pcc.g0v.ronny.tw/api/searchbytitle",
]

LAST_RESULTS = []

# 你要的欄位，直接透過 columns[] 叫 API 回傳
API_COLUMNS = [
    "機關資料:機關名稱",
    "機關資料:單位名稱",
    "機關資料:機關地址",
    "機關資料:聯絡人",
    "機關資料:聯絡電話",
    "已公告資料:標案案號",
    "已公告資料:標案名稱",
    "已公告資料:公告日",
    "招標資料:原公告日",
    "招標資料:決標方式",
    "領投開標:截止投標",
    "領投開標:開標時間",
    "領投開標:是否異動招標文件",
]

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>政府標案爬蟲系統 V4（純 API）</title>
    <style>
        body {
            font-family: Arial, "Microsoft JhengHei", sans-serif;
            margin: 30px;
            line-height: 1.6;
        }
        h1 { margin-bottom: 20px; }
        textarea, input[type="number"] {
            width: 100%;
            padding: 12px;
            font-size: 16px;
            box-sizing: border-box;
        }
        textarea { height: 180px; }
        button, a.button-link {
            display: inline-block;
            margin-top: 12px;
            margin-right: 10px;
            padding: 10px 18px;
            font-size: 16px;
            cursor: pointer;
            text-decoration: none;
            color: #000;
            border: 1px solid #888;
            background: #f5f5f5;
            border-radius: 4px;
        }
        .grid {
            display: grid;
            grid-template-columns: 1fr 220px;
            gap: 16px;
            align-items: end;
        }
        .hint, .small {
            color: #666;
        }
        .box, .summary {
            margin-top: 20px;
            padding: 12px;
            background: #f8f8f8;
            border: 1px solid #ddd;
        }
        .error-box {
            margin-top: 20px;
            padding: 12px;
            background: #fff5f5;
            border: 1px solid #d99;
            color: #900;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 24px;
            table-layout: fixed;
        }
        th, td {
            border: 1px solid #ccc;
            padding: 8px;
            vertical-align: top;
            word-break: break-word;
            text-align: left;
        }
        th { background: #f2f2f2; }
    </style>
</head>
<body>
    <h1>政府標案爬蟲系統 V4（純 API）</h1>

    <form method="POST">
        <div class="grid">
            <div>
                <label for="queries"><b>請輸入關鍵字或整段標題（每行一筆）</b></label><br>
                <textarea id="queries" name="queries" placeholder="例如：
無人機
遙控
115年 韌性臺灣 強化各類型義消科技化訓練與精進裝備中程計畫 水下工作無人機1組">{{ queries_text }}</textarea>
            </div>
            <div>
                <label for="max_results"><b>每筆最多顯示幾筆</b></label>
                <input type="number" id="max_results" name="max_results" min="1" max="100" value="{{ max_results }}">
            </div>
        </div>

        <button type="submit">查詢</button>
        {% if results %}
        <a class="button-link" href="/PSD_PCCtool/export">匯出 Excel</a>
        {% endif %}
    </form>

    <p class="hint">
        純 API 模式：輸入正規化 → 多組查詢 → searchbytitle → 相似度排序 → 直接顯示與匯出。
    </p>

    {% if summary %}
    <div class="summary">
        <b>查詢摘要</b><br>
        輸入筆數：{{ summary.query_count }}<br>
        API 候選總數：{{ summary.candidate_count }}<br>
        成功整理筆數：{{ summary.success_count }}<br>
        失敗筆數：{{ summary.fail_count }}<br>
        實際使用 API：{{ summary.api_endpoint }}
    </div>
    {% endif %}

    {% if debug_rows %}
    <div class="box">
        <b>搜尋判斷</b>
        <table>
            <thead>
                <tr>
                    <th style="width: 220px;">輸入</th>
                    <th style="width: 90px;">模式</th>
                    <th>查詢字串</th>
                </tr>
            </thead>
            <tbody>
                {% for row in debug_rows %}
                <tr>
                    <td>{{ row["input"] }}</td>
                    <td>{{ row["mode"] }}</td>
                    <td>{{ row["queries"] }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
    {% endif %}

    {% if errors %}
    <div class="error-box">
        <b>錯誤訊息</b>
        <ul>
            {% for err in errors %}
            <li>{{ err }}</li>
            {% endfor %}
        </ul>
    </div>
    {% endif %}

    {% if results %}
    <table>
        <thead>
            <tr>
                <th style="width: 120px;">原始輸入</th>
                <th style="width: 90px;">模式</th>
                <th style="width: 70px;">分數</th>
                <th style="width: 120px;">機關名稱</th>
                <th style="width: 120px;">單位名稱</th>
                <th style="width: 180px;">機關地址</th>
                <th style="width: 120px;">聯絡人</th>
                <th style="width: 120px;">聯絡電話</th>
                <th style="width: 100px;">標案案號</th>
                <th style="width: 220px;">標案名稱</th>
                <th style="width: 140px;">決標方式</th>
                <th style="width: 90px;">公告日</th>
                <th style="width: 120px;">原公告日</th>
                <th style="width: 120px;">截止投標</th>
                <th style="width: 120px;">開標時間</th>
                <th style="width: 120px;">是否異動招標文件</th>
                <th style="width: 220px;">標案 API</th>
            </tr>
        </thead>
        <tbody>
            {% for row in results %}
            <tr>
                <td>{{ row["原始輸入"] }}</td>
                <td>{{ row["比對模式"] }}</td>
                <td>{{ row["比對分數"] }}</td>
                <td>{{ row["機關名稱"] }}</td>
                <td>{{ row["單位名稱"] }}</td>
                <td>{{ row["機關地址"] }}</td>
                <td>{{ row["聯絡人"] }}</td>
                <td>{{ row["聯絡電話"] }}</td>
                <td>{{ row["標案案號"] }}</td>
                <td>{{ row["標案名稱"] }}</td>
                <td>{{ row["決標方式"] }}</td>
                <td>{{ row["公告日"] }}</td>
                <td>{{ row["原公告日"] }}</td>
                <td>{{ row["截止投標"] }}</td>
                <td>{{ row["開標時間"] }}</td>
                <td>{{ row["是否異動招標文件"] }}</td>
                <td>
                    {% if row["標案網址"] %}
                    <a href="{{ row['標案網址'] }}" target="_blank">{{ row["標案網址"] }}</a>
                    {% endif %}
                </td>
            </tr>
            {% endfor %}
        </tbody>
    </table>
    {% endif %}
</body>
</html>
"""

# ========= 文字處理 =========

def clean_text(text):
    if not text:
        return ""
    text = html.unescape(str(text))
    text = text.replace("\xa0", " ")
    text = " ".join(text.split())
    return text.strip()


def normalize_text(text):
    text = clean_text(text).lower()

    replace_map = {
        "（": "(",
        "）": ")",
        "「": "",
        "」": "",
        "『": "",
        "』": "",
        "【": "",
        "】": "",
        "－": "-",
        "—": "-",
        "–": "-",
        "，": " ",
        "、": " ",
        "。": " ",
        "：": " ",
        ":": " ",
        "/": " ",
        "\\": " ",
        "\"": "",
        "'": "",
        "　": " ",
    }

    for old, new in replace_map.items():
        text = text.replace(old, new)

    text = re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff\-\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokenize(text):
    return [t for t in normalize_text(text).split() if t]


def detect_mode(query):
    norm = normalize_text(query)
    length = len(norm.replace(" ", ""))
    has_year = bool(re.search(r"(1\d{2}|20\d{2}|21\d{2})", norm)) or "年" in query

    if length >= 18 or has_year:
        return "title"
    return "keyword"


def build_query_variants(query, mode):
    base = clean_text(query)
    norm = normalize_text(query)
    tokens = tokenize(query)

    variants = [base, norm]

    if mode == "title":
        core_tokens = [t for t in tokens if len(t) >= 2]
        if core_tokens:
            variants.append(" ".join(core_tokens[:8]))
            variants.append(" ".join(core_tokens[-5:]))
            heavy = [t for t in core_tokens if len(t) >= 3 or any(ch.isdigit() for ch in t)]
            if heavy:
                variants.append(" ".join(heavy[:6]))
    else:
        if tokens:
            variants.append(" ".join(tokens[:3]))
            variants.append(" ".join(tokens[:2]))

    deduped = []
    seen = set()
    for v in variants:
        v = clean_text(v)
        if v and v not in seen:
            seen.add(v)
            deduped.append(v)

    return deduped[:MAX_VARIANTS]


def similarity_score(user_input, candidate_title):
    a = normalize_text(user_input)
    b = normalize_text(candidate_title)

    if not a or not b:
        return 0.0

    seq = SequenceMatcher(None, a, b).ratio()
    a_tokens = set(tokenize(user_input))
    b_tokens = set(tokenize(candidate_title))

    overlap = 0.0
    if a_tokens:
        overlap = float(len(a_tokens & b_tokens)) / float(len(a_tokens))

    contain_bonus = 0.15 if (a in b or b in a) else 0.0

    year_a = set(re.findall(r"(1\d{2}|20\d{2}|21\d{2})", a))
    year_b = set(re.findall(r"(1\d{2}|20\d{2}|21\d{2})", b))
    year_bonus = 0.1 if year_a and year_b and (year_a & year_b) else 0.0

    score = (seq * 0.55) + (overlap * 0.35) + contain_bonus + year_bonus
    return round(score, 4)


# ========= API =========

def request_json(url, params=None):
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json,text/plain,*/*",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    }
    resp = requests.get(url, params=params, headers=headers, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()

    try:
        return resp.json()
    except Exception:
        text = resp.text.strip()
        if text.startswith("{") or text.startswith("["):
            return json.loads(text)
        raise ValueError("非 JSON 回應：{0}".format(url))


def search_api(query):
    """
    用 searchbytitle 直接回 records + detail。
    """
    used_endpoint = ""
    records = []
    last_errors = []

    for endpoint in SEARCH_API_ENDPOINTS:
        try:
            params = [
                ("query", query),
                ("page", 1),
            ]
            for col in API_COLUMNS:
                params.append(("columns[]", col))

            data = request_json(endpoint, params=params)

            print("API endpoint =", endpoint)
            print("API query =", query)
            print("API response type =", type(data).__name__)

            if isinstance(data, dict):
                recs = data.get("records", [])
                print("API records count =", len(recs) if isinstance(recs, list) else "not-list")

                if isinstance(recs, list):
                    used_endpoint = endpoint
                    records = recs[:MAX_API_CANDIDATES]
                    return records, used_endpoint

            last_errors.append(f"{endpoint}：回傳格式不是預期 dict/records")

        except Exception as e:
            err = f"{endpoint}：{type(e).__name__}：{str(e)}"
            print("API ERROR =", err)
            last_errors.append(err)

    if last_errors:
        raise Exception("；".join(last_errors))

    return records, used_endpoint


# ========= 資料整理 =========

def normalize_field_key(key):
    key = clean_text(key)
    key = key.replace("：", ":")
    return key


def pick_detail_value(detail, target_name):
    """
    不依賴完整 key，改用欄位尾名比對。
    例如 target_name='截止投標'，可匹配：
    - 領投開標:截止投標
    - 已公告資料:截止投標
    - 截止投標
    """
    if not detail:
        return ""

    # 如果 detail 是字串，先嘗試轉成 dict
    if isinstance(detail, str):
        try:
            detail = json.loads(detail)
        except Exception:
            return ""

    if not isinstance(detail, dict):
        return ""

    target_name = clean_text(target_name)

    # 1. 先找完全相等
    for key, value in detail.items():
        norm_key = normalize_field_key(key)
        if norm_key == target_name:
            return clean_text(value)

    # 2. 再找冒號後半段相等
    for key, value in detail.items():
        norm_key = normalize_field_key(key)
        tail = norm_key.split(":")[-1].strip()
        if tail == target_name:
            return clean_text(value)

    # 3. 最後找包含關係
    for key, value in detail.items():
        norm_key = normalize_field_key(key)
        tail = norm_key.split(":")[-1].strip()
        if target_name in tail or tail in target_name:
            return clean_text(value)

    return ""

def record_to_row(raw_query, mode, score, record):
    detail = record.get("detail", {})
    brief = record.get("brief", {}) if isinstance(record.get("brief", {}), dict) else {}

    title = pick_detail_value(detail, "標案名稱")
    if not title:
        title = clean_text(brief.get("title", ""))

    tender_url = clean_text(record.get("tender_api_url", ""))

    return {
        "原始輸入": raw_query,
        "比對模式": mode,
        "比對分數": score,
        "機關名稱": pick_detail_value(detail, "機關名稱") or clean_text(record.get("unit_name", "")),
        "單位名稱": pick_detail_value(detail, "單位名稱"),
        "機關地址": pick_detail_value(detail, "機關地址"),
        "聯絡人": pick_detail_value(detail, "聯絡人"),
        "聯絡電話": pick_detail_value(detail, "聯絡電話"),
        "標案案號": pick_detail_value(detail, "標案案號") or clean_text(record.get("job_number", "")),
        "標案名稱": title,
        "決標方式": pick_detail_value(detail, "決標方式"),
        "公告日": pick_detail_value(detail, "公告日") or clean_text(record.get("date", "")),
        "原公告日": pick_detail_value(detail, "原公告日"),
        "截止投標": pick_detail_value(detail, "截止投標"),
        "開標時間": pick_detail_value(detail, "開標時間"),
        "是否異動招標文件": pick_detail_value(detail, "是否異動招標文件"),
        "標案網址": tender_url,
    }

def collect_rows_for_query(raw_query, max_results):
    mode = detect_mode(raw_query)
    variants = build_query_variants(raw_query, mode)

    debug_info = {
        "input": raw_query,
        "mode": mode,
        "queries": " | ".join(variants),
    }

    used_api = ""
    merged = {}

    for q in variants:
    print("TRY QUERY =", q)
    records, endpoint = search_api(q)
    
        if endpoint and not used_api:
            used_api = endpoint

        for rec in records:
            title = ""
            detail = rec.get("detail", {})
            brief = rec.get("brief", {}) if isinstance(rec.get("brief", {}), dict) else {}

            if isinstance(detail, dict):
                title = clean_text(detail.get("已公告資料:標案名稱", ""))

            if not title:
                title = clean_text(brief.get("title", ""))

            tender_api_url = clean_text(rec.get("tender_api_url", ""))
            score = similarity_score(raw_query, title)

            key = tender_api_url or "{0}-{1}".format(clean_text(rec.get("job_number", "")), title)

            if key not in merged or score > merged[key]["score"]:
                merged[key] = {
                    "record": rec,
                    "title": title,
                    "score": score,
                }

    candidates = list(merged.values())
    candidates.sort(key=lambda x: x["score"], reverse=True)

    threshold = 0.45 if mode == "title" else 0.12
    filtered = [c for c in candidates if c["score"] >= threshold]

    if not filtered:
        filtered = candidates[:max_results]
    else:
        filtered = filtered[:max_results]

    rows = []
    for c in filtered:
        row = record_to_row(raw_query, mode, c["score"], c["record"])
        rows.append(row)

    return rows, debug_info, used_api, len(candidates)


# ========= Flask =========

@app.route("/PSD_PCCtool", methods=["GET", "POST"])
def index():
    global LAST_RESULTS

    results = []
    errors = []
    summary = None
    queries_text = ""
    debug_rows = []
    max_results = DEFAULT_MAX_RESULTS
    api_endpoint_used = ""

    if request.method == "POST":
        queries_text = request.form.get("queries", "").strip()

        try:
            max_results = int(request.form.get("max_results", str(DEFAULT_MAX_RESULTS)))
        except Exception:
            max_results = DEFAULT_MAX_RESULTS

        max_results = max(1, min(max_results, 100))
        inputs = [line.strip() for line in queries_text.splitlines() if line.strip()]

        all_rows = []
        candidate_total = 0

        for raw_query in inputs:
            try:
                rows, debug_info, used_api, candidate_count = collect_rows_for_query(raw_query, max_results)
                debug_rows.append(debug_info)

                if used_api and not api_endpoint_used:
                    api_endpoint_used = used_api

                candidate_total += candidate_count

                if not rows:
                    errors.append("{0}：找不到 API 候選標案，建議改短一點或換核心名詞。".format(raw_query))
                    continue

                all_rows.extend(rows)

            except Exception as e:
                errors.append("{0}：{1}".format(raw_query, str(e)))

        # 去重：同一輸入 + 同一標案案號 + 同一標案名稱
        deduped = []
        seen = set()

        for row in all_rows:
            key = (
                row["原始輸入"],
                row["標案案號"],
                row["標案名稱"],
            )
            if key not in seen:
                seen.add(key)
                deduped.append(row)

        results = deduped
        LAST_RESULTS = results

        summary = {
            "query_count": len(inputs),
            "candidate_count": candidate_total,
            "success_count": len(results),
            "fail_count": len(errors),
            "api_endpoint": api_endpoint_used or "未偵測到可用端點",
        }

    else:
        LAST_RESULTS = []

    return render_template_string(
        HTML_TEMPLATE,
        results=results,
        errors=errors,
        summary=summary,
        queries_text=queries_text,
        debug_rows=debug_rows,
        max_results=max_results,
    )


@app.route("/PSD_PCCtool/export")
def export_excel():
    global LAST_RESULTS

    if not LAST_RESULTS:
        return """
        <html>
        <head><meta charset="UTF-8"><title>無資料</title></head>
        <body style="font-family: Arial, 'Microsoft JhengHei', sans-serif; margin: 30px;">
            <h1>目前沒有可匯出的資料</h1>
            <p>請先回到首頁查詢，再匯出 Excel。</p>
        </body>
        </html>
        """

    columns = [
        "原始輸入",
        "比對模式",
        "比對分數",
        "機關名稱",
        "單位名稱",
        "機關地址",
        "聯絡人",
        "聯絡電話",
        "標案案號",
        "標案名稱",
        "決標方式",
        "公告日",
        "原公告日",
        "截止投標",
        "開標時間",
        "是否異動招標文件",
        "標案網址",
    ]

    df = pd.DataFrame(LAST_RESULTS)
    df = df.reindex(columns=columns)

    tmp_dir = tempfile.gettempdir()
    file_path = os.path.join(tmp_dir, EXPORT_FILENAME)

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="標案資料")
        ws = writer.sheets["標案資料"]
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

        width_map = {
            "A": 24, "B": 10, "C": 10, "D": 18, "E": 18,
            "F": 28, "G": 18, "H": 18, "I": 14, "J": 34,
            "K": 18, "L": 12, "M": 18, "N": 18, "O": 42,
        }

        for col, width in width_map.items():
            ws.column_dimensions[col].width = width

    return send_file(
        file_path,
        as_attachment=True,
        download_name=EXPORT_FILENAME,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True, port=5050)
