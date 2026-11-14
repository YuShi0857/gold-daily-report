import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import hashlib
import datetime
import smtplib
from email.mime.text import MIMEText
import random
import string
import io
import base64
import time
import re
import json
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Alignment

# 頁面基本設定
st.set_page_config(page_title="每日工作回報系統", layout="wide", page_icon="📝")

# 隱藏右上角選單、頂部列與部署按鈕，防止員工連回公開作品集頁面
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    .stAppDeployButton {display: none;}
    </style>
""", unsafe_allow_html=True)

# 設定台灣標準時區 (UTC+8)
TW_TZ = datetime.timezone(datetime.timedelta(hours=8))

def get_tw_now():
    return datetime.datetime.now(TW_TZ)

def get_tw_today():
    return get_tw_now().date()

# ================= 標準中文標題列定義 =================
USER_HEADERS = ["登入帳號", "密碼雜湊", "真實姓名", "綁定信箱", "系統權限", "建立時間", "所屬部門"]

# 黃金部門標題
REPORT_HEADERS = [
    "回報單號", "送出時間", "回報日期", "帳號", "姓名",
    "上班時間", "下班時間", 
    "包貨開始", "包貨結束", "包貨件數", "包貨總分鐘", "單件包貨耗時(分)",
    "客服開始", "客服結束", "客服總分鐘",
    "寄貨件數", 
    "直播開始", "直播結束", "直播總分鐘",
    "有無客訴", "客訴_Who", "客訴_Where", "客訴_When", "客訴_What", "客訴_Why", "客訴_How",
    "有無8D", "8D檔名", "8D檔案編碼",
    "今日處理事項", "提案改善", "對同仁想法", "上班總分鐘"
]

COMPACT_HEADERS = [
    "回報日期", "姓名", "上下班時段(含工時)", "包貨紀錄(時段/件數/單件耗時)", 
    "寄貨件數", "客服紀錄(時段/時數)", "直播紀錄(時段/時數)", 
    "今日處理事項明細", "提案改善", "對同仁想法", "有無客訴", "回報單號"
]

COMPLAINT_HEADERS = [
    "回報日期", "填表同仁", "1.客人(Who)", "2.管道(Where)", 
    "3.訂單與客訴日(When)", "4.問題詳情(What)", "5.原因分析(Why)", 
    "6.解決與改善(How)", "有無附8D報告", "8D檔案名稱", "回報單號"
]

# 操台部門標題
CAOTAI_HEADERS = [
    "回報單號", "送出時間", "回報日期", "帳號", "姓名",
    "工作機或賴名稱", "上班時間", "下班時間", "上班總分鐘", "上班總時數(小時)",
    "直播台名稱", "直播台粉專網址", "操台時間", "操台時數(小時)",
    "總留言_點讚_分享", "直播需求", "直播需求是否達成",
    "有無客訴", "客訴_Who", "客訴_Where", "客訴_When", "客訴_What", "客訴_Why", "客訴_How",
    "直播中客訴", "操台外處理事項", "直播台整體操台感受", "主管提醒改善事項", "提案改善",
    "需求截圖檔名", "需求截圖編碼", "有無8D", "8D檔名", "8D檔案編碼"
]

CAOTAI_COMPACT_HEADERS = [
    "回報日期", "姓名", "工作機/賴名稱", "上下班時間(當日工時)",
    "直播台名稱", "操台時間(時數)", "留言/點讚/分享",
    "直播需求", "是否達成", "有無客訴", "操台外處理事項",
    "整體操台感受", "主管提醒改善事項", "提案改善", "回報單號"
]

CAOTAI_COMPLAINT_HEADERS = [
    "回報日期", "填表同仁", "1.操台人員(Who)", "2.直播台名稱(Where)", 
    "3.發生時間點_賴群反映(When)", "4.客訴問題詳情(What)", "5.背後原因分析(Why)", 
    "6.處理方式與結果(How)", "有無附8D報告", "8D檔案名稱", "回報單號"
]

EN_TO_ZH_MAP = {
    "report_id": "回報單號", "submit_time": "送出時間", "report_date": "回報日期",
    "username": "帳號", "name": "姓名", "work_start": "上班時間", "work_end": "下班時間",
    "pack_start": "包貨開始", "pack_end": "包貨結束", "pack_count": "包貨件數",
    "pack_duration_min": "包貨總分鐘", "min_per_pack": "單件包貨耗時(分)",
    "cs_start": "客服開始", "cs_end": "客服結束", "cs_duration_min": "客服總分鐘",
    "ship_count": "寄貨件數", "live_start": "直播開始", "live_end": "直播結束",
    "live_duration_min": "直播總分鐘", "has_complaint": "有無客訴",
    "c_who": "客訴_Who", "c_where": "客訴_Where", "c_when": "客訴_When",
    "c_what": "客訴_What", "c_why": "客訴_Why", "c_how": "客訴_How",
    "has_8d": "有無8D", "file_8d_name": "8D檔名", "file_8d_b64": "8D檔案編碼",
    "daily_tasks": "今日處理事項", "improvement_proposal": "提案改善", "peer_feedback": "對同仁想法",
    "username_u": "登入帳號", "password_hash": "密碼雜湊", "email": "綁定信箱", "role": "系統權限", "created_at": "建立時間"
}

# ================= 1. 輔助與安全計算函數 (支援多段班與跨日凌晨計算) =================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def make_auth_token(username, pass_hash):
    raw = f"{username}_{pass_hash}_gold_stay_logged_in"
    return hashlib.sha256(raw.encode()).hexdigest()[:16]

def col_num_to_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def calc_minutes(t_start, t_end):
    """自動判斷跨日：若結束時間小於開始時間（如 20:00 ~ 02:30），自動視為隔日凌晨計算"""
    dt1 = datetime.datetime.combine(get_tw_today(), t_start)
    dt2 = datetime.datetime.combine(get_tw_today(), t_end)
    if dt2 < dt1:
        dt2 += datetime.timedelta(days=1)
    return round((dt2 - dt1).total_seconds() / 60.0, 2)

def extract_hhmm(s):
    m = re.search(r'(\d{1,2}:\d{2})', str(s))
    return m.group(1) if m else ""

def calc_minutes_str(s1, s2):
    """支援單段班與多段班（多行字串）的總上班分鐘數計算，含跨日凌晨自動換算"""
    if not s1 or not s2:
        return 0
    lines1 = [x.strip() for x in str(s1).split("\n") if x.strip() and x.strip() != "無"]
    lines2 = [x.strip() for x in str(s2).split("\n") if x.strip() and x.strip() != "無"]
    if not lines1 or not lines2:
        return 0
    total_m = 0.0
    for l1, l2 in zip(lines1, lines2):
        t1_str = extract_hhmm(l1)
        t2_str = extract_hhmm(l2)
        if not t1_str or not t2_str or (t1_str == "00:00" and t2_str == "00:00"):
            continue
        try:
            t1 = datetime.datetime.strptime(t1_str, "%H:%M").time()
            t2 = datetime.datetime.strptime(t2_str, "%H:%M").time()
            total_m += calc_minutes(t1, t2)
        except:
            pass
    return round(total_m, 2)

def format_shifts_display(ws_val, we_val, sep=" / "):
    """將單段或多段上下班時間格式化為清楚易讀的顯示字串（自動標註跨日）"""
    l1 = [extract_hhmm(x) for x in str(ws_val).split("\n") if extract_hhmm(x)]
    l2 = [extract_hhmm(x) for x in str(we_val).split("\n") if extract_hhmm(x)]
    if not l1 or not l2:
        return f"{ws_val} ~ {we_val}"
    pairs = []
    for idx, (a, b) in enumerate(zip(l1, l2)):
        try:
            t1 = datetime.datetime.strptime(a, "%H:%M").time()
            t2 = datetime.datetime.strptime(b, "%H:%M").time()
            cross = "(跨日)" if t2 < t1 else ""
        except:
            cross = ""
        if len(l1) > 1:
            pairs.append(f"第{idx+1}段 {a}~{b}{cross}")
        else:
            pairs.append(f"{a}~{b}{cross}")
    return sep.join(pairs)

def parse_shifts_from_multiline_text(text):
    """供管理員後台修改多段班文字框解析使用（例如：13:00-17:00 換行 20:00-02:00）"""
    starts, ends = [], []
    total_min = 0.0
    lines = [line.strip() for line in str(text).split("\n") if line.strip()]
    for line in lines:
        matches = re.findall(r'(\d{1,2}:\d{2})', line)
        if len(matches) >= 2:
            s_t, e_t = matches[0], matches[1]
            try:
                t1 = datetime.datetime.strptime(s_t, "%H:%M").time()
                t2 = datetime.datetime.strptime(e_t, "%H:%M").time()
                s_clean = t1.strftime("%H:%M")
                e_clean = t2.strftime("%H:%M")
                starts.append(s_clean)
                ends.append(f"{e_clean}(跨日)" if t2 < t1 else e_clean)
                total_min += calc_minutes(t1, t2)
            except:
                pass
    return starts, ends, round(total_min, 2)

def parse_single_op_hours(text):
    if not text or str(text).strip() in ["無", ""]:
        return 0.0
    s = str(text)
    h_matches = re.findall(r'[\(\（]\s*(\d+(?:\.\d+)?)\s*[hH小時]+[\)\）]', s)
    if h_matches:
        return round(sum(float(x) for x in h_matches), 2)
    t_matches = re.findall(r'(\d{1,2}:\d{2})\s*[-~～至]\s*(\d{1,2}:\d{2})', s)
    if t_matches:
        total_m = sum(calc_minutes_str(t1, t2) for t1, t2 in t_matches)
        return round(total_m / 60.0, 2)
    return 0.0

def parse_caotai_hours(text):
    if not text or str(text).strip() in ["無", ""]:
        return 0.0
    lines = [line.strip() for line in str(text).split("\n") if line.strip()]
    return round(sum(parse_single_op_hours(line) for line in lines), 2)

def is_caotai_row_complaint(r):
    if str(r.get("有無客訴", "")).strip() == "有":
        return True
    old_c = str(r.get("直播中客訴", "")).strip()
    if old_c not in ["", "無", "否"]:
        return True
    return False

def parse_time_safe(t_str, default_time=datetime.time(9, 0)):
    try:
        clean_s = extract_hhmm(t_str)
        if not clean_s:
            return default_time
        return datetime.datetime.strptime(clean_s, "%H:%M").time()
    except:
        return default_time

def parse_date_safe(d_str):
    try:
        return pd.to_datetime(str(d_str).strip()).date()
    except:
        return get_tw_today()

def normalize_date_str(d_str):
    try:
        return pd.to_datetime(str(d_str).strip()).strftime("%Y-%m-%d")
    except:
        return str(d_str).strip()

def parse_vals_to_df(all_vals):
    if not all_vals or len(all_vals) <= 1:
        return pd.DataFrame()
    raw_headers = [EN_TO_ZH_MAP.get(h.strip(), h.strip()) for h in all_vals[0]]
    valid_cols, clean_headers, seen = [], [], set()
    for i, h in enumerate(raw_headers):
        if h != "" and h not in seen:
            valid_cols.append(i)
            clean_headers.append(h)
            seen.add(h)
    data = []
    for r in all_vals[1:]:
        if not any(str(cell).strip() for cell in r):
            continue
        row_padded = r + [""] * max(0, len(raw_headers) - len(r))
        data.append([row_padded[i] for i in valid_cols])
    df = pd.DataFrame(data, columns=clean_headers)
    if "回報日期" in df.columns:
        df["回報日期"] = df["回報日期"].apply(normalize_date_str)
    return df

# ================= 2. 資料庫連線與快取機制 =================
def beautify_google_sheet(ws, num_cols):
    try:
        ws.freeze(rows=1)
        end_col = col_num_to_letter(min(num_cols, 26))
        ws.format(f"A1:{end_col}1", {
            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
            "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True, "fontSize": 11},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        })
        ws.format(f"A2:{end_col}500", {
            "wrapStrategy": "WRAP",
            "verticalAlignment": "TOP"
        })
    except:
        pass

@st.cache_resource
def get_worksheets_cached():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    client = gspread.authorize(creds)
    sheet_name = st.secrets.get("sheet_name", "黃金部門工作回報系統")
    sh = client.open(sheet_name)
    
    ws_users = sh.worksheet("users")
    ws_reports = sh.worksheet("reports")
    
    try:
        ws_compact = sh.worksheet("📊 每日作業總覽")
    except:
        ws_compact = sh.add_worksheet(title="📊 每日作業總覽", rows="1000", cols="14")
        ws_compact.append_row(COMPACT_HEADERS)
        
    try:
        ws_comp = sh.worksheet("🚨 客訴與異常紀錄")
    except:
        ws_comp = sh.add_worksheet(title="🚨 客訴與異常紀錄", rows="500", cols="12")
        ws_comp.append_row(COMPLAINT_HEADERS)

    try:
        ws_caotai = sh.worksheet("caotai_reports")
    except:
        ws_caotai = sh.add_worksheet(title="caotai_reports", rows="1000", cols="36")
        ws_caotai.append_row(CAOTAI_HEADERS)
        try:
            ws_caotai.hide()
        except:
            pass

    try:
        ws_caotai_compact = sh.worksheet("🎛️ 操台作業總覽")
    except:
        ws_caotai_compact = sh.add_worksheet(title="🎛️ 操台作業總覽", rows="1000", cols="17")
        ws_caotai_compact.append_row(CAOTAI_COMPACT_HEADERS)
        beautify_google_sheet(ws_caotai_compact, len(CAOTAI_COMPACT_HEADERS))

    try:
        ws_caotai_comp = sh.worksheet("🚨 操台客訴紀錄")
    except:
        ws_caotai_comp = sh.add_worksheet(title="🚨 操台客訴紀錄", rows="500", cols="12")
        ws_caotai_comp.append_row(CAOTAI_COMPLAINT_HEADERS)
        beautify_google_sheet(ws_caotai_comp, len(CAOTAI_COMPLAINT_HEADERS))
        
    return ws_users, ws_reports, ws_compact, ws_comp, ws_caotai, ws_caotai_compact, ws_caotai_comp

@st.cache_data(ttl=180)
def load_users_df():
    ws_users, _, _, _, _, _, _ = get_worksheets_cached()
    for attempt in range(3):
        try:
            all_vals = ws_users.get_all_values()
            if all_vals and len(all_vals[0]) < 7:
                ws_users.update(raw=False, range_name="A1:G1", values=[USER_HEADERS])
                all_vals[0] = USER_HEADERS
            df = parse_vals_to_df(all_vals)
            if "帳號" in df.columns and "登入帳號" not in df.columns:
                df.rename(columns={"帳號": "登入帳號", "姓名": "真實姓名"}, inplace=True)
            if not df.empty and "所屬部門" not in df.columns:
                df["所屬部門"] = "黃金部門"
            elif not df.empty:
                df["所屬部門"] = df["所屬部門"].apply(lambda x: x.strip() if str(x).strip() != "" else "黃金部門")
            return df
        except Exception as e:
            if attempt == 2:
                raise e
            time.sleep(1.5)

@st.cache_data(ttl=180)
def load_reports_df():
    _, ws_reports, _, _, _, _, _ = get_worksheets_cached()
    for attempt in range(3):
        try:
            return parse_vals_to_df(ws_reports.get_all_values())
        except Exception as e:
            if attempt == 2:
                raise e
            time.sleep(1.5)

@st.cache_data(ttl=180)
def load_caotai_df():
    _, _, _, _, ws_caotai, _, _ = get_worksheets_cached()
    for attempt in range(3):
        try:
            return parse_vals_to_df(ws_caotai.get_all_values())
        except Exception as e:
            if attempt == 2:
                raise e
            time.sleep(1.5)

def clear_all_caches():
    load_users
