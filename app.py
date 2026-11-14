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

# 隱藏右上角選單與部署按鈕（注意保留 header 結構避免側邊欄開關消失）
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stAppDeployButton {display: none;}
    </style>
""", unsafe_allow_html=True)

TW_TZ = datetime.timezone(datetime.timedelta(hours=8))

def get_tw_now():
    return datetime.datetime.now(TW_TZ)

def get_tw_today():
    return get_tw_now().date()

# ================= 標題列定義 =================
USER_HEADERS = ["登入帳號", "密碼雜湊", "真實姓名", "綁定信箱", "系統權限", "建立時間", "所屬部門"]

REPORT_HEADERS = [
    "回報單號", "送出時間", "回報日期", "帳號", "姓名",
    "上班時間", "下班時間", "包貨開始", "包貨結束", "包貨件數", "包貨總分鐘", "單件包貨耗時(分)",
    "客服開始", "客服結束", "客服總分鐘", "寄貨件數", "直播開始", "直播結束", "直播總分鐘",
    "有無客訴", "客訴_Who", "客訴_Where", "客訴_When", "客訴_What", "客訴_Why", "客訴_How",
    "有無8D", "8D檔名", "8D檔案編碼", "今日處理事項", "提案改善", "對同仁想法", "上班總分鐘"
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

# ================= 1. 輔助與工時計算函數 =================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def make_auth_token(username, pass_hash):
    return hashlib.sha256(f"{username}_{pass_hash}_gold_stay_logged_in".encode()).hexdigest()[:16]

def col_num_to_letter(n):
    res = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(65 + rem) + res
    return res

def calc_minutes(t_start, t_end):
    """若結束時間小於開始時間（如 20:00 ~ 02:30），自動視為隔日凌晨計算"""
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
        t1_s, t2_s = extract_hhmm(l1), extract_hhmm(l2)
        if not t1_s or not t2_s or (t1_s == "00:00" and t2_s == "00:00"):
            continue
        try:
            t1 = datetime.datetime.strptime(t1_s, "%H:%M").time()
            t2 = datetime.datetime.strptime(t2_s, "%H:%M").time()
            total_m += calc_minutes(t1, t2)
        except:
            pass
    return round(total_m, 2)

def format_shifts_display(ws_val, we_val, sep=" / "):
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
        pairs.append(f"第{idx+1}段 {a}~{b}{cross}" if len(l1) > 1 else f"{a}~{b}{cross}")
    return sep.join(pairs)

def format_shifts_for_edit(ws_val, we_val):
    l1 = [extract_hhmm(x) for x in str(ws_val).split("\n") if extract_hhmm(x)]
    l2 = [extract_hhmm(x) for x in str(we_val).split("\n") if extract_hhmm(x)]
    if not l1 or not l2:
        return "13:00-22:00"
    return "\n".join([f"{a}-{b}" for a, b in zip(l1, l2)])

def parse_shifts_from_multiline_text(text):
    starts, ends = [], []
    total_min = 0.0
    for line in [l.strip() for l in str(text).split("\n") if l.strip()]:
        matches = re.findall(r'(\d{1,2}:\d{2})', line)
        if len(matches) >= 2:
            try:
                t1 = datetime.datetime.strptime(matches[0], "%H:%M").time()
                t2 = datetime.datetime.strptime(matches[1], "%H:%M").time()
                s_c, e_c = t1.strftime("%H:%M"), t2.strftime("%H:%M")
                starts.append(s_c)
                ends.append(f"{e_c}(跨日)" if t2 < t1 else e_c)
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
        return round(sum(calc_minutes_str(t1, t2) for t1, t2 in t_matches) / 60.0, 2)
    return 0.0

def parse_caotai_hours(text):
    if not text or str(text).strip() in ["無", ""]:
        return 0.0
    lines = [line.strip() for line in str(text).split("\n") if line.strip()]
    return round(sum(parse_single_op_hours(line) for line in lines), 2)

def is_caotai_row_complaint(r):
    if str(r.get("有無客訴", "")).strip() == "有":
        return True
    return str(r.get("直播中客訴", "")).strip() not in ["", "無", "否"]

def parse_time_safe(t_str, default_time=datetime.time(9, 0)):
    try:
        clean_s = extract_hhmm(t_str)
        return datetime.datetime.strptime(clean_s, "%H:%M").time() if clean_s else default_time
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

# ================= 2. 資料庫連線與快取 =================
def beautify_google_sheet(ws, num_cols):
    try:
        ws.freeze(rows=1)
        end_col = col_num_to_letter(min(num_cols, 26))
        ws.format(f"A1:{end_col}1", {
            "backgroundColor": {"red": 0.12, "green": 0.31, "blue": 0.47},
            "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True, "fontSize": 11},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"
        })
        ws.format(f"A2:{end_col}500", {"wrapStrategy": "WRAP", "verticalAlignment": "TOP"})
    except:
        pass

@st.cache_resource
def get_worksheets_cached():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    sh = client.open(st.secrets.get("sheet_name", "黃金部門工作回報系統"))
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
        try: ws_caotai.hide()
        except: pass
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
    ws_users = get_worksheets_cached()[0]
    for attempt in range(3):
        try:
            all_vals = ws_users.get_all_values()
            if all_vals and len(all_vals[0]) < 7:
                ws_users.update(raw=False, range_name="A1:G1", values=[USER_HEADERS])
                all_vals[0] = USER_HEADERS
            df = parse_vals_to_df(all_vals)
            if "帳號" in df.columns and "登入帳號" not in df.columns:
                df.rename(columns={"帳號": "登入帳號", "姓名": "真實姓名"}, inplace=True)
            if not df.empty:
                if "所屬部門" not in df.columns:
                    df["所屬部門"] = "黃金部門"
                else:
                    df["所屬部門"] = df["所屬部門"].apply(lambda x: x.strip() if str(x).strip() else "黃金部門")
            return df
        except Exception as e:
            if attempt == 2: raise e
            time.sleep(1.5)

@st.cache_data(ttl=180)
def load_reports_df():
    ws_reports = get_worksheets_cached()[1]
    for attempt in range(3):
        try: return parse_vals_to_df(ws_reports.get_all_values())
        except Exception as e:
            if attempt == 2: raise e
            time.sleep(1.5)

@st.cache_data(ttl=180)
def load_caotai_df():
    ws_caotai = get_worksheets_cached()[4]
    for attempt in range(3):
        try: return parse_vals_to_df(ws_caotai.get_all_values())
        except Exception as e:
            if attempt == 2: raise e
            time.sleep(1.5)

def clear_all_caches():
    load_users_df.clear()
    load_reports_df.clear()
    load_caotai_df.clear()

def write_dict_to_sheet(ws, expected_headers, data_dict, target_report_id=None):
    all_vals = ws.get_all_values()
    if not all_vals:
        ws.append_row(expected_headers)
        headers = expected_headers
    else:
        headers = [EN_TO_ZH_MAP.get(h.strip(), h.strip()) for h in all_vals[0]]
        missing_cols = [h for h in expected_headers if h not in headers]
        if missing_cols:
            headers = headers + missing_cols
            ws.update(raw=False, range_name=f"A1:{col_num_to_letter(len(headers))}1", values=[headers])
    row_to_write = [data_dict.get(h, "") for h in headers]
    if target_report_id is not None:
        for idx in range(1, len(all_vals)):
            if len(all_vals[idx]) > 0 and str(all_vals[idx][0]).strip() == str(target_report_id).strip():
                r_num = idx + 1
                ws.update(raw=False, range_name=f"A{r_num}:{col_num_to_letter(len(headers))}{r_num}", values=[row_to_write])
                return True
        return False
    ws.append_row(row_to_write)
    return True

def sync_pretty_sheets(ws_reports, ws_compact, ws_comp, apply_style=False):
    df = parse_vals_to_df(ws_reports.get_all_values())
    compact_rows, comp_rows = [COMPACT_HEADERS], [COMPLAINT_HEADERS]
    if not df.empty:
        for _, r in df.sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).iterrows():
            w_hr = round(calc_minutes_str(r.get("上班時間", "09:00"), r.get("下班時間", "18:00")) / 60.0, 2)
            p_cnt = pd.to_numeric(r.get("包貨件數", 0), errors="coerce") or 0
            p_dur = pd.to_numeric(r.get("包貨總分鐘", 0), errors="coerce") or 0
            if p_dur <= 0 and p_cnt > 0:
                p_dur = calc_minutes_str(r.get("包貨開始", "無"), r.get("包貨結束", "無"))
            m_per_p = round(p_dur / p_cnt, 2) if p_cnt > 0 else 0
            cs_min = pd.to_numeric(r.get("客服總分鐘", 0), errors="coerce") or calc_minutes_str(r.get("客服開始", "無"), r.get("客服結束", "無"))
            lv_min = pd.to_numeric(r.get("直播總分鐘", 0), errors="coerce") or calc_minutes_str(r.get("直播開始", "無"), r.get("直播結束", "無"))
            pack_str = f"{r.get('包貨開始','無')}~{r.get('包貨結束','無')}\n({int(p_cnt)}件 / {m_per_p}分)" if p_cnt > 0 else "無"
            cs_str = f"{r.get('客服開始','無')}~{r.get('客服結束','無')}\n({round(cs_min/60, 1)}小時)" if cs_min > 0 else "無"
            lv_str = f"{r.get('直播開始','無')}~{r.get('直播結束','無')}\n({round(lv_min/60, 1)}小時)" if lv_min > 0 else "無"
            compact_rows.append([
                str(r.get("回報日期", "")), str(r.get("姓名", "")),
                f"{r.get('上班時間', '')} ~ {r.get('下班時間', '')}\n({w_hr}小時)",
                pack_str, str(r.get("寄貨件數", 0)), cs_str, lv_str,
                str(r.get("今日處理事項", "")), str(r.get("提案改善", "")) or "無",
                str(r.get("對同仁想法", "")) or "無", str(r.get("有無客訴", "無")), str(r.get("回報單號", ""))
            ])
            if r.get("有無客訴") == "有" or r.get("有無8D") == "是":
                comp_rows.append([
                    str(r.get("回報日期", "")), str(r.get("姓名", "")), str(r.get("客訴_Who", "")),
                    str(r.get("客訴_Where", "")), str(r.get("客訴_When", "")), str(r.get("客訴_What", "")),
                    str(r.get("客訴_Why", "")), str(r.get("客訴_How", "")), str(r.get("有無8D", "否")),
                    str(r.get("8D檔名", "")), str(r.get("回報單號", ""))
                ])
    ws_compact.clear(); ws_compact.update(compact_rows)
    ws_comp.clear(); ws_comp.update(comp_rows)
    if apply_style:
        beautify_google_sheet(ws_compact, len(COMPACT_HEADERS))
        beautify_google_sheet(ws_comp, len(COMPLAINT_HEADERS))

def sync_caotai_pretty_sheet(ws_caotai, ws_caotai_compact, ws_caotai_comp, apply_style=False):
    df = parse_vals_to_df(ws_caotai.get_all_values())
    compact_rows, comp_rows = [CAOTAI_COMPACT_HEADERS], [CAOTAI_COMPLAINT_HEADERS]
    if not df.empty:
        for _, r in df.sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).iterrows():
            w_hr = round(calc_minutes_str(r.get("上班時間", ""), r.get("下班時間", "")) / 60.0, 2)
            ct_hr = parse_caotai_hours(r.get("操台時間", ""))
            has_c = is_caotai_row_complaint(r)
            shifts_disp = format_shifts_display(r.get("上班時間", ""), r.get("下班時間", ""), sep="\n")
            compact_rows.append([
                str(r.get("回報日期", "")), str(r.get("姓名", "")), str(r.get("工作機或賴名稱", "")),
                f"{shifts_disp}\n(共 {w_hr} 小時)", str(r.get("直播台名稱", "")),
                f"{r.get('操台時間', '')}\n(共 {ct_hr} H)", str(r.get("總留言_點讚_分享", "")),
                str(r.get("直播需求", "")), str(r.get("直播需求是否達成", "")),
                "有" if has_c else "無", str(r.get("操台外處理事項", "")),
                str(r.get("直播台整體操台感受", "")), str(r.get("主管提醒改善事項", "")) or "無",
                str(r.get("提案改善", "")), str(r.get("回報單號", ""))
            ])
            if has_c or r.get("有無8D") == "是":
                comp_rows.append([
                    str(r.get("回報日期", "")), str(r.get("姓名", "")),
                    str(r.get("客訴_Who", "")) or str(r.get("姓名", "")),
                    str(r.get("客訴_Where", "")) or str(r.get("直播台名稱", "")),
                    str(r.get("客訴_When", "")),
                    str(r.get("客訴_What", "")) or str(r.get("直播中客訴", "")),
                    str(r.get("客訴_Why", "")), str(r.get("客訴_How", "")),
                    str(r.get("有無8D", "否")), str(r.get("8D檔名", "")), str(r.get("回報單號", ""))
                ])
    ws_caotai_compact.clear(); ws_caotai_compact.update(compact_rows)
    ws_caotai_comp.clear(); ws_caotai_comp.update(comp_rows)
    if apply_style:
        beautify_google_sheet(ws_caotai_compact, len(CAOTAI_COMPACT_HEADERS))
        beautify_google_sheet(ws_caotai_comp, len(CAOTAI_COMPLAINT_HEADERS))

def find_user_row_idx(ws_users, target_username):
    for idx, val in enumerate(ws_users.col_values(1)):
        if idx > 0 and val.strip() == str(target_username).strip():
            return idx + 1
    return None

def delete_report_by_id(ws_reports, ws_compact, ws_comp, target_report_id):
    col_vals = ws_reports.col_values(1)
    deleted = False
    for idx in range(len(col_vals) - 1, 0, -1):
        if str(col_vals[idx]).strip() == str(target_report_id).strip():
            ws_reports.delete_rows(idx + 1)
            deleted = True
    if deleted:
        sync_pretty_sheets(ws_reports, ws_compact, ws_comp)
        clear_all_caches()
    return deleted

def delete_caotai_by_id(ws_caotai, ws_caotai_compact, ws_caotai_comp, target_report_id):
    col_vals = ws_caotai.col_values(1)
    deleted = False
    for idx in range(len(col_vals) - 1, 0, -1):
        if str(col_vals[idx]).strip() == str(target_report_id).strip():
            ws_caotai.delete_rows(idx + 1)
            deleted = True
    if deleted:
        sync_caotai_pretty_sheet(ws_caotai, ws_caotai_compact, ws_caotai_comp)
        clear_all_caches()
    return deleted

def delete_existing_same_day_report(ws, username, report_date_str):
    all_vals = ws.get_all_values()
    overwritten = False
    target_d = normalize_date_str(report_date_str)
    for idx in range(len(all_vals) - 1, 0, -1):
        r = all_vals[idx]
        if len(r) >= 4 and normalize_date_str(r[2]) == target_d and str(r[3]).strip() == str(username).strip():
            ws.delete_rows(idx + 1)
            overwritten = True
    return overwritten

def get_quarter_label(date_str):
    try:
        dt = pd.to_datetime(date_str)
        return f"{dt.year}-Q{(dt.month - 1) // 3 + 1}"
    except:
        return "未知季度"

def style_excel_File(writer):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    for sheet_name in writer.sheets:
        ws = writer.sheets[sheet_name]
        ws.views.sheetView[0].showGridLines = True
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                if cell.row == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                lines = str(cell.value or "").split("\n")
                line_max = max([len(l.encode("utf-8")) for l in lines]) if lines else 0
                if line_max > max_len:
                    max_len = line_max
            ws.column_dimensions[col_letter].width = min(max(max_len * 0.65 + 4, 12), 45)

def send_reset_email(receiver_email, new_pass):
    try:
        sender_email = st.secrets["smtp"]["sender_email"]
        sender_password = st.secrets["smtp"]["app_password"]
        msg = MIMEText(f"您好，\n\n您的「每日工作回報系統」密碼已重設。\n臨時新密碼為：{new_pass}\n\n請登入後至左側選單「修改密碼」更新。", "plain", "utf-8")
        msg["Subject"] = "【系統通知】每日工作回報系統 - 密碼重設"
        msg["From"] = sender_email
        msg["To"] = receiver_email
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [receiver_email], msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"郵件發送失敗：{e}")
        return False

# ================= 3. 狀態管理與防掉線自動登入 =================
for k, v in {
    "logged_in": False, "user_info": {}, "task_rows": 3,
    "ct_shift_rows": 1, "ct_task_rows": 3, "ct_stream_rows": 1,
    "form_ver": 0, "submit_success_msg": "", "confirm_del_id": None, "edit_report_id": None
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

if not st.session_state.logged_in:
    saved_u = st.query_params.get("u", "")
    saved_t = st.query_params.get("t", "")
    if saved_u and saved_t:
        try:
            users_df = load_users_df()
            if not users_df.empty and saved_u in users_df["登入帳號"].values:
                urow = users_df[users_df["登入帳號"] == saved_u].iloc[0]
                if saved_t == make_auth_token(saved_u, urow["密碼雜湊"]):
                    st.session_state.logged_in = True
                    st.session_state.user_info = {
                        "username": urow["登入帳號"], "password_hash": urow["密碼雜湊"],
                        "name": urow["真實姓名"], "email": urow["綁定信箱"],
                        "role": urow["系統權限"], "dept": urow.get("所屬部門", "黃金部門")
                    }
        except:
            pass

# ================= 4. 未登入畫面 =================
if not st.session_state.logged_in:
    st.title("📝 每日工作回報系統")
    tab1, tab2, tab3 = st.tabs(["🔑 登入系統", "📝 新人註冊", "❓ 忘記密碼"])
    
    with tab1:
        st.subheader("員工 / 管理員登入")
        with st.form("login_form"):
            login_user = st.text_input("帳號")
            login_pass = st.text_input("密碼", type="password")
            login_submit = st.form_submit_button("🔑 登入系統", type="primary", use_container_width=True)
        if login_submit:
            with st.spinner("🔄 正在驗證帳號..."):
                users_df = load_users_df()
                if not users_df.empty and login_user.strip() in users_df["登入帳號"].values:
                    user_row = users_df[users_df["登入帳號"] == login_user.strip()].iloc[0]
                    if user_row["密碼雜湊"] == hash_password(login_pass.strip()):
                        st.session_state.logged_in = True
                        st.session_state.user_info = {
                            "username": user_row["登入帳號"], "password_hash": user_row["密碼雜湊"],
                            "name": user_row["真實姓名"], "email": user_row["綁定信箱"],
                            "role": user_row["系統權限"], "dept": user_row.get("所屬部門", "黃金部門")
                        }
                        st.query_params["u"] = user_row["登入帳號"]
                        st.query_params["t"] = make_auth_token(user_row["登入帳號"], user_row["密碼雜湊"])
                        st.rerun()
                    else:
                        st.error("❌ 密碼錯誤！")
                else:
                    st.error("❌ 找不到此帳號，請先切換至「📝 新人註冊」建立帳號！")

    with tab2:
        st.subheader("註冊新帳號")
        with st.form("register_form"):
            reg_dept = st.selectbox("請選擇您的所屬部門", ["黃金部門", "操台部門", "跨部門(兩部門皆可填寫)"])
            reg_name = st.text_input("真實姓名 (將自動綁定未來的回報單)")
            reg_email = st.text_input("常用信箱 (忘記密碼寄信用，請務必填寫正確)")
            reg_user = st.text_input("設定登入帳號")
            reg_pass1 = st.text_input("設定密碼", type="password")
            reg_pass2 = st.text_input("再次確認密碼", type="password")
            reg_submit = st.form_submit_button("📝 立即註冊帳號", type="primary", use_container_width=True)
        if reg_submit:
            if not all([reg_name.strip(), reg_email.strip(), reg_user.strip(), reg_pass1.strip(), reg_pass2.strip()]):
                st.warning("⚠️ 所有欄位皆為必填！")
            elif reg_pass1 != reg_pass2:
                st.error("❌ 兩次輸入的密碼不一致，請重新確認！")
            else:
                with st.spinner("🔄 正在建立您的帳號..."):
                    ws_users = get_worksheets_cached()[0]
                    users_df = load_users_df()
                    if not users_df.empty and reg_user.strip() in users_df["登入帳號"].values:
                        st.error("❌ 此帳號已被註冊，請換一個！")
                    else:
                        ws_users.append_row([
                            reg_user.strip(), hash_password(reg_pass1.strip()), reg_name.strip(),
                            reg_email.strip(), "staff", get_tw_now().strftime("%Y-%m-%d %H:%M:%S"), reg_dept
                        ])
                        clear_all_caches()
                        st.success(f"✅ 註冊成功（所屬部門：{reg_dept}）！請切換至左上角「🔑 登入系統」進行登入。")

    with tab3:
        st.subheader("忘記密碼 / 重新寄送臨時密碼")
        with st.form("fp_form"):
            fp_user = st.text_input("請輸入您的帳號")
            fp_email = st.text_input("請輸入註冊時綁定的信箱")
            fp_submit = st.form_submit_button("📧 寄送重設密碼信", type="primary")
        if fp_submit:
            with st.spinner("🔄 正在驗證信箱並發送郵件..."):
                ws_users = get_worksheets_cached()[0]
                users_df = load_users_df()
                if not users_df.empty:
                    match = users_df[(users_df["登入帳號"] == fp_user.strip()) & (users_df["綁定信箱"] == fp_email.strip())]
                    if not match.empty:
                        new_temp_pass = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
                        row_idx = find_user_row_idx(ws_users, fp_user.strip())
                        if row_idx:
                            ws_users.update_cell(row_idx, 2, hash_password(new_temp_pass))
                            clear_all_caches()
                            if send_reset_email(fp_email.strip(), new_temp_pass):
                                st.success("✅ 臨時新密碼已寄出！請至信箱收信並重新登入。")
                    else:
                        st.error("❌ 帳號與綁定信箱不符合！")
                else:
                    st.error("❌ 系統尚無使用者資料！")

# ================= 5. 已登入畫面 =================
else:
    try:
        ws_users, ws_reports, ws_compact, ws_comp, ws_caotai, ws_caotai_compact, ws_caotai_comp = get_worksheets_cached()
    except Exception:
        st.warning("⏳ 系統正在冷卻連線頻率，請等待約 10 秒後按 F5 重新整理網頁（登入狀態已鎖定保護，不會被登出）！")
        st.stop()

    user = st.session_state.user_info
    user_dept = "跨部門(兩部門皆可填寫)" if user["role"] == "admin" else user.get("dept", "黃金部門")

    st.sidebar.title(f"👋 您好，{user['name']}")
    st.sidebar.caption(f"身份：{'👑 管理員' if user['role'] == 'admin' else f'👤 {user_dept}'}")
    st.sidebar.caption(f"🕒 台灣時間：{get_tw_now().strftime('%Y-%m-%d %H:%M')}")
    st.sidebar.success("🔒 防掉線保護已啟動（重整不登出）")
    
    menu_options = []
    if user["role"] == "admin":
        menu_options.extend(["📊 主管後台數據中心", "👥 帳號與權限管理"])
    if user_dept in ["黃金部門", "跨部門(兩部門皆可填寫)"]:
        menu_options.append("📝 黃金部門 - 每日工作回報")
    if user_dept in ["操台部門", "跨部門(兩部門皆可填寫)"]:
        menu_options.append("🎛️ 操台部門 - 每日工作回報")
    menu_options.append("🔒 修改密碼")
        
    choice = st.sidebar.radio("功能選單", menu_options)
    
    if st.sidebar.button("🚪 登出系統"):
        st.session_state.logged_in = False
        st.session_state.user_info = {}
        st.query_params.clear()
        st.rerun()

    # ---------------- 功能 A1：黃金部門 - 每日工作回報 ----------------
    if choice == "📝 黃金部門 - 每日工作回報":
        st.title("✨ 黃金部門 - 每日工作回報單")
        if st.session_state.submit_success_msg:
            st.balloons()
            st.success(st.session_state.submit_success_msg)
            st.session_state.submit_success_msg = ""

        st.info(f"📌 填表人：**{user['name']}** ｜ ⏱️ 系統將自動計算您的當日上班時數與包貨效率（若跨日下班亦會自動辨識換算）！")
        fv = st.session_state.form_ver
        
        st.markdown("### 📌 步驟一：請先勾選今日有執行的作業項目")
        chk_c1, chk_c2, chk_c3, chk_c4, chk_c5, chk_c6 = st.columns(6)
        has_pack = chk_c1.checkbox("📦 今日有包貨", key=f"hp_{fv}")
        has_ship = chk_c2.checkbox("🚚 今日有寄貨", key=f"hs_{fv}")
        has_cs = chk_c3.checkbox("💬 今日有客服", key=f"hcs_{fv}")
        has_live = chk_c4.checkbox("🎥 今日有直播", key=f"hlv_{fv}")
        has_complaint_chk = chk_c5.checkbox("🚨 今日有客訴", key=f"hc_{fv}")
        has_8d = chk_c6.checkbox("📄 需附 8D 檔案", key=f"h8d_{fv}")

        st.divider()
        st.markdown("### 📝 步驟二：填寫今日工作內容")
        st.markdown("#### 一、 基礎出勤與作業數據")
        report_date = st.date_input("📅 回報歸屬日期 (預設為台灣今日)", get_tw_today(), key=f"g_date_{fv}")
        col1, col2 = st.columns(2)
        work_start = col1.time_input("🕒 上班時間", datetime.time(9, 0), step=60, key=f"g_ws_{fv}")
        work_end = col2.time_input("🕕 下班時間", datetime.time(18, 0), step=60, key=f"g_we_{fv}")
            
        pack_start, pack_end, pack_count = datetime.time(10, 0), datetime.time(12, 0), 0
        if has_pack:
            st.markdown("##### 📦 包貨紀錄")
            c_p1, c_p2, c_p3 = st.columns(3)
            pack_start = c_p1.time_input("包貨開始時間", datetime.time(10, 0), step=60, key=f"g_ps_{fv}")
            pack_end = c_p2.time_input("包貨結束時間", datetime.time(12, 0), step=60, key=f"g_pe_{fv}")
            pack_count = c_p3.number_input("個人包貨件數", min_value=1, value=1, step=1, key=f"g_pc_{fv}")

        ship_count = 0
        if has_ship:
            st.markdown("##### 🚚 寄貨紀錄")
            ship_count = st.number_input("今日寄貨件數", min_value=1, value=1, step=1, key=f"g_sc_{fv}")
            
        cs_start, cs_end = datetime.time(9, 0), datetime.time(18, 0)
        if has_cs:
            st.markdown("##### 💬 客服回覆紀錄")
            c_l1, c_l2 = st.columns(2)
            cs_start = c_l1.time_input("客服回覆開始", datetime.time(9, 0), step=60, key=f"g_css_{fv}")
            cs_end = c_l2.time_input("客服回覆結束", datetime.time(18, 0), step=60, key=f"g_cse_{fv}")

        live_start, live_end = datetime.time(19, 0), datetime.time(22, 0)
        if has_live:
            st.markdown("##### 🎥 直播紀錄")
            c_l3, c_l4 = st.columns(2)
            live_start = c_l3.time_input("今日直播開始", datetime.time(19, 0), step=60, key=f"g_lvs_{fv}")
            live_end = c_l4.time_input("今日直播結束", datetime.time(22, 0), step=60, key=f"g_lve_{fv}")

        c_who = c_where = c_when = c_what = c_why = c_how = ""
        if has_complaint_chk:
            st.divider()
            st.markdown("#### 二、 🚨 黃金部門 - 客訴 5W1H 詳細回報")
            c_who = st.text_input("1. 【Who】客人本名 + TK名 / 帳號名稱", key=f"g_cwho_{fv}")
            c_where = st.selectbox("2. 【Where】客訴發生管道", ["直播上", "LINE客服上", "LINE社群", "其他"], key=f"g_cwhere_{fv}")
            c_when = st.text_input("3. 【When】幾月幾號的訂單 及 客訴發生日期", key=f"g_cwhen_{fv}")
            c_what = st.text_area("4. 【What】發生什麼事 / 商品的問題 (請詳細打字說明)", key=f"g_cwhat_{fv}")
            c_why = st.text_area("5. 【Why】客訴背後原因分析", key=f"g_cwhy_{fv}")
            c_how = st.text_area("6. 【How】執行方法 / 解決方案 / 未來如何改善", key=f"g_chow_{fv}")

        uploaded_8d = None
        if has_8d:
            st.divider()
            st.markdown("#### 📄 重大異常 8D 報告檔案上傳")
            uploaded_8d = st.file_uploader("請上傳填妥之 8D 檔案 (支援 Excel, Word, PDF)", type=["xlsx", "xls", "docx", "doc", "pdf"], key=f"g_8d_{fv}")

        st.divider()
        st.markdown("#### 三、 📝 今日處理事項明細")
        st.caption("請填寫幾點幾分到幾點幾分做了哪些事（空白未填的列會自動忽略，正上方與正下方皆可隨時增減列數）")

        gt_b1, gt_b2, _ = st.columns([1.5, 1.5, 4])
        if gt_b1.button(f"➕ 新增一列處理事項 (目前 {st.session_state.task_rows} 列)", key="g_add_top", use_container_width=True):
            st.session_state.task_rows += 1; st.rerun()
        if gt_b2.button("➖ 減少最後一列", key="g_sub_top", disabled=(st.session_state.task_rows <= 1), use_container_width=True):
            st.session_state.task_rows = max(1, st.session_state.task_rows - 1); st.rerun()
        
        raw_tasks_inputs = []
        for i in range(st.session_state.task_rows):
            t_col1, t_col2, t_col3 = st.columns([1.5, 1.5, 5])
            ts = t_col1.time_input(f"開始時間 #{i+1}", datetime.time(min(9 + i, 22), 0), key=f"ts_{fv}_{i}", step=60)
            te = t_col2.time_input(f"結束時間 #{i+1}", datetime.time(min(10 + i, 23), 0), key=f"te_{fv}_{i}", step=60)
            td = t_col3.text_input(f"處理事項內容 #{i+1}", placeholder="請輸入該時段具體完成的事項（無則留空即可）...", key=f"td_{fv}_{i}")
            raw_tasks_inputs.append((ts, te, td))

        gb_b1, gb_b2, _ = st.columns([1.5, 1.5, 4])
        if gb_b1.button("➕ 在下方繼續新增一列", key="g_add_bot", use_container_width=True):
            st.session_state.task_rows += 1; st.rerun()
        if gb_b2.button("➖ 移除最後一列", key="g_sub_bot", disabled=(st.session_state.task_rows <= 1), use_container_width=True):
            st.session_state.task_rows = max(1, st.session_state.task_rows - 1); st.rerun()

        st.divider()
        st.markdown("#### 四、 提案改善與同仁想法回饋")
        improvement_proposal = st.text_area("💡 提案改善欄位", key=f"g_imp_{fv}")
        peer_feedback = st.text_area("🤝 針對其他同仁有任何想法或建議 (僅主管後台可見)", key=f"g_peer_{fv}")

        if st.button("🚀 確認無誤，送出黃金部門今日回報", type="primary", use_container_width=True, key="g_sub_main"):
            tasks_list = [f"【{ts.strftime('%H:%M')} - {te.strftime('%H:%M')}】{td.strip()}" for ts, te, td in raw_tasks_inputs if td.strip()]
            has_complaint = "有" if has_complaint_chk else "無"
            file_8d_name, file_8d_b64 = "", ""
            if has_8d and uploaded_8d is not None:
                file_8d_name = uploaded_8d.name
                file_8d_b64 = base64.b64encode(uploaded_8d.read()).decode("utf-8")

            if has_complaint == "有" and not all([c_who.strip(), c_when.strip(), c_what.strip(), c_why.strip(), c_how.strip()]):
                st.error("❌ 您勾選了「今日有客訴」，請將 5W1H 所有欄位填寫完整！")
            elif has_8d and not file_8d_name:
                st.error("❌ 您勾選了「需附 8D 檔案」，請務必選擇檔案上傳！")
            elif not tasks_list:
                st.error("❌ 請至少填寫一條「今日處理事項」！")
            else:
                with st.spinner("🚀 正在安全傳送您的回報資料..."):
                    work_duration = calc_minutes(work_start, work_end)
                    pack_duration = calc_minutes(pack_start, pack_end) if has_pack else 0
                    min_per_pack = round(pack_duration / pack_count, 2) if (has_pack and pack_count > 0) else 0
                    cs_duration = calc_minutes(cs_start, cs_end) if has_cs else 0
                    live_duration = calc_minutes(live_start, live_end) if has_live else 0

                    clean_rdate = normalize_date_str(report_date)
                    was_overwritten = delete_existing_same_day_report(ws_reports, user["username"], clean_rdate)
                    now_tw = get_tw_now()
                    report_id = now_tw.strftime("%Y%m%d%H%M%S") + "_" + user["username"]
                    
                    new_data_dict = {
                        "回報單號": report_id, "送出時間": now_tw.strftime("%Y-%m-%d %H:%M:%S"),
                        "回報日期": clean_rdate, "帳號": user["username"], "姓名": user["name"],
                        "上班時間": work_start.strftime("%H:%M"), "下班時間": work_end.strftime("%H:%M"),
                        "上班總分鐘": work_duration,
                        "包貨開始": pack_start.strftime("%H:%M") if has_pack else "無",
                        "包貨結束": pack_end.strftime("%H:%M") if has_pack else "無",
                        "包貨件數": pack_count if has_pack else 0, "包貨總分鐘": pack_duration, "單件包貨耗時(分)": min_per_pack,
                        "客服開始": cs_start.strftime("%H:%M") if has_cs else "無",
                        "客服結束": cs_end.strftime("%H:%M") if has_cs else "無", "客服總分鐘": cs_duration,
                        "寄貨件數": ship_count if has_ship else 0,
                        "直播開始": live_start.strftime("%H:%M") if has_live else "無",
                        "直播結束": live_end.strftime("%H:%M") if has_live else "無", "直播總分鐘": live_duration,
                        "有無客訴": has_complaint, "客訴_Who": c_who, "客訴_Where": c_where, "客訴_When": c_when,
                        "客訴_What": c_what, "客訴_Why": c_why, "客訴_How": c_how,
                        "有無8D": "是" if has_8d else "否", "8D檔名": file_8d_name, "8D檔案編碼": file_8d_b64,
                        "今日處理事項": "\n".join(tasks_list), "提案改善": improvement_proposal, "對同仁想法": peer_feedback
                    }
                    write_dict_to_sheet(ws_reports, REPORT_HEADERS, new_data_dict)
                    sync_pretty_sheets(ws_reports, ws_compact, ws_comp)
                    clear_all_caches()
                    st.session_state.task_rows = 3
                    st.session_state.form_ver += 1
                    w_hr_disp = round(work_duration / 60.0, 2)
                    st.session_state.submit_success_msg = f"{'✅ 【覆蓋成功】' if was_overwritten else '🎉 【送出成功】'}您於 {clean_rdate} 的黃金部門回報已完成（當日上班時數：{w_hr_disp} 小時）！"
                    st.rerun()

        st.divider()
        with st.expander("🗂️ 檢視 / 撤回我近期已送出的黃金部門回報"):
            df_check = load_reports_df()
            if not df_check.empty:
                my_df = df_check[df_check["帳號"] == user["username"]].sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).head(10)
                for _, mrow in my_df.iterrows():
                    m_id = mrow["回報單號"]
                    mc1, mc2 = st.columns([4, 2])
                    mc1.markdown(f"**📅 日期：{mrow['回報日期']}** ｜ 上下班：`{mrow['上班時間']}~{mrow['下班時間']}` ｜ 包貨：{mrow['包貨件數']}件 / 寄貨：{mrow['寄貨件數']}件")
                    with mc2:
                        if st.session_state.confirm_del_id == f"self_{m_id}":
                            sc1, sc2 = st.columns(2)
                            if sc1.button("✅ 確定刪除", key=f"yes_self_{m_id}", type="primary", use_container_width=True):
                                delete_report_by_id(ws_reports, ws_compact, ws_comp, m_id)
                                st.session_state.confirm_del_id = None; st.rerun()
                            if sc2.button("❌ 取消", key=f"no_self_{m_id}", use_container_width=True):
                                st.session_state.confirm_del_id = None; st.rerun()
                        else:
                            if st.button("🗑️ 撤回刪除", key=f"self_del_{m_id}", use_container_width=True):
                                st.session_state.confirm_del_id = f"self_{m_id}"; st.rerun()

    # ---------------- 功能 A2：操台部門 - 每日工作回報 (支援多段班/跨日凌晨 & 多直播台動態新增) ----------------
    elif choice == "🎛️ 操台部門 - 每日工作回報":
        st.title("🎛️ 操台部門 - 每日工作回報單")
        if st.session_state.submit_success_msg:
            st.balloons()
            st.success(st.session_state.submit_success_msg)
            st.session_state.submit_success_msg = ""

        st.info(f"📌 填表人：**{user['name']}** ｜ ⏱️ 支援**兩段班／三段班新增**與**跨日到隔天凌晨自動計時**，系統將自動加總當日與當月工時！")
        fv = st.session_state.form_ver
        
        st.markdown("### 📌 步驟一：請先勾選今日是否有客訴或操台外處理事項")
        ct_chk1, ct_chk2, ct_chk3 = st.columns(3)
        ct_has_other_chk = ct_chk1.checkbox("🛠️ 今日有操台外處理事項 (遠端無協助則免勾)", value=True, key=f"ct_ho_{fv}")
        ct_has_comp_chk = ct_chk2.checkbox("🚨 今日直播中有客訴", value=False, key=f"ct_hc_{fv}")
        ct_has_8d_chk = ct_chk3.checkbox("📄 重大異常需附 8D 報告檔案", value=False, key=f"ct_h8d_{fv}")

        st.divider()
        st.markdown("### 📝 步驟二：填寫今日操台工作內容")
        
        # 一、出勤時間與工作設備（支援多段班 & 跨日凌晨即時預覽）
        st.markdown("#### 一、 出勤時間與工作設備（支援多段班與跨日凌晨自動計時）")
        ct_col1, ct_col2 = st.columns(2)
        ct_date = ct_col1.date_input("📅 歸屬日期 * (若跨日到凌晨，請選上機當天的日期)", get_tw_today(), key=f"ct_date_{fv}")
        with ct_col2:
            ct_device = st.text_input("📱 工作機或賴名稱 *", placeholder="遠端人員無工作機填寫私人賴名稱", key=f"ct_dev_{fv}")
            st.caption("遠端人員無工作機填寫私人賴名稱")

        st.caption("⚠️ 上下班時間不可湊整，請務必與打卡紙上相同（遠端人員請務必與上下機時間相同；若跨日到凌晨如 `20:00 ~ 02:30` 直接選時間即可，系統會自動算跨日！）")

        shift_entries = []
        preview_total_min = 0.0
        default_shift_times = [
            (datetime.time(13, 0), datetime.time(22, 0)),
            (datetime.time(19, 0), datetime.time(23, 30)),
            (datetime.time(0, 0), datetime.time(3, 0))
        ]
        for sh_i in range(st.session_state.ct_shift_rows):
            def_ws, def_we = default_shift_times[min(sh_i, 2)]
            sh_c1, sh_c2, sh_c3 = st.columns([2.2, 2.2, 2.6])
            seg_lbl = f" (第 {sh_i+1} 段)" if st.session_state.ct_shift_rows > 1 else ""
            s_ws = sh_c1.time_input(f"🕒 上班時間{seg_lbl} *", def_ws, step=60, key=f"ct_ws_{fv}_{sh_i}")
            s_we = sh_c2.time_input(f"🕕 下班時間{seg_lbl} *", def_we, step=60, key=f"ct_we_{fv}_{sh_i}")
            seg_min = calc_minutes(s_ws, s_we)
            seg_hr = round(seg_min / 60.0, 2)
            is_cross = (s_we < s_ws)
            preview_total_min += seg_min
            with sh_c3:
                st.write("")
                if is_cross:
                    st.info(f"🌙 跨日到隔天凌晨｜本段：**{seg_hr} 小時**")
                else:
                    st.success(f"⏱️ 本段工時：**{seg_hr} 小時** ({int(seg_min)} 分)")
            shift_entries.append((s_ws, s_we, is_cross, seg_min))

        sh_b1, sh_b2, sh_b3 = st.columns([2.0, 1.5, 3.5])
        if sh_b1.button(f"➕ 新增一段上下班時間 (兩段班/三段班)", key="sh_add_btn", use_container_width=True):
            st.session_state.ct_shift_rows += 1; st.rerun()
        if sh_b2.button("➖ 移除最後一段時間", key="sh_sub_btn", disabled=(st.session_state.ct_shift_rows <= 1), use_container_width=True):
            st.session_state.ct_shift_rows = max(1, st.session_state.ct_shift_rows - 1); st.rerun()
        if st.session_state.ct_shift_rows > 1:
            sh_b3.markdown(f"👉 **今日共 {st.session_state.ct_shift_rows} 段班，合計上班總工時：`{round(preview_total_min/60.0, 2)}` 小時**")

        st.divider()
        # 二、動態多台直播操台數據與需求紀錄
        st.markdown("#### 二、 直播操台數據與需求紀錄（可依實際操台數量往下新增直播台）")
        st.caption("💡 只要輸入『直播台名稱』並按 Enter，右方『總留言量、總點讚、總分享』前方就會自動帶出該直播台名稱，您只需填寫數字（如：8553-523-325）即可！")

        st_top_b1, st_top_b2, _ = st.columns([1.8, 1.4, 3.8])
        if st_top_b1.button(f"➕ 新增一台直播台 (目前共 {st.session_state.ct_stream_rows} 台)", key="st_add_top", use_container_width=True):
            st.session_state.ct_stream_rows += 1; st.rerun()
        if st_top_b2.button("➖ 移除最後一台", key="st_sub_top", disabled=(st.session_state.ct_stream_rows <= 1), use_container_width=True):
            st.session_state.ct_stream_rows = max(1, st.session_state.ct_stream_rows - 1); st.rerun()

        stream_entries = []
        for s_idx in range(st.session_state.ct_stream_rows):
            with st.container(border=True):
                st.markdown(f"##### 🎬 第 {s_idx + 1} 台直播操台紀錄")
                sc1, sc2 = st.columns(2)
                s_name = sc1.text_input(f"🎥 直播台名稱 #{s_idx+1} *", placeholder="請輸入直播台名稱（例如：廣天）", key=f"ct_sn_{fv}_{s_idx}")
                s_url = sc2.text_input(f"🔗 直播台粉專網址 #{s_idx+1} *", placeholder="請貼上該直播台對應粉專網址", key=f"ct_fu_{fv}_{s_idx}")

                mc1, mc2 = st.columns(2)
                with mc1:
                    s_opt = st.text_input(f"⏱️ 操台時間 #{s_idx+1} *", placeholder="格式範例：21:06-23:33(2.5H)", key=f"ct_opt_{fv}_{s_idx}")
                    st.caption("格式範例：21:06-23:33(2.5H)")
                with mc2:
                    prefix_title = f"【{s_name.strip()}】" if s_name.strip() else f"#{s_idx+1}"
                    st.markdown(f"<div style='font-size:14px; margin-bottom:4px;'>📊 {prefix_title} 總留言量、總點讚、總分享 *</div>", unsafe_allow_html=True)
                    p_col1, p_col2 = st.columns([1.3, 3.2])
                    p_col1.text_input("前綴", value=f"{s_name.strip()} -" if s_name.strip() else "(先填台名) -", disabled=True, label_visibility="collapsed", key=f"ct_mp_{fv}_{s_idx}")
                    s_met_nums = p_col2.text_input("數據", placeholder="只需填數字，例如：8553-523-325", label_visibility="collapsed", key=f"ct_met_{fv}_{s_idx}")
                    st.caption(f"自動組合為：{s_name.strip() or '直播台'}-{s_met_nums.strip() or '留言-點讚-分享'}")

                s_req = st.text_area(f"📋 直播需求 #{s_idx+1} *", placeholder="請詳細說明該台直播需求，禁止只打加單關鍵字反饋互動...", height=80, key=f"ct_req_{fv}_{s_idx}")
                s_img = st.file_uploader(f"📸 附上【{s_name.strip() or f'第 {s_idx+1} 台'}】直播需求調整截圖 (選填)", type=["png", "jpg", "jpeg", "pdf"], key=f"ct_rimg_{fv}_{s_idx}")
                s_rmet = st.text_area(f"✅ 直播需求是否達成 #{s_idx+1} *", placeholder="請依上題所述需求詳細說明，如未達成請詳述原因...", height=70, key=f"ct_rmet_{fv}_{s_idx}")

                stream_entries.append({
                    "idx": s_idx + 1, "name": s_name.strip(), "url": s_url.strip(),
                    "opt": s_opt.strip(), "met_nums": s_met_nums.strip(),
                    "req": s_req.strip(), "img": s_img, "rmet": s_rmet.strip()
                })

        st_bot_b1, st_bot_b2, _ = st.columns([1.8, 1.4, 3.8])
        if st_bot_b1.button("➕ 在下方接著新增一台直播台", key="st_add_bot", use_container_width=True):
            st.session_state.ct_stream_rows += 1; st.rerun()
        if st_bot_b2.button("➖ 移除最後一台直播台", key="st_sub_bot", disabled=(st.session_state.ct_stream_rows <= 1), use_container_width=True):
            st.session_state.ct_stream_rows = max(1, st.session_state.ct_stream_rows - 1); st.rerun()

        # 三、客訴與 8D
        ct_c_who = ct_c_where = ct_c_when = ct_c_what = ct_c_why = ct_c_how = ""
        if ct_has_comp_chk:
            st.divider()
            st.markdown("#### 三、 🚨 直播中客訴 5W1H 詳細回報（有客訴必填）")
            cw1, cw2, cw3 = st.columns(3)
            ct_c_who = cw1.text_input("1. 【Who】操台人員姓名 *", value=user["name"], key=f"ct_cwho_{fv}")
            ct_c_where = cw2.text_input("2. 【Where】直播台名稱 *", value="、".join([e["name"] for e in stream_entries if e["name"]]), key=f"ct_cwhere_{fv}")
            ct_c_when = cw3.text_input("3. 【When】發生時間點 (幾點幾分在賴群組反映) *", placeholder="例如：21:35 在賴群組反映", key=f"ct_cwhen_{fv}")
            ct_c_what = st.text_area("4. 【What】客訴問題詳情（客人反應什麼問題／當下發生什麼事） *", height=75, key=f"ct_cwhat_{fv}")
            ct_c_why = st.text_area("5. 【Why】客訴背後原因分析（為何會發生此狀況） *", height=75, key=f"ct_cwhy_{fv}")
            ct_c_how = st.text_area("6. 【How】當下處理方式 ／ 是否已處理完成 ／ 最終處理結果與未來改善 *", height=75, key=f"ct_chow_{fv}")

        ct_uploaded_8d = None
        if ct_has_8d_chk:
            st.divider()
            st.markdown("#### 📄 重大異常 8D 報告檔案上傳")
            ct_uploaded_8d = st.file_uploader("請上傳填妥之 8D 檔案", type=["xlsx", "xls", "docx", "doc", "pdf"], key=f"ct_u8d_{fv}")

        # 四、操台外處理事項明細
        ct_raw_tasks_inputs = []
        if ct_has_other_chk:
            st.divider()
            st.markdown("#### 四、 🛠️ 操台外處理事項明細")
            st.caption("請依照上班時間軸清楚標註幾點幾分到幾點幾分做了哪些事（正上方與正下方皆可直接點擊增減列數）")

            ct_top_b1, ct_top_b2, _ = st.columns([1.6, 1.4, 4])
            if ct_top_b1.button(f"➕ 新增一列操台外事項 (目前 {st.session_state.ct_task_rows} 列)", key="ct_add_top", use_container_width=True):
                st.session_state.ct_task_rows += 1; st.rerun()
            if ct_top_b2.button("➖ 減少最後一列", key="ct_sub_top", disabled=(st.session_state.ct_task_rows <= 1), use_container_width=True):
                st.session_state.ct_task_rows = max(1, st.session_state.ct_task_rows - 1); st.rerun()

            for i in range(st.session_state.ct_task_rows):
                tc1, tc2, tc3 = st.columns([1.5, 1.5, 5])
                cts = tc1.time_input(f"開始時間 #{i+1}", datetime.time(min(13 + i, 22), 0), key=f"ct_ts_{fv}_{i}", step=60)
                cte = tc2.time_input(f"結束時間 #{i+1}", datetime.time(min(14 + i, 23), 0), key=f"ct_te_{fv}_{i}", step=60)
                ctd = tc3.text_input(f"操台外處理事項內容 #{i+1}", placeholder="請輸入該時段具體完成的操台外事項（無則留空即可）...", key=f"ct_td_{fv}_{i}")
                ct_raw_tasks_inputs.append((cts, cte, ctd))

            ct_bot_b1, ct_bot_b2, _ = st.columns([1.6, 1.4, 4])
            if ct_bot_b1.button("➕ 在下方繼續新增一列", key="ct_add_bot", use_container_width=True):
                st.session_state.ct_task_rows += 1; st.rerun()
            if ct_bot_b2.button("➖ 移除最後一列", key="ct_sub_bot", disabled=(st.session_state.ct_task_rows <= 1), use_container_width=True):
                st.session_state.ct_task_rows = max(1, st.session_state.ct_task_rows - 1); st.rerun()

        st.divider()
        st.markdown("#### 五、 💬 整體操台感受、主管提醒與提案改善")
        ct_feeling = st.text_area("💬 直播台整體操台感受 *", placeholder="請詳述整體節奏、客人加單狀況、客人互動狀況、直播客人熱絡度等等...", height=85, key=f"ct_feel_{fv}")
        ct_sup_remind = st.text_area("📌 主管提醒改善事項 *", placeholder="請填寫主管提醒需改善與注意之事項與執行狀況（若無請填無）...", height=85, key=f"ct_sup_{fv}")
        ct_proposal = st.text_area("💡 提案改善", placeholder="請提供可以調整的地方或是建議，也可以在這邊跟我們聊聊天說說話釋放壓力（只有主管看得到呦~~~ 如果無也可以打無）", height=85, key=f"ct_prop_{fv}")

        if st.button("🚀 確認無誤，送出操台部門今日回報", type="primary", use_container_width=True, key="ct_sub_main"):
            ct_tasks_list = [f"【{cts.strftime('%H:%M')} - {cte.strftime('%H:%M')}】{ctd.strip()}" for cts, cte, ctd in ct_raw_tasks_inputs if ctd.strip()] if ct_has_other_chk else ["無"]
            ct_8d_name, ct_8d_b64 = "", ""
            if ct_has_8d_chk and ct_uploaded_8d is not None:
                ct_8d_name = ct_uploaded_8d.name
                ct_8d_b64 = base64.b64encode(ct_uploaded_8d.read()).decode("utf-8")

            missing_stream_idx = next((e["idx"] for e in stream_entries if not all([e["name"], e["url"], e["opt"], e["met_nums"], e["req"], e["rmet"]])), None)

            if not all([ct_device.strip(), ct_feeling.strip(), ct_sup_remind.strip()]):
                st.error("❌ 「工作機或賴名稱」、「整體操台感受」、「主管提醒改善事項」皆為必填！（若無主管提醒改善事項，請填『無』）")
            elif missing_stream_idx is not None:
                st.error(f"❌ 「第 {missing_stream_idx} 台直播操台紀錄」尚有必填欄位未填寫完整！")
            elif ct_has_other_chk and not ct_tasks_list:
                st.error("❌ 您勾選了「今日有操台外處理事項」，請至少填寫一列！（若遠端無協助請取消上方勾選）")
            elif ct_has_comp_chk and not all([ct_c_who.strip(), ct_c_where.strip(), ct_c_when.strip(), ct_c_what.strip(), ct_c_why.strip(), ct_c_how.strip()]):
                st.error("❌ 您勾選了「今日直播中有客訴」，請將客訴 5W1H 共 6 個欄位填寫完整！")
            elif ct_has_8d_chk and not ct_8d_name:
                st.error("❌ 您勾選了「重大異常需附 8D 報告檔案」，請務必上傳 8D 檔案！")
            else:
                with st.spinner("🚀 正在安全傳送您的操台回報資料..."):
                    ws_strs = [s_ws.strftime("%H:%M") for s_ws, _, _, _ in shift_entries]
                    we_strs = [f"{s_we.strftime('%H:%M')}(跨日)" if is_cr else s_we.strftime("%H:%M") for _, s_we, is_cr, _ in shift_entries]
                    work_min = round(sum(seg_m for _, _, _, seg_m in shift_entries), 2)
                    work_hr = round(work_min / 60.0, 2)

                    is_multi = len(stream_entries) > 1
                    names_joined = "、".join([e["name"] for e in stream_entries])
                    urls_list, opts_list, mets_list, reqs_list, rmets_list = [], [], [], [], []
                    img_names_list, img_b64_list = [], []
                    total_op_hr = 0.0

                    for e in stream_entries:
                        sn = e["name"]
                        total_op_hr += parse_single_op_hours(e["opt"])
                        clean_nums = e["met_nums"]
                        full_met = clean_nums if (clean_nums.startswith(f"{sn}-") or clean_nums.startswith(f"{sn} -")) else f"{sn}-{clean_nums.lstrip('- ')}"
                        mets_list.append(full_met)
                        urls_list.append(f"【{sn}】{e['url']}" if is_multi else e["url"])
                        opts_list.append(f"【{sn}】{e['opt']}" if is_multi else e["opt"])
                        reqs_list.append(f"【{sn}】{e['req']}" if is_multi else e["req"])
                        rmets_list.append(f"【{sn}】{e['rmet']}" if is_multi else e["rmet"])
                        if e["img"] is not None:
                            img_names_list.append(f"{sn}_{e['img'].name}")
                            img_b64_list.append({"name": f"{sn}_{e['img'].name}", "b64": base64.b64encode(e["img"].read()).decode("utf-8")})

                    op_hr = round(total_op_hr, 2)
                    if len(img_b64_list) == 0: img_name, img_b64 = "", ""
                    elif len(img_b64_list) == 1: img_name, img_b64 = img_b64_list[0]["name"], img_b64_list[0]["b64"]
                    else: img_name, img_b64 = " / ".join(img_names_list), "MULTI_JSON:" + json.dumps(img_b64_list, ensure_ascii=False)

                    clean_cdate = normalize_date_str(ct_date)
                    was_overwritten = delete_existing_same_day_report(ws_caotai, user["username"], clean_cdate)
                    now_tw = get_tw_now()
                    report_id = "CT_" + now_tw.strftime("%Y%m%d%H%M%S") + "_" + user["username"]
                    summary_comp_str = f"【Who】{ct_c_who}｜【Where】{ct_c_where}｜【When】{ct_c_when}｜【What】{ct_c_what}｜【Why】{ct_c_why}｜【How】{ct_c_how}" if ct_has_comp_chk else "無"

                    caotai_dict = {
                        "回報單號": report_id, "送出時間": now_tw.strftime("%Y-%m-%d %H:%M:%S"),
                        "回報日期": clean_cdate, "帳號": user["username"], "姓名": user["name"],
                        "工作機或賴名稱": ct_device.strip(),
                        "上班時間": "\n".join(ws_strs), "下班時間": "\n".join(we_strs),
                        "上班總分鐘": work_min, "上班總時數(小時)": work_hr,
                        "直播台名稱": names_joined, "直播台粉專網址": "\n".join(urls_list),
                        "操台時間": "\n".join(opts_list), "操台時數(小時)": op_hr,
                        "總留言_點讚_分享": "\n".join(mets_list), "直播需求": "\n".join(reqs_list),
                        "直播需求是否達成": "\n".join(rmets_list),
                        "有無客訴": "有" if ct_has_comp_chk else "無",
                        "客訴_Who": ct_c_who.strip() if ct_has_comp_chk else "",
                        "客訴_Where": ct_c_where.strip() if ct_has_comp_chk else "",
                        "客訴_When": ct_c_when.strip() if ct_has_comp_chk else "",
                        "客訴_What": ct_c_what.strip() if ct_has_comp_chk else "",
                        "客訴_Why": ct_c_why.strip() if ct_has_comp_chk else "",
                        "客訴_How": ct_c_how.strip() if ct_has_comp_chk else "",
                        "直播中客訴": summary_comp_str, "操台外處理事項": "\n".join(ct_tasks_list),
                        "直播台整體操台感受": ct_feeling.strip(), "主管提醒改善事項": ct_sup_remind.strip(),
                        "提案改善": ct_proposal.strip() or "無",
                        "需求截圖檔名": img_name, "需求截圖編碼": img_b64,
                        "有無8D": "是" if ct_has_8d_chk else "否", "8D檔名": ct_8d_name, "8D檔案編碼": ct_8d_b64
                    }

                    write_dict_to_sheet(ws_caotai, CAOTAI_HEADERS, caotai_dict)
                    sync_caotai_pretty_sheet(ws_caotai, ws_caotai_compact, ws_caotai_comp)
                    clear_all_caches()

                    st.session_state.ct_shift_rows = 1
                    st.session_state.ct_task_rows = 3
                    st.session_state.ct_stream_rows = 1
                    st.session_state.form_ver += 1
                    st.session_state.submit_success_msg = f"{'✅ 【覆蓋成功】' if was_overwritten else '🎉 【送出成功】'}您於 {clean_cdate} 的操台回報已完成（共 {len(shift_entries)} 段班，總上班時數：{work_hr} 小時 / 操台時數：{op_hr} 小時）！"
                    st.rerun()

        st.divider()
        with st.expander("🗂️ 檢視 / 撤回我近期已送出的操台部門回報"):
            df_ct_check = load_caotai_df()
            if not df_ct_check.empty:
                my_ct_df = df_ct_check[df_ct_check["帳號"] == user["username"]].sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).head(10)
                for _, mrow in my_ct_df.iterrows():
                    m_id = mrow["回報單號"]
                    mc1, mc2 = st.columns([4, 2])
                    sh_txt = format_shifts_display(mrow.get("上班時間", ""), mrow.get("下班時間", ""))
                    mc1.markdown(f"**📅 日期：{mrow['回報日期']}** ｜ 上下班：`{sh_txt}` ({mrow.get('上班總時數(小時)', '')}H) ｜ 直播台：{mrow['直播台名稱']}")
                    with mc2:
                        if st.session_state.confirm_del_id == f"self_ct_{m_id}":
                            sc1, sc2 = st.columns(2)
                            if sc1.button("✅ 確定刪除", key=f"yes_sct_{m_id}", type="primary", use_container_width=True):
                                delete_caotai_by_id(ws_caotai, ws_caotai_compact, ws_caotai_comp, m_id)
                                st.session_state.confirm_del_id = None; st.rerun()
                            if sc2.button("❌ 取消", key=f"no_sct_{m_id}", use_container_width=True):
                                st.session_state.confirm_del_id = None; st.rerun()
                        else:
                            if st.button("🗑️ 撤回刪除", key=f"self_ct_del_{m_id}", use_container_width=True):
                                st.session_state.confirm_del_id = f"self_ct_{m_id}"; st.rerun()

    # ---------------- 功能 B：修改密碼 ----------------
    elif choice == "🔒 修改密碼":
        st.title("🔒 修改個人密碼")
        with st.form("change_pw_form"):
            old_p = st.text_input("請輸入舊密碼 (或信箱收到的臨時密碼)", type="password")
            new_p1 = st.text_input("請輸入新密碼", type="password")
            new_p2 = st.text_input("請再次確認新密碼", type="password")
            pw_submit = st.form_submit_button("更新密碼", type="primary")
        if pw_submit:
            if new_p1 != new_p2:
                st.error("❌ 兩次新密碼輸入不一致！")
            elif hash_password(old_p) != user["password_hash"]:
                st.error("❌ 舊密碼輸入錯誤！")
            else:
                row_idx = find_user_row_idx(ws_users, user["username"])
                if row_idx:
                    new_hash = hash_password(new_p1)
                    ws_users.update_cell(row_idx, 2, new_hash)
                    clear_all_caches()
                    st.session_state.user_info["password_hash"] = new_hash
                    st.query_params["t"] = make_auth_token(user["username"], new_hash)
                    st.success("✅ 密碼更新成功！")

    # ---------------- 功能 C：帳號與權限管理 ----------------
    elif choice == "👥 帳號與權限管理" and user["role"] == "admin":
        st.title("👥 雙部門帳號、部門歸屬與權限管理中心")
        users_df = load_users_df()
        if not users_df.empty:
            display_users = users_df.copy()
            display_users["目前權限"] = display_users["系統權限"].apply(lambda r: "👑 管理員 (admin)" if r == "admin" else "👤 一般同仁 (staff)")
            st.dataframe(display_users[["登入帳號", "真實姓名", "所屬部門", "綁定信箱", "目前權限", "建立時間"]], use_container_width=True)

        st.divider()
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.subheader("🛡️ 調整現有帳號部門 / 權限 / 密碼")
            if not users_df.empty:
                user_options = [f"{r['真實姓名']} (部門: {r.get('所屬部門','黃金部門')} | 帳號: {r['登入帳號']} | 權限: {r['系統權限']})" for _, r in users_df.iterrows()]
                selected_opt = st.selectbox("請選擇要操作的帳號", user_options)
                target_row = users_df.iloc[user_options.index(selected_opt)]
                target_uname = target_row["登入帳號"]
                
                dept_list = ["黃金部門", "操台部門", "跨部門(兩部門皆可填寫)"]
                curr_dept = target_row.get("所屬部門", "黃金部門")
                new_dept = st.selectbox("調整所屬部門：", dept_list, index=dept_list.index(curr_dept) if curr_dept in dept_list else 0)
                new_role = st.radio("權限等級：", ["staff", "admin"], index=0 if target_row["系統權限"] == "staff" else 1,
                                    format_func=lambda x: "👤 一般同仁" if x == "staff" else "👑 管理員", horizontal=True)
                if st.button("💾 儲存部門與權限變更", type="primary"):
                    if target_uname == "admin" and new_role != "admin":
                        st.error("❌ 預設最高管理員 `admin` 不可降級！")
                    else:
                        sheet_row = find_user_row_idx(ws_users, target_uname)
                        if sheet_row:
                            ws_users.update_cell(sheet_row, 5, new_role)
                            ws_users.update_cell(sheet_row, 7, new_dept)
                            clear_all_caches()
                            if target_uname == user["username"]:
                                st.session_state.user_info["role"] = new_role
                                st.session_state.user_info["dept"] = new_dept
                            st.success("✅ 更新成功！"); st.rerun()
                            
                st.markdown("---")
                force_pass = st.text_input("強制設定新密碼 (免信箱驗證)", placeholder="例如：gold8888", key="f_pass")
                if st.button("🔑 直接更新該帳號密碼") and force_pass.strip():
                    sheet_row = find_user_row_idx(ws_users, target_uname)
                    if sheet_row:
                        ws_users.update_cell(sheet_row, 2, hash_password(force_pass.strip()))
                        clear_all_caches()
                        st.success(f"✅ 已將密碼重設為：`{force_pass.strip()}`")

                st.markdown("---")
                confirm_del = st.checkbox(f"確認永久刪除「{target_row['真實姓名']} ({target_uname})」")
                if st.button("🗑️ 刪除此帳號", disabled=not confirm_del):
                    if target_uname in ["admin", user["username"]]:
                        st.error("❌ 無法刪除 `admin` 或目前登入中的帳號！")
                    else:
                        sheet_row = find_user_row_idx(ws_users, target_uname)
                        if sheet_row:
                            ws_users.delete_rows(sheet_row)
                            clear_all_caches(); st.rerun()

        with col_m2:
            st.subheader("➕ 由後台直接幫同仁建立帳號")
            adm_dept = st.selectbox("選擇所屬部門", ["黃金部門", "操台部門", "跨部門(兩部門皆可填寫)"], key="a_dept")
            adm_name = st.text_input("同仁真實姓名", key="a_name")
            adm_email = st.text_input("同仁常用信箱", key="a_email")
            adm_user = st.text_input("設定登入帳號", key="a_user")
            adm_pass = st.text_input("設定預設密碼", value="gold8888", key="a_pass")
            adm_role = st.selectbox("初始權限", ["staff", "admin"], format_func=lambda x: "👤 一般同仁" if x == "staff" else "👑 管理員")
            if st.button("➕ 立即建立帳號", type="primary"):
                if not all([adm_name.strip(), adm_email.strip(), adm_user.strip(), adm_pass.strip()]):
                    st.warning("⚠️ 請填寫所有欄位！")
                elif not users_df.empty and adm_user.strip() in users_df["登入帳號"].values:
                    st.error("❌ 此登入帳號已經存在！")
                else:
                    ws_users.append_row([adm_user.strip(), hash_password(adm_pass.strip()), adm_name.strip(), adm_email.strip(), adm_role, get_tw_now().strftime("%Y-%m-%d %H:%M:%S"), adm_dept])
                    clear_all_caches()
                    st.success(f"✅ 已建立 **{adm_name}** 帳號！"); st.rerun()

    # ---------------- 功能 D：主管後台數據中心 ----------------
    elif choice == "📊 主管後台數據中心" and user["role"] == "admin":
        top_c1, top_c2 = st.columns([3.5, 1.2])
        top_c1.title("📊 主管後台數據中心")
        with top_c2:
            st.write("")
            if st.button("🔄 立即刷新最新數據與圖表", use_container_width=True):
                clear_all_caches(); st.rerun()

        dept_tab1, dept_tab2, dept_tab3 = st.tabs([
            "✨ 【黃金部門】數據與回報管理",
            "🎛️ 【操台部門】工時統計與回報管理",
            "🚨 【季稽核專區】跨部門客訴區間查詢與匯出"
        ])

        # 分頁 1：黃金部門後台
        with dept_tab1:
            df = load_reports_df()
            if df.empty:
                st.info("目前黃金部門尚無任何回報資料。")
            else:
                for col in ["包貨件數", "包貨總分鐘", "單件包貨耗時(分)", "寄貨件數", "客服總分鐘", "直播總分鐘"]:
                    df[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0)
                df["上班總分鐘"] = df.apply(lambda r: calc_minutes_str(r.get("上班時間", "09:00"), r.get("下班時間", "18:00")), axis=1)
                df["上班時數(小時)"] = (df["上班總分鐘"] / 60.0).round(2)
                df["包貨總分鐘"] = df.apply(lambda r: calc_minutes_str(r.get("包貨開始", "無"), r.get("包貨結束", "無")) or r["包貨總分鐘"], axis=1)
                df["單件包貨耗時(分)"] = df.apply(lambda r: round(r["包貨總分鐘"] / r["包貨件數"], 2) if r["包貨件數"] > 0 else 0, axis=1)
                df["客服總分鐘"] = df.apply(lambda r: calc_minutes_str(r.get("客服開始", "無"), r.get("客服結束", "無")) or r["客服總分鐘"], axis=1)
                df["直播總分鐘"] = df.apply(lambda r: calc_minutes_str(r.get("直播開始", "無"), r.get("直播結束", "無")) or r["直播總分鐘"], axis=1)
                df["月份"] = df["回報日期"].astype(str).str.slice(0, 7)
                df["季度"] = df["回報日期"].apply(get_quarter_label)
                
                col_f1, col_f2, col_f3 = st.columns(3)
                filter_mode = col_f1.radio("統計週期維度", ["📅 依月份檢視", "📆 依季度檢視 (Q1~Q4)"], horizontal=True, key="g_fmode")
                if filter_mode == "📅 依月份檢視":
                    sel_period = col_f2.selectbox("選擇統計月份", ["全部月份"] + sorted(df["月份"].unique().tolist(), reverse=True), key="g_fper")
                else:
                    sel_period = col_f2.selectbox("選擇統計季度", ["全部季度"] + sorted(df["季度"].unique().tolist(), reverse=True), key="g_fqtr")
                sel_name = col_f3.selectbox("👤 選擇統計同仁", ["全部同仁"] + sorted(df["姓名"].unique().tolist()), key="g_fname")
                    
                filtered_df = df.copy()
                if filter_mode == "📅 依月份檢視" and sel_period != "全部月份":
                    filtered_df = filtered_df[filtered_df["月份"] == sel_period]
                elif filter_mode != "📅 依月份檢視" and sel_period != "全部季度":
                    filtered_df = filtered_df[filtered_df["季度"] == sel_period]
                if sel_name != "全部同仁":
                    filtered_df = filtered_df[filtered_df["姓名"] == sel_name]

                total_packs = int(filtered_df["包貨件數"].sum())
                total_pack_min = float(filtered_df["包貨總分鐘"].sum())
                k0, k1, k2, k3, k4, k5 = st.columns(6)
                k0.metric("🕒 區間上班總時數", f"{round(filtered_df['上班總分鐘'].sum() / 60.0, 1)} 小時")
                k1.metric("📦 區間總包貨件數", f"{total_packs} 件")
                k2.metric("⚡ 平均單件包貨耗時", f"{round(total_pack_min / total_packs, 2) if total_packs > 0 else 0} 分/件")
                k3.metric("🚚 區間總寄貨件數", f"{int(filtered_df['寄貨件數'].sum())} 件")
                k4.metric("💬 客服累計總時數", f"{round(filtered_df['客服總分鐘'].sum() / 60.0, 1)} 小時")
                k5.metric("🎥 直播累計總時數", f"{round(filtered_df['直播總分鐘'].sum() / 60.0, 1)} 小時")

                st.divider()
                if not filtered_df.empty:
                    summary_df = filtered_df.groupby("姓名").agg(
                        出勤回報天數=("回報單號", "count"),
                        當月上班總時數_小時=("上班總分鐘", lambda x: round(x.sum() / 60.0, 2)),
                        總包貨件數=("包貨件數", "sum"),
                        總包貨耗時_分=("包貨總分鐘", "sum"),
                        總寄貨件數=("寄貨件數", "sum"),
                        客服總時數_小時=("客服總分鐘", lambda x: round(x.sum() / 60.0, 2)),
                        直播總時數_小時=("直播總分鐘", lambda x: round(x.sum() / 60.0, 2)),
                        客訴發生次數=("有無客訴", lambda x: (x == "有").sum())
                    ).reset_index()
                    summary_df["平均每件包貨耗時(分/件)"] = summary_df.apply(lambda r: round(r["總包貨耗時_分"] / r["總包貨件數"], 2) if r["總包貨件數"] > 0 else 0, axis=1)
                    summary_df = summary_df[["姓名", "出勤回報天數", "當月上班總時數_小時", "總包貨件數", "總包貨耗時_分", "平均每件包貨耗時(分/件)", "總寄貨件數", "客服總時數_小時", "直播總時數_小時", "客訴發生次數"]]
                    st.subheader(f"📈 黃金部門 - 各同仁統計總表 ({sel_period} - {sel_name})")
                    st.dataframe(summary_df, use_container_width=True)

                    st.divider()
                    st.subheader("📈 每日包貨效率折線圖（橫軸：天 ／ 縱軸：單件包貨耗時分鐘）")
                    daily_pack_df = filtered_df[filtered_df["包貨件數"] > 0].copy()
                    if not daily_pack_df.empty:
                        daily_grouped = daily_pack_df.groupby(["回報日期", "姓名"], as_index=False).agg(當日包貨件數=("包貨件數", "sum"), 當日包貨總分鐘=("包貨總分鐘", "sum"))
                        daily_grouped["單件包貨耗時(分/件)"] = (daily_grouped["當日包貨總分鐘"] / daily_grouped["當日包貨件數"]).round(2)
                        sorted_all_dates = sorted(daily_grouped["回報日期"].unique().tolist())
                        daily_grouped = daily_grouped.sort_values(["回報日期", "姓名"])

                        fig_compare = px.line(daily_grouped, x="回報日期", y="單件包貨耗時(分/件)", color="姓名", markers=True, text="單件包貨耗時(分/件)",
                                              hover_data=["當日包貨件數", "當日包貨總分鐘"], category_orders={"回報日期": sorted_all_dates}, title="👥 各同仁包貨效率比較折線圖")
                        fig_compare.update_xaxes(type="category", categoryorder="array", categoryarray=sorted_all_dates)
                        fig_compare.update_traces(textposition="top center", connectgaps=True, marker=dict(size=9))
                        st.plotly_chart(fig_compare, use_container_width=True)

                        staff_list = sorted(daily_grouped["姓名"].unique().tolist())
                        target_staff = st.selectbox("👤 單獨檢視折線圖的同仁：", ["📋 一次展開所有同仁各自的獨立折線圖"] + staff_list)
                        if target_staff == "📋 一次展開所有同仁各自的獨立折線圖":
                            for i in range(0, len(staff_list), 2):
                                c_list = st.columns(2)
                                for j in range(2):
                                    if i + j < len(staff_list):
                                        st_name = staff_list[i + j]
                                        single_df = daily_grouped[daily_grouped["姓名"] == st_name].sort_values("回報日期")
                                        s_dates = sorted(single_df["回報日期"].unique().tolist())
                                        fig_s = px.line(single_df, x="回報日期", y="單件包貨耗時(分/件)", markers=True, text="單件包貨耗時(分/件)", title=f"👤 {st_name} - 個人包貨效率")
                                        fig_s.update_xaxes(type="category", categoryorder="array", categoryarray=s_dates)
                                        fig_s.update_traces(line_color="#00B4D8", textposition="top center", marker=dict(size=9))
                                        c_list[j].plotly_chart(fig_s, use_container_width=True)
                        else:
                            single_df = daily_grouped[daily_grouped["姓名"] == target_staff].sort_values("回報日期")
                            s_dates = sorted(single_df["回報日期"].unique().tolist())
                            fig_s = px.line(single_df, x="回報日期", y="單件包貨耗時(分/件)", markers=True, text="單件包貨耗時(分/件)", title=f"👤 {target_staff} - 個人包貨效率")
                            fig_s.update_xaxes(type="category", categoryorder="array", categoryarray=s_dates)
                            fig_s.update_traces(line_color="#00B4D8", textposition="top center", marker=dict(size=9))
                            st.plotly_chart(fig_s, use_container_width=True)

                st.divider()
                col_dl1, col_dl2 = st.columns([2, 1])
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    if not filtered_df.empty: summary_df.to_excel(writer, index=False, sheet_name="1_人員工時與統計彙整")
                    filtered_df.drop(columns=["8D檔案編碼"], errors="ignore").to_excel(writer, index=False, sheet_name="2_每日作業明細")
                    style_excel_File(writer)
                safe_p = sel_period if "全部" not in sel_period else "ALL"
                col_dl1.download_button("📊 點擊下載【黃金部門】Excel 報表 (.xlsx)", data=output.getvalue(), file_name=f"Gold_Report_{safe_p}_{get_tw_today().strftime('%Y%m%d')}.xlsx", type="primary", use_container_width=True)
                if col_dl2.button("✨ 重新美化並同步 Google 試算表", key="sync_g_btn", use_container_width=True):
                    sync_pretty_sheets(ws_reports, ws_compact, ws_comp, apply_style=True); clear_all_caches(); st.success("✅ 已同步！")

                st.divider()
                st.subheader("📋 黃金部門 - 每日回報詳細內容檢視、修改與刪除管理")
                det_c1, det_c2, det_c3, det_c4 = st.columns(4)
                all_m_det = ["全部月份"] + sorted(df["月份"].unique().tolist(), reverse=True)
                det_month = det_c1.selectbox("📅 檢視月份", all_m_det, key="det_m_sel")
                det_name = det_c2.selectbox("👤 檢視同仁", ["全部同仁"] + sorted(df["姓名"].unique().tolist()), key="det_n_sel")
                detail_df = df.copy()
                if det_month != "全部月份": detail_df = detail_df[detail_df["月份"] == det_month]
                if det_name != "全部同仁": detail_df = detail_df[detail_df["姓名"] == det_name]
                det_date = det_c3.selectbox("📆 特定日期", ["全部日期"] + sorted(detail_df["回報日期"].unique().tolist(), reverse=True), key="det_d_sel")
                det_status = det_c4.selectbox("🚨 快篩", ["全部回報", "🚨 僅看有客訴", "📎 僅看有附8D報告"], key="det_s_sel")
                if det_date != "全部日期": detail_df = detail_df[detail_df["回報日期"] == det_date]
                if det_status == "🚨 僅看有客訴": detail_df = detail_df[detail_df["有無客訴"] == "有"]
                elif det_status == "📎 僅看有附8D報告": detail_df = detail_df[detail_df["有無8D"] == "是"]

                for idx, row in detail_df.sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).reset_index(drop=True).iterrows():
                    r_id = row.get("回報單號", f"idx_{idx}")
                    alert_b = " 🚨【有客訴】" if row.get("有無客訴") == "有" else ""
                    exp_t = f"📅 {row.get('回報日期', '')} ｜ 👤 {row.get('姓名', '')} ｜ 🕒 {row.get('上班時間', '')}~{row.get('下班時間', '')} ({row.get('上班時數(小時)', 0)}H) ｜ 📦 包貨：{int(row.get('包貨件數', 0))}件 ({row.get('單件包貨耗時(分)', 0)}分/件) / 🚚 寄貨：{int(row.get('寄貨件數', 0))}件{alert_b}"
                    is_ed = (st.session_state.edit_report_id == r_id)
                    with st.expander(exp_t, expanded=(idx == 0 or is_ed)):
                        c1, c2, c3, c4 = st.columns(4)
                        c1.markdown(f"**🕒 出勤時間**\n\n`{row.get('上班時間', '')} ~ {row.get('下班時間', '')}` (**{row.get('上班時數(小時)', 0)}** H)")
                        c2.markdown(f"**📦 包貨紀錄**\n\n`{row.get('包貨開始', '')} ~ {row.get('包貨結束', '')}` (**{int(row.get('包貨件數', 0))}** 件)" if int(row.get("包貨件數", 0)) > 0 else "**📦 包貨紀錄**\n\n`無`")
                        c3.markdown(f"**🚚 寄貨 / 💬 客服**\n\n寄貨：**{int(row.get('寄貨件數', 0))}** 件｜客服：**{round(float(row.get('客服總分鐘', 0))/60.0, 2)}** H")
                        c4.markdown(f"**🎥 直播紀錄**\n\n`{row.get('直播開始', '無')} ~ {row.get('直播結束', '無')}` (**{round(float(row.get('直播總分鐘', 0))/60.0, 2)}** H)")
                        st.markdown("---")
                        st.markdown("#### 📝 今日處理事項明細")
                        for line in str(row.get("今日處理事項", "")).split("\n"):
                            if line.strip(): st.markdown(f"- {line}")
                        col_p1, col_p2 = st.columns(2)
                        col_p1.info(f"💡 提案改善：{row.get('提案改善', '') or '無'}")
                        col_p2.warning(f"🤝 對同仁想法：{row.get('對同仁想法', '') or '無'}")
                        if row.get("有無客訴") == "有":
                            st.error(f"🚨 **客訴報告**：Who:{row.get('客訴_Who','')} | Where:{row.get('客訴_Where','')} | When:{row.get('客訴_When','')}\n- What: {row.get('客訴_What','')}\n- Why: {row.get('客訴_Why','')}\n- How: {row.get('客訴_How','')}")
                        if row.get("有無8D") == "是" and row.get("8D檔案編碼"):
                            st.download_button(f"📎 下載 8D 檔案 ({row.get('8D檔名', '8D.xlsx')})", data=base64.b64decode(row["8D檔案編碼"]), file_name=f"8D_{r_id}_{row.get('8D檔名', 'report.xlsx')}", key=f"dl_card_{r_id}")
                        
                        foot_c1, foot_c2, foot_c3 = st.columns([2.5, 1.2, 1.3])
                        foot_c1.caption(f"單號：{r_id} ｜ 送出時間：{row.get('送出時間', '')}")
                        if is_ed:
                            if foot_c2.button("❌ 收合修改", key=f"cl_ed_{r_id}", use_container_width=True): st.session_state.edit_report_id = None; st.rerun()
                        else:
                            if foot_c2.button("✏️ 修改內容", key=f"adm_ed_{r_id}", use_container_width=True): st.session_state.edit_report_id = r_id; st.rerun()
                        with foot_c3:
                            if st.session_state.confirm_del_id == f"adm_{r_id}":
                                dc1, dc2 = st.columns(2)
                                if dc1.button("✅ 確定", key=f"y_adm_{r_id}", type="primary"): delete_report_by_id(ws_reports, ws_compact, ws_comp, r_id); st.session_state.confirm_del_id = None; st.rerun()
                                if dc2.button("❌ 取消", key=f"n_adm_{r_id}"): st.session_state.confirm_del_id = None; st.rerun()
                            else:
                                if st.button("🗑️ 刪除此筆", key=f"adm_del_{r_id}", use_container_width=True): st.session_state.confirm_del_id = f"adm_{r_id}"; st.rerun()

                        if is_ed:
                            with st.form(key=f"edit_form_{r_id}"):
                                ec1, ec2, ec3, ec4 = st.columns(4)
                                e_date = ec1.date_input("回報日期", parse_date_safe(row.get("回報日期", "")))
                                e_name = ec2.text_input("姓名", value=str(row.get("姓名", "")))
                                e_ws = ec3.time_input("上班時間", parse_time_safe(row.get("上班時間", "09:00")), step=60)
                                e_we = ec4.time_input("下班時間", parse_time_safe(row.get("下班時間", "18:00"), datetime.time(18, 0)), step=60)
                                ep1, ep2, ep3, ep4 = st.columns(4)
                                e_ps = ep1.time_input("包貨開始", parse_time_safe(row.get("包貨開始", "10:00"), datetime.time(10, 0)), step=60)
                                e_pe = ep2.time_input("包貨結束", parse_time_safe(row.get("包貨結束", "12:00"), datetime.time(12, 0)), step=60)
                                e_pc = ep3.number_input("包貨件數", min_value=0, value=int(row.get("包貨件數", 0)), step=1)
                                e_sc = ep4.number_input("寄貨件數", min_value=0, value=int(row.get("寄貨件數", 0)), step=1)
                                e_tasks = st.text_area("今日處理事項明細", value=str(row.get("今日處理事項", "")), height=90)
                                e_prop = st.text_area("提案改善", value=str(row.get("提案改善", "")), height=60)
                                e_peer = st.text_area("對同仁想法", value=str(row.get("對同仁想法", "")), height=60)
                                if st.form_submit_button("💾 儲存修改", type="primary", use_container_width=True):
                                    new_w_dur = calc_minutes(e_ws, e_we)
                                    new_p_dur = calc_minutes(e_ps, e_pe) if e_pc > 0 else 0
                                    new_min_p = round(new_p_dur / e_pc, 2) if e_pc > 0 else 0
                                    upd = dict(row)
                                    upd.update({
                                        "回報日期": normalize_date_str(e_date), "姓名": e_name.strip(),
                                        "上班時間": e_ws.strftime("%H:%M"), "下班時間": e_we.strftime("%H:%M"), "上班總分鐘": new_w_dur,
                                        "包貨開始": e_ps.strftime("%H:%M") if e_pc > 0 else "無", "包貨結束": e_pe.strftime("%H:%M") if e_pc > 0 else "無",
                                        "包貨件數": e_pc, "包貨總分鐘": new_p_dur, "單件包貨耗時(分)": new_min_p, "寄貨件數": e_sc,
                                        "今日處理事項": e_tasks.strip(), "提案改善": e_prop.strip(), "對同仁想法": e_peer.strip()
                                    })
                                    write_dict_to_sheet(ws_reports, REPORT_HEADERS, upd, target_report_id=r_id)
                                    sync_pretty_sheets(ws_reports, ws_compact, ws_comp); clear_all_caches()
                                    st.session_state.edit_report_id = None; st.rerun()

        # 分頁 2：操台部門後台
        with dept_tab2:
            df_ct = load_caotai_df()
            if df_ct.empty:
                st.info("目前操台部門尚無任何回報資料。")
            else:
                if "主管提醒改善事項" not in df_ct.columns: df_ct["主管提醒改善事項"] = "無"
                df_ct["上班總分鐘"] = df_ct.apply(lambda r: calc_minutes_str(r.get("上班時間", ""), r.get("下班時間", "")), axis=1)
                df_ct["上班總時數(小時)"] = (df_ct["上班總分鐘"] / 60.0).round(2)
                df_ct["操台時數(小時)"] = df_ct.apply(lambda r: parse_caotai_hours(r.get("操台時間", "")), axis=1)
                df_ct["判定有客訴"] = df_ct.apply(is_caotai_row_complaint, axis=1)
                df_ct["月份"] = df_ct["回報日期"].astype(str).str.slice(0, 7)

                ct_f1, ct_f2 = st.columns(2)
                ct_months = ["全部月份"] + sorted(df_ct["月份"].unique().tolist(), reverse=True)
                sel_ct_month = ct_f1.selectbox("📅 選擇統計月份", ct_months, key="ct_m_top")
                ct_names = ["全部同仁"] + sorted(df_ct["姓名"].unique().tolist())
                sel_ct_name = ct_f2.selectbox("👤 選擇統計同仁", ct_names, key="ct_n_top")

                f_ct_df = df_ct.copy()
                if sel_ct_month != "全部月份": f_ct_df = f_ct_df[f_ct_df["月份"] == sel_ct_month]
                if sel_ct_name != "全部同仁": f_ct_df = f_ct_df[f_ct_df["姓名"] == sel_ct_name]

                ck1, ck2, ck3, ck4 = st.columns(4)
                ck1.metric("📅 區間總出勤天數", f"{len(f_ct_df)} 天")
                ck2.metric("🕒 區間累計上班總時數", f"{round(f_ct_df['上班總時數(小時)'].sum(), 2)} 小時")
                ck3.metric("🎛️ 區間累計操台總時數", f"{round(f_ct_df['操台時數(小時)'].sum(), 2)} 小時")
                ck4.metric("🚨 直播有客訴筆數", f"{int(f_ct_df['判定有客訴'].sum())} 筆")

                st.divider()
                if not f_ct_df.empty:
                    ct_summary_df = f_ct_df.groupby("姓名").agg(
                        出勤回報天數=("回報單號", "count"),
                        當月上班總時數_小時=("上班總時數(小時)", lambda x: round(x.sum(), 2)),
                        平均每日上班時數_小時=("上班總時數(小時)", lambda x: round(x.mean(), 2)),
                        當月操台總時數_小時=("操台時數(小時)", lambda x: round(x.sum(), 2)),
                        直播客訴發生次數=("判定有客訴", lambda x: int(x.sum()))
                    ).reset_index()
                    st.subheader(f"📈 操台部門 - 各同仁當月上班總時數統計表 ({sel_ct_month} - {sel_ct_name})")
                    st.dataframe(ct_summary_df, use_container_width=True)

                st.divider()
                ct_dl1, ct_dl2 = st.columns([2, 1])
                ct_export_df = pd.DataFrame({
                    "回報日期": f_ct_df["回報日期"], "姓名": f_ct_df["姓名"], "工作機或賴名稱": f_ct_df["工作機或賴名稱"],
                    "上下班各段時間": f_ct_df.apply(lambda r: format_shifts_display(r.get("上班時間", ""), r.get("下班時間", ""), sep="\n"), axis=1),
                    "當日上班總時數(小時)": f_ct_df["上班總時數(小時)"],
                    "直播台名稱": f_ct_df["直播台名稱"], "粉專網址": f_ct_df["直播台粉專網址"],
                    "操台時間": f_ct_df["操台時間"], "操台總時數(小時)": f_ct_df["操台時數(小時)"],
                    "總留言/點讚/分享": f_ct_df["總留言_點讚_分享"], "直播需求": f_ct_df["直播需求"],
                    "需求是否達成": f_ct_df["直播需求是否達成"], "有無客訴": f_ct_df["判定有客訴"].apply(lambda x: "有" if x else "無"),
                    "操台外處理事項": f_ct_df["操台外處理事項"], "整體操台感受": f_ct_df["直播台整體操台感受"],
                    "主管提醒改善事項": f_ct_df["主管提醒改善事項"], "提案改善": f_ct_df["提案改善"]
                })
                ct_output = io.BytesIO()
                with pd.ExcelWriter(ct_output, engine="openpyxl") as writer:
                    if not f_ct_df.empty: ct_summary_df.to_excel(writer, index=False, sheet_name="1_操台人員月度總工時統計")
                    ct_export_df.to_excel(writer, index=False, sheet_name="2_操台每日回報明細")
                    style_excel_File(writer)
                safe_ct_p = sel_ct_month if "全部" not in sel_ct_month else "ALL"
                ct_dl1.download_button("📊 點擊下載【操台部門】工時統計與明細 Excel (.xlsx)", data=ct_output.getvalue(), file_name=f"Caotai_Report_{safe_ct_p}_{get_tw_today().strftime('%Y%m%d')}.xlsx", type="primary", use_container_width=True)
                if ct_dl2.button("✨ 重新美化並同步操台 Google 試算表", key="sync_ct_btn", use_container_width=True):
                    sync_caotai_pretty_sheet(ws_caotai, ws_caotai_compact, ws_caotai_comp, apply_style=True); clear_all_caches(); st.success("✅ 已同步！")

                st.divider()
                st.subheader("📋 操台部門 - 每日回報詳細內容檢視、修改與刪除管理")
                cd1, cd2, cd3, cd4 = st.columns(4)
                ct_det_m = cd1.selectbox("📅 篩選月份", ct_months, index=ct_months.index(sel_ct_month), key="ct_dm")
                ct_det_n = cd2.selectbox("👤 篩選同仁", ct_names, index=ct_names.index(sel_ct_name), key="ct_dn")
                ct_detail_df = df_ct.copy()
                if ct_det_m != "全部月份": ct_detail_df = ct_detail_df[ct_detail_df["月份"] == ct_det_m]
                if ct_det_n != "全部同仁": ct_detail_df = ct_detail_df[ct_detail_df["姓名"] == ct_det_n]
                ct_det_d = cd3.selectbox("📆 篩選日期", ["全部日期"] + sorted(ct_detail_df["回報日期"].unique().tolist(), reverse=True), key="ct_dd")
                ct_det_c = cd4.selectbox("🚨 快篩", ["全部回報", "🚨 僅看有客訴", "📎 僅看有附8D報告"], key="ct_dc")
                if ct_det_d != "全部日期": ct_detail_df = ct_detail_df[ct_detail_df["回報日期"] == ct_det_d]
                if ct_det_c == "🚨 僅看有客訴": ct_detail_df = ct_detail_df[ct_detail_df["判定有客訴"]]
                elif ct_det_c == "📎 僅看有附8D報告": ct_detail_df = ct_detail_df[ct_detail_df.get("有無8D", "否") == "是"]

                for idx, row in ct_detail_df.sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).reset_index(drop=True).iterrows():
                    r_id = row.get("回報單號", f"ct_{idx}")
                    has_ct_comp = bool(row.get("判定有客訴", False))
                    sh_inline = format_shifts_display(row.get("上班時間", ""), row.get("下班時間", ""))
                    short_opt = str(row.get("操台時間", "")).replace("\n", " / ")
                    exp_title = f"📅 {row.get('回報日期', '')} ｜ 👤 {row.get('姓名', '')} ｜ 🕒 {sh_inline} (總工時：{row.get('上班總時數(小時)', 0)}H) ｜ 🎥 直播台：{row.get('直播台名稱', '')} ｜ ⏱️ 操台：{short_opt}{' 🚨【有客訴】' if has_ct_comp else ''}"

                    is_editing_ct = (st.session_state.edit_report_id == r_id)
                    with st.expander(exp_title, expanded=(idx == 0 or is_editing_ct)):
                        c1, c2, c3, c4 = st.columns(4)
                        sh_lines_md = "\n".join([f"- `{s}`" for s in format_shifts_display(row.get("上班時間", ""), row.get("下班時間", ""), sep="\n").split("\n")])
                        c1.markdown(f"**🕒 上下班時段 (共 {row.get('上班總時數(小時)', 0)} 小時)**\n\n{sh_lines_md}")
                        c2.markdown(f"**📱 工作機或賴名稱**\n\n`{row.get('工作機或賴名稱', '')}`")
                        urls_md = "\n".join([f"- {u}" for u in str(row.get("直播台粉專網址", "")).split("\n") if u.strip()])
                        c3.markdown(f"**🎥 直播台名稱 / 粉專**\n\n**{row.get('直播台名稱', '')}**\n{urls_md}")
                        opts_md = "\n".join([f"- `{o}`" for o in str(row.get("操台時間", "")).split("\n") if o.strip()])
                        mets_md = "\n".join([f"- `{m}`" for m in str(row.get("總留言_點讚_分享", "")).split("\n") if m.strip()])
                        c4.markdown(f"**⏱️ 操台 ({row.get('操台時數(小時)', 0)}H) / 互動**\n\n{opts_md}\n{mets_md}")

                        st.markdown("---")
                        r_c1, r_c2 = st.columns(2)
                        with r_c1:
                            st.markdown("#### 📋 直播需求")
                            for l in str(row.get("直播需求", "")).split("\n"):
                                if l.strip(): st.write(l)
                            raw_img_b64 = str(row.get("需求截圖編碼", ""))
                            if raw_img_b64:
                                if raw_img_b64.startswith("MULTI_JSON:"):
                                    try:
                                        for m_idx, item in enumerate(json.loads(raw_img_b64[len("MULTI_JSON:"):])):
                                            st.download_button(f"📸 下載截圖 ({item['name']})", data=base64.b64decode(item["b64"]), file_name=f"{r_id}_{item['name']}", key=f"dl_cti_{r_id}_{m_idx}")
                                    except: pass
                                else:
                                    st.download_button(f"📸 下載截圖 ({row.get('需求截圖檔名', 'img.png')})", data=base64.b64decode(raw_img_b64), file_name=f"{r_id}_{row.get('需求截圖檔名', 'img.png')}", key=f"dl_cti_{r_id}")
                        with r_c2:
                            st.markdown("#### ✅ 直播需求是否達成")
                            for l in str(row.get("直播需求是否達成", "")).split("\n"):
                                if l.strip(): st.write(l)

                        st.markdown("---")
                        st.markdown("#### 🛠️ 操台外處理事項")
                        ot_t = str(row.get("操台外處理事項", "")).strip()
                        if ot_t and ot_t != "無":
                            for l in ot_t.split("\n"): st.markdown(f"- {l}")
                        else: st.caption("無")

                        if has_ct_comp:
                            st.error(f"🚨 **操台客訴 5W1H**：【Who】{row.get('客訴_Who','')}｜【Where】{row.get('客訴_Where','')}｜【When】{row.get('客訴_When','')}\n- 【What】{row.get('客訴_What','') or row.get('直播中客訴','')}\n- 【Why】{row.get('客訴_Why','')}\n- 【How】{row.get('客訴_How','')}")
                        if row.get("有無8D") == "是" and row.get("8D檔案編碼"):
                            st.download_button(f"📎 下載 8D 檔案 ({row.get('8D檔名', '8D.xlsx')})", data=base64.b64decode(row["8D檔案編碼"]), file_name=f"8D_{r_id}_{row.get('8D檔名', '8D.xlsx')}", key=f"dl_ct8d_{r_id}")

                        st.markdown("---")
                        f_c1, f_c2, f_c3 = st.columns(3)
                        f_c1.info(f"💬 整體操台感受：\n\n{row.get('直播台整體操台感受', '')}")
                        f_c2.success(f"📌 主管提醒改善事項：\n\n{row.get('主管提醒改善事項', '') or '無'}")
                        f_c3.warning(f"💡 提案改善：\n\n{row.get('提案改善', '無')}")

                        fc1, fc2, fc3 = st.columns([2.5, 1.2, 1.3])
                        fc1.caption(f"單號：{r_id} ｜ 送出時間：{row.get('送出時間', '')}")
                        if is_editing_ct:
                            if fc2.button("❌ 收合修改", key=f"cl_cted_{r_id}", use_container_width=True): st.session_state.edit_report_id = None; st.rerun()
                        else:
                            if fc2.button("✏️ 修改操台回報", key=f"ct_ed_{r_id}", use_container_width=True): st.session_state.edit_report_id = r_id; st.rerun()
                        with fc3:
                            if st.session_state.confirm_del_id == f"adm_ct_{r_id}":
                                dc1, dc2 = st.columns(2)
                                if dc1.button("✅ 確定", key=f"y_ct_{r_id}", type="primary"): delete_caotai_by_id(ws_caotai, ws_caotai_compact, ws_caotai_comp, r_id); st.session_state.confirm_del_id = None; st.rerun()
                                if dc2.button("❌ 取消", key=f"n_ct_{r_id}"): st.session_state.confirm_del_id = None; st.rerun()
                            else:
                                if st.button("🗑️ 刪除此筆", key=f"ct_del_{r_id}", use_container_width=True): st.session_state.confirm_del_id = f"adm_ct_{r_id}"; st.rerun()

                        if is_editing_ct:
                            with st.form(key=f"edit_ct_form_{r_id}"):
                                ec1, ec2, ec3 = st.columns(3)
                                ed_ct_date = ec1.date_input("📅 日期", parse_date_safe(row.get("回報日期", "")))
                                ed_ct_name = ec2.text_input("👤 姓名", value=str(row.get("姓名", "")))
                                ed_ct_dev = ec3.text_input("📱 工作機或賴名稱", value=str(row.get("工作機或賴名稱", "")))
                                ed_shifts_txt = st.text_area("🕒 上下班時間 (多段班請一行寫一段，如 13:00-17:00 換行 20:00-02:30，系統自動算跨日與總工時)",
                                                             value=format_shifts_for_edit(row.get("上班時間", ""), row.get("下班時間", "")), height=70)
                                ed_ct_sname = st.text_input("🎥 直播台名稱", value=str(row.get("直播台名稱", "")))
                                ed_ct_furl = st.text_area("🔗 粉專網址", value=str(row.get("直播台粉專網址", "")), height=60)
                                ec8, ec9 = st.columns(2)
                                ed_ct_optime = ec8.text_area("⏱️ 操台時間", value=str(row.get("操台時間", "")), height=60)
                                ed_ct_metrics = ec9.text_area("📊 留言/點讚/分享", value=str(row.get("總留言_點讚_分享", "")), height=60)
                                ed_ct_req = st.text_area("📋 直播需求", value=str(row.get("直播需求", "")), height=65)
                                ed_ct_met = st.text_area("✅ 需求是否達成", value=str(row.get("直播需求是否達成", "")), height=60)
                                ed_ct_other = st.text_area("🛠️ 操台外處理事項", value=str(row.get("操台外處理事項", "")), height=70)
                                ed_ct_feel = st.text_area("💬 整體操台感受", value=str(row.get("直播台整體操台感受", "")), height=60)
                                ed_ct_sup = st.text_area("📌 主管提醒改善事項", value=str(row.get("主管提醒改善事項", "")) or "無", height=60)
                                ed_ct_prop = st.text_area("💡 提案改善", value=str(row.get("提案改善", "")), height=60)

                                if st.form_submit_button("💾 儲存操台修改內容 (自動重算多段工時)", type="primary", use_container_width=True):
                                    s_list, e_list, new_w_m = parse_shifts_from_multiline_text(ed_shifts_txt)
                                    if not s_list:
                                        st.error("❌ 上下班時間格式有誤，請輸入如 13:00-22:00 的格式！")
                                    else:
                                        upd_ct = dict(row)
                                        upd_ct.update({
                                            "回報日期": normalize_date_str(ed_ct_date), "姓名": ed_ct_name.strip(),
                                            "工作機或賴名稱": ed_ct_dev.strip(),
                                            "上班時間": "\n".join(s_list), "下班時間": "\n".join(e_list),
                                            "上班總分鐘": new_w_m, "上班總時數(小時)": round(new_w_m / 60.0, 2),
                                            "直播台名稱": ed_ct_sname.strip(), "直播台粉專網址": ed_ct_furl.strip(),
                                            "操台時間": ed_ct_optime.strip(), "操台時數(小時)": parse_caotai_hours(ed_ct_optime),
                                            "總留言_點讚_分享": ed_ct_metrics.strip(), "直播需求": ed_ct_req.strip(),
                                            "直播需求是否達成": ed_ct_met.strip(), "操台外處理事項": ed_ct_other.strip(),
                                            "直播台整體操台感受": ed_ct_feel.strip(), "主管提醒改善事項": ed_ct_sup.strip() or "無",
                                            "提案改善": ed_ct_prop.strip()
                                        })
                                        write_dict_to_sheet(ws_caotai, CAOTAI_HEADERS, upd_ct, target_report_id=r_id)
                                        sync_caotai_pretty_sheet(ws_caotai, ws_caotai_compact, ws_caotai_comp)
                                        clear_all_caches()
                                        st.session_state.edit_report_id = None; st.rerun()

        # 分頁 3：季稽核專區
        with dept_tab3:
            st.subheader("🚨 季稽核專區：自訂時間區間 / 季度 客訴調閱與匯出中心")
            df_gold_all, df_ct_all = load_reports_df(), load_caotai_df()
            audit_rows = []
            if not df_gold_all.empty:
                for _, r in df_gold_all[(df_gold_all["有無客訴"] == "有") | (df_gold_all.get("有無8D", "否") == "是")].iterrows():
                    audit_rows.append({
                        "回報單號": r.get("回報單號", ""), "部門": "黃金部門",
                        "回報日期": normalize_date_str(r.get("回報日期", "")), "季度": get_quarter_label(r.get("回報日期", "")),
                        "填表同仁": r.get("姓名", ""), "1.【Who】對象/人員": f"客人：{r.get('客訴_Who', '')} (填表：{r.get('姓名', '')})",
                        "2.【Where】管道/直播台": r.get("客訴_Where", ""), "3.【When】發生時間/訂單日": r.get("客訴_When", ""),
                        "4.【What】客訴問題詳情": r.get("客訴_What", ""), "5.【Why】背後原因分析": r.get("客訴_Why", ""),
                        "6.【How】處理方式與改善結果": r.get("客訴_How", ""),
                        "有無8D": r.get("有無8D", "否"), "8D檔名": r.get("8D檔名", ""), "8D檔案編碼": r.get("8D檔案編碼", "")
                    })
            if not df_ct_all.empty:
                for _, r in df_ct_all.iterrows():
                    if is_caotai_row_complaint(r) or r.get("有無8D") == "是":
                        audit_rows.append({
                            "回報單號": r.get("回報單號", ""), "部門": "操台部門",
                            "回報日期": normalize_date_str(r.get("回報日期", "")), "季度": get_quarter_label(r.get("回報日期", "")),
                            "填表同仁": r.get("姓名", ""), "1.【Who】對象/人員": r.get("客訴_Who", "") or r.get("姓名", ""),
                            "2.【Where】管道/直播台": r.get("客訴_Where", "") or r.get("直播台名稱", ""),
                            "3.【When】發生時間/訂單日": r.get("客訴_When", ""),
                            "4.【What】客訴問題詳情": r.get("客訴_What", "") or r.get("直播中客訴", ""),
                            "5.【Why】背後原因分析": r.get("客訴_Why", ""), "6.【How】處理方式與改善結果": r.get("客訴_How", ""),
                            "有無8D": r.get("有無8D", "否"), "8D檔名": r.get("8D檔名", ""), "8D檔案編碼": r.get("8D檔案編碼", "")
                        })
            df_audit = pd.DataFrame(audit_rows)

            au_c1, au_c2, au_c3, au_c4 = st.columns([1.3, 1.5, 1.2, 1.2])
            audit_time_mode = au_c1.radio("⏳ 時間篩選方式", ["📅 自訂起訖日期區間", "📆 依季度快速選擇"], key="au_tmode")
            with au_c2:
                if audit_time_mode == "📅 自訂起訖日期區間":
                    start_d = st.date_input("開始日期", get_tw_today() - datetime.timedelta(days=90), key="au_sd")
                    end_d = st.date_input("結束日期", get_tw_today(), key="au_ed")
                    period_label = f"{start_d.strftime('%Y%m%d')}_{end_d.strftime('%Y%m%d')}"
                else:
                    q_list = ["全部季度"] + (sorted(df_audit["季度"].unique().tolist(), reverse=True) if not df_audit.empty else [])
                    sel_au_q = st.selectbox("選擇稽核季度", q_list, key="au_qsel")
                    period_label = sel_au_q
            sel_au_dept = au_c3.selectbox("🏢 篩選部門", ["全部部門", "操台部門", "黃金部門"], key="au_dept")
            sel_au_staff = au_c4.selectbox("👤 篩選同仁", ["全部同仁"] + (sorted(df_audit["填表同仁"].unique().tolist()) if not df_audit.empty else []), key="au_staff")

            if df_audit.empty:
                st.success("🎉 目前系統內尚無任何客訴紀錄！")
            else:
                f_audit = df_audit.copy()
                if audit_time_mode == "📅 自訂起訖日期區間":
                    f_audit = f_audit[(f_audit["回報日期"] >= normalize_date_str(start_d)) & (f_audit["回報日期"] <= normalize_date_str(end_d))]
                elif sel_au_q != "全部季度":
                    f_audit = f_audit[f_audit["季度"] == sel_au_q]
                if sel_au_dept != "全部部門": f_audit = f_audit[f_audit["部門"] == sel_au_dept]
                if sel_au_staff != "全部同仁": f_audit = f_audit[f_audit["填表同仁"] == sel_au_staff]

                st.divider()
                ak1, ak2, ak3 = st.columns(3)
                ak1.metric("🚨 區間客訴總件數", f"{len(f_audit)} 件")
                ak2.metric("🎛️ 操台部門客訴", f"{len(f_audit[f_audit['部門'] == '操台部門'])} 件")
                ak3.metric("✨ 黃金部門客訴", f"{len(f_audit[f_audit['部門'] == '黃金部門'])} 件")

                if not f_audit.empty:
                    f_audit = f_audit.sort_values(by="回報日期", ascending=False).reset_index(drop=True)
                    au_stat_df = f_audit.groupby(["部門", "填表同仁"]).agg(
                        區間客訴總次數=("回報單號", "count"),
                        涉及直播台或管道=("2.【Where】管道/直播台", lambda x: "、".join(sorted(set(str(v) for v in x if str(v).strip()))))
                    ).reset_index().sort_values(by="區間客訴總次數", ascending=False)
                    st.dataframe(au_stat_df, use_container_width=True)

                    au_output = io.BytesIO()
                    with pd.ExcelWriter(au_output, engine="openpyxl") as writer:
                        au_stat_df.to_excel(writer, index=False, sheet_name="1_季稽核客訴次數統計")
                        f_audit.drop(columns=["回報單號", "8D檔案編碼"], errors="ignore").to_excel(writer, index=False, sheet_name="2_跨部門客訴5W1H總表")
                        style_excel_File(writer)
                    st.download_button(f"📥 點擊下載【季稽核客訴專案報表】Excel ({period_label}.xlsx)", data=au_output.getvalue(), file_name=f"Audit_Complaints_{period_label}.xlsx", type="primary", use_container_width=True)

                    for idx, a_row in f_audit.iterrows():
                        with st.expander(f"🚨 {a_row['回報日期']} ({a_row['季度']}) ｜ 【{a_row['部門']}】{a_row['填表同仁']} ｜ 📍 {a_row['2.【Where】管道/直播台']}", expanded=(idx == 0)):
                            st.markdown(f"- **1. 【Who】：** {a_row['1.【Who】對象/人員']}\n- **2. 【Where】：** {a_row['2.【Where】管道/直播台']}\n- **3. 【When】：** {a_row['3.【When】發生時間/訂單日']}\n- **4. 【What】：** {a_row['4.【What】客訴問題詳情']}\n- **5. 【Why】：** {a_row['5.【Why】背後原因分析']}\n- **6. 【How】：** {a_row['6.【How】處理方式與改善結果']}")
                            if a_row.get("有無8D") == "是" and a_row.get("8D檔案編碼"):
                                st.download_button(f"📎 下載 8D ({a_row.get('8D檔名', '8D.xlsx')})", data=base64.b64decode(a_row["8D檔案編碼"]), file_name=f"8D_{a_row['回報單號']}_{a_row.get('8D檔名', '8D.xlsx')}", key=f"dl_au_{idx}")
# === 程式碼完整結束 ===
