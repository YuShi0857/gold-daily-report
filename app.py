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
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Alignment

# 頁面基本設定
st.set_page_config(page_title="黃金部門工作回報系統", layout="wide", page_icon="✨")

# 設定台灣標準時區 (UTC+8)
TW_TZ = datetime.timezone(datetime.timedelta(hours=8))

def get_tw_now():
    return datetime.datetime.now(TW_TZ)

def get_tw_today():
    return get_tw_now().date()

# 標準中文標題列定義
REPORT_HEADERS = [
    "回報單號", "送出時間", "回報日期", "帳號", "姓名",
    "上班時間", "下班時間", 
    "包貨開始", "包貨結束", "包貨件數", "包貨總分鐘", "單件包貨耗時(分)",
    "客服開始", "客服結束", "客服總分鐘",
    "寄貨件數", 
    "直播開始", "直播結束", "直播總分鐘",
    "有無客訴", "客訴_Who", "客訴_Where", "客訴_When", "客訴_What", "客訴_Why", "客訴_How",
    "有無8D", "8D檔名", "8D檔案編碼",
    "今日處理事項", "提案改善", "對同仁想法"
]

COMPACT_HEADERS = [
    "回報日期", "姓名", "上下班時段", "包貨紀錄(時段/件數/單件耗時)", 
    "寄貨件數", "客服紀錄(時段/時數)", "直播紀錄(時段/時數)", 
    "今日處理事項明細", "提案改善", "對同仁想法", "有無客訴", "回報單號"
]

COMPLAINT_HEADERS = [
    "回報日期", "填表同仁", "1.客人(Who)", "2.管道(Where)", 
    "3.訂單與客訴日(When)", "4.問題詳情(What)", "5.原因分析(Why)", 
    "6.解決與改善(How)", "有無附8D報告", "8D檔案名稱", "回報單號"
]

# ================= 1. 輔助與安全函數 (含防掉線 Token 驗證) =================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def make_auth_token(username, pass_hash):
    """產生專屬防掉線安全憑證，讓網頁就算重整也不會踢出登入"""
    raw = f"{username}_{pass_hash}_gold_stay_logged_in"
    return hashlib.sha256(raw.encode()).hexdigest()[:16]

def calc_minutes(t_start, t_end):
    dt1 = datetime.datetime.combine(get_tw_today(), t_start)
    dt2 = datetime.datetime.combine(get_tw_today(), t_end)
    if dt2 < dt1:
        dt2 += datetime.timedelta(days=1)
    return round((dt2 - dt1).total_seconds() / 60.0, 2)

def calc_minutes_str(s1, s2):
    try:
        if not s1 or not s2 or s1 == "無" or s2 == "無" or (s1 == "00:00" and s2 == "00:00"):
            return 0
        t1 = datetime.datetime.strptime(s1.strip(), "%H:%M").time()
        t2 = datetime.datetime.strptime(s2.strip(), "%H:%M").time()
        return calc_minutes(t1, t2)
    except:
        return 0

def parse_vals_to_df(all_vals):
    if not all_vals or len(all_vals) <= 1:
        return pd.DataFrame()
    headers = [h.strip() for h in all_vals[0]]
    valid_cols, clean_headers, seen = [], [], set()
    for i, h in enumerate(headers):
        if h != "" and h not in seen:
            valid_cols.append(i)
            clean_headers.append(h)
            seen.add(h)
    data = []
    for r in all_vals[1:]:
        if not any(str(cell).strip() for cell in r):
            continue
        row_padded = r + [""] * max(0, len(headers) - len(r))
        data.append([row_padded[i] for i in valid_cols])
    return pd.DataFrame(data, columns=clean_headers)

# ================= 2. 極速資料庫連線與快取機制 =================
def beautify_google_sheet(ws, num_cols):
    try:
        ws.freeze(rows=1)
        end_col = chr(64 + min(num_cols, 26))
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
        
    return ws_users, ws_reports, ws_compact, ws_comp

@st.cache_data(ttl=180)
def load_users_df():
    ws_users, _, _, _ = get_worksheets_cached()
    for attempt in range(3):
        try:
            return parse_vals_to_df(ws_users.get_all_values())
        except Exception as e:
            if attempt == 2:
                raise e
            time.sleep(1.5)

@st.cache_data(ttl=180)
def load_reports_df():
    _, ws_reports, _, _ = get_worksheets_cached()
    for attempt in range(3):
        try:
            return parse_vals_to_df(ws_reports.get_all_values())
        except Exception as e:
            if attempt == 2:
                raise e
            time.sleep(1.5)

def clear_all_caches():
    load_users_df.clear()
    load_reports_df.clear()

def sync_pretty_sheets(ws_reports, ws_compact, ws_comp, apply_style=False):
    all_vals = ws_reports.get_all_values()
    df = parse_vals_to_df(all_vals)
    compact_rows = [COMPACT_HEADERS]
    comp_rows = [COMPLAINT_HEADERS]
    
    if not df.empty:
        df_sorted = df.sort_values(by=["回報日期", "送出時間"], ascending=[False, False])
        for _, r in df_sorted.iterrows():
            p_cnt = pd.to_numeric(r.get("包貨件數", 0), errors="coerce") or 0
            cs_min = pd.to_numeric(r.get("客服總分鐘", 0), errors="coerce") or 0
            lv_min = pd.to_numeric(r.get("直播總分鐘", 0), errors="coerce") or 0
            
            pack_str = f"{r.get('包貨開始','無')}~{r.get('包貨結束','無')}\n({int(p_cnt)}件 / {r.get('單件包貨耗時(分)',0)}分)" if p_cnt > 0 else "無"
            cs_str = f"{r.get('客服開始','無')}~{r.get('客服結束','無')}\n({round(cs_min/60, 1)}小時)" if cs_min > 0 else "無"
            lv_str = f"{r.get('直播開始','無')}~{r.get('直播結束','無')}\n({round(lv_min/60, 1)}小時)" if lv_min > 0 else "無"
            
            compact_rows.append([
                str(r.get("回報日期", "")),
                str(r.get("姓名", "")),
                f"{r.get('上班時間', '')} ~ {r.get('下班時間', '')}",
                pack_str,
                str(r.get("寄貨件數", 0)),
                cs_str,
                lv_str,
                str(r.get("今日處理事項", "")),
                str(r.get("提案改善", "")) or "無",
                str(r.get("對同仁想法", "")) or "無",
                str(r.get("有無客訴", "無")),
                str(r.get("回報單號", ""))
            ])
            
            if r.get("有無客訴") == "有" or r.get("有無8D") == "是":
                comp_rows.append([
                    str(r.get("回報日期", "")),
                    str(r.get("姓名", "")),
                    str(r.get("客訴_Who", "")),
                    str(r.get("客訴_Where", "")),
                    str(r.get("客訴_When", "")),
                    str(r.get("客訴_What", "")),
                    str(r.get("客訴_Why", "")),
                    str(r.get("客訴_How", "")),
                    str(r.get("有無8D", "否")),
                    str(r.get("8D檔名", "")),
                    str(r.get("回報單號", ""))
                ])
                
    ws_compact.clear()
    ws_compact.update(compact_rows)
    ws_comp.clear()
    ws_comp.update(comp_rows)
    
    if apply_style:
        beautify_google_sheet(ws_compact, len(COMPACT_HEADERS))
        beautify_google_sheet(ws_comp, len(COMPLAINT_HEADERS))

def find_user_row_idx(ws_users, target_username):
    col_vals = ws_users.col_values(1)
    for idx, val in enumerate(col_vals):
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

def delete_existing_same_day_report(ws_reports, username, report_date_str):
    all_vals = ws_reports.get_all_values()
    overwritten = False
    for idx in range(len(all_vals) - 1, 0, -1):
        r = all_vals[idx]
        if len(r) >= 4 and str(r[2]).strip() == str(report_date_str).strip() and str(r[3]).strip() == str(username).strip():
            ws_reports.delete_rows(idx + 1)
            overwritten = True
    return overwritten

def get_quarter_label(date_str):
    try:
        dt = pd.to_datetime(date_str)
        q = (dt.month - 1) // 3 + 1
        return f"{dt.year}-Q{q}"
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
                val_str = str(cell.value or "")
                lines = val_str.split("\n")
                line_max = max([len(l.encode("utf-8")) for l in lines]) if lines else 0
                if line_max > max_len:
                    max_len = line_max
            ws.column_dimensions[col_letter].width = min(max(max_len * 0.65 + 4, 12), 45)

def send_reset_email(receiver_email, new_pass):
    try:
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = st.secrets["smtp"]["sender_email"]
        sender_password = st.secrets["smtp"]["app_password"]

        msg = MIMEText(f"您好，\n\n您的「黃金部門工作回報系統」密碼已重設。\n您的臨時新密碼為：{new_pass}\n\n請使用此密碼登入後，立即至左側選單「修改密碼」更新您的密碼。", "plain", "utf-8")
        msg["Subject"] = "【系統通知】黃金部門工作回報系統 - 密碼重設"
        msg["From"] = sender_email
        msg["To"] = receiver_email

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, [receiver_email], msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"郵件發送失敗，請確認後台設定：{e}")
        return False

# ================= 3. 狀態管理與「自動接回登入」機制 =================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_info = {}
if "task_rows" not in st.session_state:
    st.session_state.task_rows = 3 # 預設直接提供 3 列時段，減少頻繁按按鈕重整
if "form_ver" not in st.session_state:
    st.session_state.form_ver = 0
if "submit_success_msg" not in st.session_state:
    st.session_state.submit_success_msg = ""
if "confirm_del_id" not in st.session_state:
    st.session_state.confirm_del_id = None

# 【關鍵防掉線機制】：若網頁因手機網路瞬斷或重整導致 session 重置，自動從網址安全憑證無感恢復登入！
if not st.session_state.logged_in:
    saved_u = st.query_params.get("u", "")
    saved_t = st.query_params.get("t", "")
    if saved_u and saved_t:
        try:
            users_df = load_users_df()
            if not users_df.empty and saved_u in users_df["登入帳號"].values:
                urow = users_df[users_df["登入帳號"] == saved_u].iloc[0]
                expected_t = make_auth_token(saved_u, urow["密碼雜湊"])
                if saved_t == expected_t:
                    st.session_state.logged_in = True
                    st.session_state.user_info = {
                        "username": urow["登入帳號"],
                        "password_hash": urow["密碼雜湊"],
                        "name": urow["真實姓名"],
                        "email": urow["綁定信箱"],
                        "role": urow["系統權限"]
                    }
        except:
            pass

# ================= 4. 未登入畫面 =================
if not st.session_state.logged_in:
    st.title("✨ 黃金部門工作回報系統")
    tab1, tab2, tab3 = st.tabs(["🔑 登入系統", "📝 新人註冊", "❓ 忘記密碼"])
    
    with tab1:
        st.subheader("員工/管理員登入")
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
                            "username": user_row["登入帳號"],
                            "password_hash": user_row["密碼雜湊"],
                            "name": user_row["真實姓名"],
                            "email": user_row["綁定信箱"],
                            "role": user_row["系統權限"]
                        }
                        # 寫入網址防掉線憑證，確保重整絕不踢出
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
                    ws_users, _, _, _ = get_worksheets_cached()
                    users_df = load_users_df()
                    if not users_df.empty and reg_user.strip() in users_df["登入帳號"].values:
                        st.error("❌ 此帳號已被註冊，請換一個！")
                    else:
                        ws_users.append_row([
                            reg_user.strip(), hash_password(reg_pass1.strip()), reg_name.strip(), reg_email.strip(), "staff", get_tw_now().strftime("%Y-%m-%d %H:%M:%S")
                        ])
                        clear_all_caches()
                        st.success("✅ 註冊成功！請切換至左上角「🔑 登入系統」分頁進行登入。")

    with tab3:
        st.subheader("忘記密碼 / 重新寄送臨時密碼")
        with st.form("fp_form"):
            fp_user = st.text_input("請輸入您的帳號")
            fp_email = st.text_input("請輸入註冊時綁定的信箱")
            fp_submit = st.form_submit_button("📧 寄送重設密碼信", type="primary")
            
        if fp_submit:
            with st.spinner("🔄 正在驗證信箱並發送郵件..."):
                ws_users, _, _, _ = get_worksheets_cached()
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
        ws_users, ws_reports, ws_compact, ws_comp = get_worksheets_cached()
    except Exception as e:
        st.warning("⏳ 系統正在冷卻連線頻率，請等待約 10 秒後按 F5 重新整理網頁（您的登入狀態已鎖定保護，不會被登出）！")
        st.stop()

    user = st.session_state.user_info
    st.sidebar.title(f"👋 您好，{user['name']}")
    st.sidebar.caption(f"身份：{'👑 管理員' if user['role'] == 'admin' else '👤 部門同仁'}")
    st.sidebar.caption(f"🕒 台灣時間：{get_tw_now().strftime('%Y-%m-%d %H:%M')}")
    st.sidebar.success("🔒 防掉線保護已啟動（重整不登出）")
    
    menu_options = ["📝 每日工作回報", "🔒 修改密碼"]
    if user["role"] == "admin":
        menu_options = ["📊 主管後台數據中心", "👥 帳號與權限管理", "📝 每日工作回報", "🔒 修改密碼"]
        
    choice = st.sidebar.radio("功能選單", menu_options)
    
    if st.sidebar.button("🚪 登出系統"):
        st.session_state.logged_in = False
        st.session_state.user_info = {}
        st.query_params.clear()
        st.rerun()

    # ---------------- 功能 A：每日工作回報 ----------------
    if choice == "📝 每日工作回報":
        st.title("📝 黃金部門 - 每日工作回報單")
        
        if st.session_state.submit_success_msg:
            st.balloons()
            st.success(st.session_state.submit_success_msg)
            st.session_state.submit_success_msg = ""

        st.info(f"📌 填表人：**{user['name']}** ｜ 🛡️ **防消失升級**：目前已開啟「穩定填表模式」，調整時間與打字時網頁**不會閃爍刷新**，且登入狀態已鎖定，請安心填寫！")
        
        fv = st.session_state.form_ver
        
        # 先在最上方選擇今日有哪些特殊項目（勾選後下方表單一次展開，打字過程零刷新！）
        st.markdown("### 📌 步驟一：請先勾選今日有執行的作業項目")
        st.caption("（有執行的打勾即可，下方填表區在您打字與調整時間時，完全不會刷新畫面！）")
        chk_c1, chk_c2, chk_c3, chk_c4, chk_c5, chk_c6 = st.columns(6)
        with chk_c1:
            has_pack = st.checkbox("📦 今日有包貨", key=f"hp_{fv}")
        with chk_c2:
            has_ship = st.checkbox("🚚 今日有寄貨", key=f"hs_{fv}")
        with chk_c3:
            has_cs = st.checkbox("💬 今日有客服", key=f"hcs_{fv}")
        with chk_c4:
            has_live = st.checkbox("🎥 今日有直播", key=f"hlv_{fv}")
        with chk_c5:
            has_complaint_chk = st.checkbox("🚨 今日有客訴", key=f"hc_{fv}")
        with chk_c6:
            has_8d = st.checkbox("📄 需附 8D 檔案", key=f"h8d_{fv}")

        col_btn1, col_btn2, _ = st.columns([1.5, 1.5, 4])
        with col_btn1:
            if st.button("➕ 增加處理事項欄位 (目前 " + str(st.session_state.task_rows) + " 條)"):
                st.session_state.task_rows += 2
                st.rerun()
        with col_btn2:
            if st.button("➖ 減少處理事項欄位") and st.session_state.task_rows > 1:
                st.session_state.task_rows = max(1, st.session_state.task_rows - 1)
                st.rerun()

        # 【核心防消失設計】：使用 st.form 將所有時間與文字輸入框保護起來！
        # 員工在裡面打字、改時間，瀏覽器完全不會重整、不會發送請求，直到最後按下送出！
        with st.form(key=f"daily_report_form_{fv}", clear_on_submit=False):
            st.markdown("### 📝 步驟二：填寫今日工作內容（打字過程零刷新保護中）")
            report_date = st.date_input("📅 回報歸屬日期 (預設為台灣今日)", get_tw_today())
            
            st.markdown("#### 一、 基礎出勤與作業數據")
            col1, col2 = st.columns(2)
            with col1:
                work_start = st.time_input("🕒 上班時間", datetime.time(9, 0), step=60)
            with col2:
                work_end = st.time_input("🕕 下班時間", datetime.time(18, 0), step=60)
                
            pack_start = datetime.time(10, 0)
            pack_end = datetime.time(12, 0)
            pack_count = 0
            if has_pack:
                st.markdown("##### 📦 包貨紀錄")
                c_p1, c_p2, c_p3 = st.columns(3)
                with c_p1:
                    pack_start = st.time_input("包貨開始時間", datetime.time(10, 0), step=60)
                with c_p2:
                    pack_end = st.time_input("包貨結束時間", datetime.time(12, 0), step=60)
                with c_p3:
                    pack_count = st.number_input("個人包貨件數", min_value=1, value=1, step=1)

            ship_count = 0
            if has_ship:
                st.markdown("##### 🚚 寄貨紀錄")
                ship_count = st.number_input("今日寄貨件數", min_value=1, value=1, step=1)
                
            cs_start = datetime.time(9, 0)
            cs_end = datetime.time(18, 0)
            if has_cs:
                st.markdown("##### 💬 客服回覆紀錄")
                c_l1, c_l2 = st.columns(2)
                with c_l1:
                    cs_start = st.time_input("客服回覆開始", datetime.time(9, 0), step=60)
                with c_l2:
                    cs_end = st.time_input("客服回覆結束", datetime.time(18, 0), step=60)

            live_start = datetime.time(19, 0)
            live_end = datetime.time(22, 0)
            if has_live:
                st.markdown("##### 🎥 直播紀錄")
                c_l3, c_l4 = st.columns(2)
                with c_l3:
                    live_start = st.time_input("今日直播開始", datetime.time(19, 0), step=60)
                with c_l4:
                    live_end = st.time_input("今日直播結束", datetime.time(22, 0), step=60)

            c_who = c_where = c_when = c_what = c_why = c_how = ""
            if has_complaint_chk:
                st.divider()
                st.markdown("#### 二、 🚨 客訴 5W1H 詳細回報")
                c_who = st.text_input("1. 【Who】客人本名 + TK名 / 帳號名稱")
                c_where = st.selectbox("2. 【Where】客訴發生管道", ["直播上", "LINE客服上", "LINE社群", "其他"])
                c_when = st.text_input("3. 【When】幾月幾號的訂單 及 客訴發生日期")
                c_what = st.text_area("4. 【What】發生什麼事 / 商品的問題 (請詳細打字說明)")
                c_why = st.text_area("5. 【Why】客訴背後原因分析")
                c_how = st.text_area("6. 【How】執行方法 / 解決方案 / 未來如何改善")

            uploaded_8d = None
            if has_8d:
                st.divider()
                st.markdown("#### 📄 重大異常 8D 報告檔案上傳")
                uploaded_8d = st.file_uploader("請上傳填妥之 8D 檔案 (支援 Excel, Word, PDF，建議小於 1MB)", type=["xlsx", "xls", "docx", "doc", "pdf"])

            st.divider()
            st.markdown("#### 三、 今日處理事項明細")
            st.caption("請填寫幾點幾分到幾點幾分做了哪些事（空白未填的列會自動忽略，不夠可在上方點『➕ 增加處理事項欄位』）")
            
            raw_tasks_inputs = []
            for i in range(st.session_state.task_rows):
                t_col1, t_col2, t_col3 = st.columns([1.5, 1.5, 5])
                default_h1 = min(9 + i, 22)
                default_h2 = min(10 + i, 23)
                with t_col1:
                    ts = st.time_input(f"開始時間 #{i+1}", datetime.time(default_h1, 0), key=f"ts_{fv}_{i}", step=60)
                with t_col2:
                    te = st.time_input(f"結束時間 #{i+1}", datetime.time(default_h2, 0), key=f"te_{fv}_{i}", step=60)
                with t_col3:
                    td = st.text_input(f"處理事項內容 #{i+1}", placeholder="請輸入該時段具體完成的事項（無則留空即可）...", key=f"td_{fv}_{i}")
                raw_tasks_inputs.append((ts, te, td))

            st.divider()
            st.markdown("#### 四、 提案改善與同仁想法回饋")
            improvement_proposal = st.text_area("💡 提案改善欄位 (針對作業流程、商品、工具有任何優化建議皆可提出)")
            peer_feedback = st.text_area("🤝 針對其他同仁有任何想法或建議 (僅主管後台可見)")

            submit_btn = st.form_submit_button("🚀 確認無誤，送出今日回報", type="primary", use_container_width=True)

        if submit_btn:
            tasks_list = []
            for ts, te, td in raw_tasks_inputs:
                if td.strip():
                    tasks_list.append(f"【{ts.strftime('%H:%M')} - {te.strftime('%H:%M')}】{td.strip()}")
                    
            has_complaint = "有" if has_complaint_chk else "無"
            file_8d_name, file_8d_b64 = "", ""
            if has_8d and uploaded_8d is not None:
                file_8d_name = uploaded_8d.name
                file_8d_b64 = base64.b64encode(uploaded_8d.read()).decode("utf-8")

            if has_complaint == "有" and not all([c_who.strip(), c_when.strip(), c_what.strip(), c_why.strip(), c_how.strip()]):
                st.error("❌ 您勾選了「今日有客訴」，請將 Who / When / What / Why / How 所有欄位填寫完整！（您剛剛填寫的文字都還保留在上方格子內，補齊後再按送出即可）")
            elif has_8d and not file_8d_name:
                st.error("❌ 您勾選了「需附 8D 檔案」，請務必選擇檔案上傳！")
            elif not tasks_list:
                st.error("❌ 請至少填寫一條「今日處理事項」！（填寫後再按送出即可）")
            else:
                with st.spinner("🚀 正在安全傳送您的回報資料..."):
                    pack_start_str = pack_start.strftime("%H:%M") if has_pack else "無"
                    pack_end_str = pack_end.strftime("%H:%M") if has_pack else "無"
                    pack_duration = calc_minutes(pack_start, pack_end) if has_pack else 0
                    min_per_pack = round(pack_duration / pack_count, 2) if (has_pack and pack_count > 0) else 0

                    cs_start_str = cs_start.strftime("%H:%M") if has_cs else "無"
                    cs_end_str = cs_end.strftime("%H:%M") if has_cs else "無"
                    cs_duration = calc_minutes(cs_start, cs_end) if has_cs else 0

                    live_start_str = live_start.strftime("%H:%M") if has_live else "無"
                    live_end_str = live_end.strftime("%H:%M") if has_live else "無"
                    live_duration = calc_minutes(live_start, live_end) if has_live else 0

                    was_overwritten = delete_existing_same_day_report(ws_reports, user["username"], str(report_date))
                    
                    now_tw = get_tw_now()
                    report_id = now_tw.strftime("%Y%m%d%H%M%S") + "_" + user["username"]
                    ws_reports.append_row([
                        report_id,
                        now_tw.strftime("%Y-%m-%d %H:%M:%S"),
                        str(report_date),
                        user["username"],
                        user["name"],
                        work_start.strftime("%H:%M"),
                        work_end.strftime("%H:%M"),
                        pack_start_str,
                        pack_end_str,
                        pack_count if has_pack else 0,
                        pack_duration,
                        min_per_pack,
                        cs_start_str,
                        cs_end_str,
                        cs_duration,
                        ship_count if has_ship else 0,
                        live_start_str,
                        live_end_str,
                        live_duration,
                        has_complaint,
                        c_who, c_where, c_when, c_what, c_why, c_how,
                        "是" if has_8d else "否",
                        file_8d_name,
                        file_8d_b64,
                        "\n".join(tasks_list),
                        improvement_proposal,
                        peer_feedback
                    ])
                    
                    sync_pretty_sheets(ws_reports, ws_compact, ws_comp)
                    clear_all_caches()
                    
                    st.session_state.task_rows = 3
                    st.session_state.form_ver += 1
                    if was_overwritten:
                        st.session_state.submit_success_msg = f"✅ 【覆蓋成功】您於 {report_date} 的舊回報已自動刪除，並更新為最新送出的內容（送出時間：{now_tw.strftime('%Y-%m-%d %H:%M:%S')}）！下方表單已自動清空。"
                    else:
                        st.session_state.submit_success_msg = f"🎉 【送出成功】您於 {report_date} 的工作回報已成功送出（送出時間：{now_tw.strftime('%Y-%m-%d %H:%M:%S')}）！下方表單已自動清空。"
                    st.rerun()

        # 同仁自行撤回近期回報紀錄（含二次確認防呆機制）
        st.divider()
        with st.expander("🗂️ 檢視 / 撤回我近期已送出的回報紀錄"):
            df_check = load_reports_df()
            if not df_check.empty:
                my_df = df_check[df_check["帳號"] == user["username"]].sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).head(10)
                if not my_df.empty:
                    for _, mrow in my_df.iterrows():
                        m_id = mrow["回報單號"]
                        mc1, mc2 = st.columns([4, 2])
                        mc1.markdown(f"**📅 日期：{mrow['回報日期']}** ｜ 送出時間：`{mrow['送出時間']}` ｜ 包貨：{mrow['包貨件數']}件 / 寄貨：{mrow['寄貨件數']}件 ｜ 事項：{str(mrow['今日處理事項'])[:30]}...")
                        
                        with mc2:
                            if st.session_state.confirm_del_id == f"self_{m_id}":
                                st.warning(f"⚠️ 確定要撤回刪除 {mrow['回報日期']} 的回報嗎？")
                                sc1, sc2 = st.columns(2)
                                if sc1.button("✅ 確定刪除", key=f"yes_self_{m_id}", type="primary", use_container_width=True):
                                    delete_report_by_id(ws_reports, ws_compact, ws_comp, m_id)
                                    st.session_state.confirm_del_id = None
                                    st.success(f"✅ 已成功刪除您於 {mrow['回報日期']} 的回報紀錄！")
                                    st.rerun()
                                if sc2.button("❌ 取消", key=f"no_self_{m_id}", use_container_width=True):
                                    st.session_state.confirm_del_id = None
                                    st.rerun()
                            else:
                                if st.button("🗑️ 撤回刪除", key=f"self_del_{m_id}", use_container_width=True):
                                    st.session_state.confirm_del_id = f"self_{m_id}"
                                    st.rerun()
                else:
                    st.caption("您目前尚無任何已送出的回報紀錄。")
            else:
                st.caption("目前尚無任何已送出的回報紀錄。")

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
        st.title("👥 部門帳號與權限管理中心")
        users_df = load_users_df()
        
        st.subheader("📋 目前已註冊帳號總覽")
        if not users_df.empty:
            display_users = users_df.copy()
            display_users["目前權限"] = display_users["系統權限"].apply(lambda r: "👑 管理員 (admin)" if r == "admin" else "👤 一般同仁 (staff)")
            st.dataframe(display_users[["登入帳號", "真實姓名", "綁定信箱", "目前權限", "建立時間"]], use_container_width=True)
        else:
            st.info("目前尚無帳號資料。")

        st.divider()
        col_m1, col_m2 = st.columns(2)
        
        with col_m1:
            st.subheader("🛡️ 調整現有帳號權限 / 重設密碼")
            if not users_df.empty:
                user_options = [f"{r['真實姓名']} (帳號: {r['登入帳號']} | 權限: {r['系統權限']})" for _, r in users_df.iterrows()]
                selected_opt = st.selectbox("請選擇要操作的帳號", user_options)
                target_idx = user_options.index(selected_opt)
                target_row = users_df.iloc[target_idx]
                target_uname = target_row["登入帳號"]
                
                st.markdown("##### 1️⃣ 變更系統權限")
                new_role = st.radio(
                    "選擇賦予該帳號的權限等級：",
                    ["staff", "admin"],
                    index=0 if target_row["系統權限"] == "staff" else 1,
                    format_func=lambda x: "👤 一般同仁 (僅可填寫每日回報)" if x == "staff" else "👑 管理員 (可查看數據後台 & 管理所有帳號)",
                    horizontal=True
                )
                if st.button("💾 儲存權限變更", type="primary"):
                    if target_uname == "admin" and new_role != "admin":
                        st.error("❌ 系統保護機制：預設最高管理員 `admin` 不可降級為一般同仁！")
                    else:
                        sheet_row = find_user_row_idx(ws_users, target_uname)
                        if sheet_row:
                            ws_users.update_cell(sheet_row, 5, new_role)
                            clear_all_caches()
                            if target_uname == user["username"]:
                                st.session_state.user_info["role"] = new_role
                            st.success(f"✅ 已將 **{target_row['真實姓名']}** 的權限更新為：`{new_role}`！")
                            st.rerun()
                            
                st.markdown("---")
                st.markdown("##### 2️⃣ 強制設定新密碼 (免信箱驗證)")
                force_pass = st.text_input("輸入要幫該同仁設定的新密碼", placeholder="例如：gold8888", key="f_pass")
                if st.button("🔑 直接更新該帳號密碼"):
                    if not force_pass.strip():
                        st.warning("⚠️ 請先輸入新密碼！")
                    else:
                        sheet_row = find_user_row_idx(ws_users, target_uname)
                        if sheet_row:
                            ws_users.update_cell(sheet_row, 2, hash_password(force_pass.strip()))
                            clear_all_caches()
                            st.success(f"✅ 已成功將 **{target_row['真實姓名']}** 的密碼重設為：`{force_pass.strip()}`")

                st.markdown("---")
                st.markdown("##### 3️⃣ 移除離職或誤建帳號")
                confirm_del = st.checkbox(f"我確認要永久刪除「{target_row['真實姓名']} ({target_uname})」的帳號")
                if st.button("🗑️ 刪除此帳號", disabled=not confirm_del):
                    if target_uname in ["admin", user["username"]]:
                        st.error("❌ 無法刪除最高管理員 `admin` 或您目前正在登入中的帳號！")
                    else:
                        sheet_row = find_user_row_idx(ws_users, target_uname)
                        if sheet_row:
                            ws_users.delete_rows(sheet_row)
                            clear_all_caches()
                            st.success(f"✅ 已刪除帳號：{target_row['真實姓名']} ({target_uname})")
                            st.rerun()

        with col_m2:
            st.subheader("➕ 由後台直接幫同仁建立帳號")
            st.caption("您可以先幫同仁建好帳號與預設密碼，同仁登入後可自行修改密碼。")
            adm_name = st.text_input("同仁真實姓名 (將綁定回報單)", key="a_name")
            adm_email = st.text_input("同仁常用信箱 (供未來忘記密碼收信用)", key="a_email")
            adm_user = st.text_input("設定登入帳號", key="a_user")
            adm_pass = st.text_input("設定預設密碼", value="gold8888", key="a_pass")
            adm_role = st.selectbox(
                "初始權限設定",
                ["staff", "admin"],
                format_func=lambda x: "👤 一般同仁 (staff)" if x == "staff" else "👑 管理員 (admin)"
            )
            if st.button("➕ 立即建立帳號", type="primary"):
                if not all([adm_name.strip(), adm_email.strip(), adm_user.strip(), adm_pass.strip()]):
                    st.warning("⚠️ 請填寫所有欄位！")
                elif not users_df.empty and adm_user.strip() in users_df["登入帳號"].values:
                    st.error("❌ 此登入帳號已經存在，請換一個名稱！")
                else:
                    ws_users.append_row([
                        adm_user.strip(),
                        hash_password(adm_pass.strip()),
                        adm_name.strip(),
                        adm_email.strip(),
                        adm_role,
                        get_tw_now().strftime("%Y-%m-%d %H:%M:%S")
                    ])
                    clear_all_caches()
                    st.success(f"✅ 已成功建立 **{adm_name}** 的帳號！預設密碼為：`{adm_pass.strip()}`")
                    st.rerun()

    # ---------------- 功能 D：主管後台數據中心 ----------------
    elif choice == "📊 主管後台數據中心" and user["role"] == "admin":
        st.title("📊 黃金部門 - 主管後台數據中心")
        
        df = load_reports_df()
        if df.empty:
            st.info("目前尚無任何同仁提交回報資料。您可以先點選左側「📝 每日工作回報」測試送出一筆！")
        else:
            for col in ["包貨件數", "包貨總分鐘", "單件包貨耗時(分)", "寄貨件數", "客服總分鐘", "直播總分鐘"]:
                if col not in df.columns:
                    df[col] = 0
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                
            df["客服總分鐘"] = df.apply(
                lambda r: r["客服總分鐘"] if r["客服總分鐘"] > 0 else calc_minutes_str(r.get("客服開始", "無"), r.get("客服結束", "無")), axis=1
            )
            df["直播總分鐘"] = df.apply(
                lambda r: r["直播總分鐘"] if r["直播總分鐘"] > 0 else calc_minutes_str(r.get("直播開始", "無"), r.get("直播結束", "無")), axis=1
            )
            
            df["月份"] = df["回報日期"].astype(str).str.slice(0, 7)
            df["季度"] = df["回報日期"].apply(get_quarter_label)
            
            # 1. 頂部統計區篩選器
            st.markdown("### 🔍 統計看板與圖表篩選器")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                filter_mode = st.radio("統計週期維度", ["📅 依月份檢視", "📆 依季度檢視 (Q1~Q4)"], horizontal=True)
            with col_f2:
                if filter_mode == "📅 依月份檢視":
                    periods = ["全部月份"] + sorted(df["月份"].unique().tolist(), reverse=True)
                    sel_period = st.selectbox("選擇統計月份", periods)
                else:
                    periods = ["全部季度"] + sorted(df["季度"].unique().tolist(), reverse=True)
                    sel_period = st.selectbox("選擇統計季度", periods)
            with col_f3:
                names = ["全部同仁"] + sorted(df["姓名"].unique().tolist())
                sel_name = st.selectbox("👤 選擇統計同仁", names)
                
            filtered_df = df.copy()
            if filter_mode == "📅 依月份檢視" and sel_period != "全部月份":
                filtered_df = filtered_df[filtered_df["月份"] == sel_period]
            elif filter_mode != "📅 依月份檢視" and sel_period != "全部季度":
                filtered_df = filtered_df[filtered_df["季度"] == sel_period]
                
            if sel_name != "全部同仁":
                filtered_df = filtered_df[filtered_df["姓名"] == sel_name]

            # 2. 頂部關鍵指標大字卡
            total_packs = int(filtered_df["包貨件數"].sum())
            total_pack_min = float(filtered_df["包貨總分鐘"].sum())
            avg_pack_speed = round(total_pack_min / total_packs, 2) if total_packs > 0 else 0
            total_ships = int(filtered_df["寄貨件數"].sum())
            total_cs_hr = round(filtered_df["客服總分鐘"].sum() / 60.0, 1)
            total_live_hr = round(filtered_df["直播總分鐘"].sum() / 60.0, 1)

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("📦 區間總包貨件數", f"{total_packs} 件")
            k2.metric("⚡ 平均單件包貨耗時", f"{avg_pack_speed} 分鐘/件")
            k3.metric("🚚 區間總寄貨件數", f"{total_ships} 件")
            k4.metric("💬 客服累計總時數", f"{total_cs_hr} 小時")
            k5.metric("🎥 直播累計總時數", f"{total_live_hr} 小時")

            st.divider()

            # 3. 各同仁當月/當季個別統計總表
            st.subheader(f"📈 各同仁個別統計總表 ({sel_period} - {sel_name})")
            if not filtered_df.empty:
                summary_df = filtered_df.groupby("姓名").agg(
                    出勤回報天數=("回報單號", "count"),
                    總包貨件數=("包貨件數", "sum"),
                    總包貨耗時_分=("包貨總分鐘", "sum"),
                    總寄貨件數=("寄貨件數", "sum"),
                    客服總時數_小時=("客服總分鐘", lambda x: round(x.sum() / 60.0, 2)),
                    直播總時數_小時=("直播總分鐘", lambda x: round(x.sum() / 60.0, 2)),
                    客訴發生次數=("有無客訴", lambda x: (x == "有").sum())
                ).reset_index()
                
                summary_df["平均每件包貨耗時(分/件)"] = summary_df.apply(
                    lambda r: round(r["總包貨耗時_分"] / r["總包貨件數"], 2) if r["總包貨件數"] > 0 else 0, axis=1
                )
                summary_df = summary_df[[
                    "姓名", "出勤回報天數", "總包貨件數", "總包貨耗時_分", "平均每件包貨耗時(分/件)",
                    "總寄貨件數", "客服總時數_小時", "直播總時數_小時", "客訴發生次數"
                ]]
                st.dataframe(summary_df, use_container_width=True)

                # 4. 包貨效率折線圖專區（橫軸：天 / 縱軸：效率）
                st.divider()
                st.subheader("📈 每日包貨效率折線圖（橫軸：天 ／ 縱軸：單件包貨耗時分鐘）")
                
                daily_pack_df = filtered_df[filtered_df["包貨件數"] > 0].copy()
                if not daily_pack_df.empty:
                    daily_grouped = daily_pack_df.groupby(["回報日期", "姓名"], as_index=False).agg(
                        當日包貨件數=("包貨件數", "sum"),
                        當日包貨總分鐘=("包貨總分鐘", "sum")
                    )
                    daily_grouped["單件包貨耗時(分/件)"] = (
                        daily_grouped["當日包貨總分鐘"] / daily_grouped["當日包貨件數"]
                    ).round(2)
                    daily_grouped = daily_grouped.sort_values("回報日期")

                    st.markdown("#### 👥 各同仁包貨效率比較折線圖")
                    fig_compare = px.line(
                        daily_grouped,
                        x="回報日期",
                        y="單件包貨耗時(分/件)",
                        color="姓名",
                        markers=True,
                        text="單件包貨耗時(分/件)",
                        hover_data=["當日包貨件數", "當日包貨總分鐘"],
                        labels={
                            "回報日期": "日期 (天)",
                            "單件包貨耗時(分/件)": "效率：單件包貨耗時 (分/件)"
                        }
                    )
                    fig_compare.update_xaxes(type="category")
                    fig_compare.update_traces(textposition="top center")
                    st.plotly_chart(fig_compare, use_container_width=True)

                    st.markdown("#### 👤 單獨同仁包貨效率折線圖")
                    staff_list = sorted(daily_grouped["姓名"].unique().tolist())
                    col_s1, col_s2 = st.columns([1, 2])
                    with col_s1:
                        target_staff = st.selectbox(
                            "請選擇要單獨檢視折線圖的同仁：",
                            ["📋 一次展開所有同仁各自的獨立折線圖"] + staff_list
                        )
                    
                    if target_staff == "📋 一次展開所有同仁各自的獨立折線圖":
                        cols_per_row = 2
                        for i in range(0, len(staff_list), cols_per_row):
                            c_list = st.columns(cols_per_row)
                            for j in range(cols_per_row):
                                if i + j < len(staff_list):
                                    st_name = staff_list[i + j]
                                    single_df = daily_grouped[daily_grouped["姓名"] == st_name]
                                    fig_single = px.line(
                                        single_df,
                                        x="回報日期",
                                        y="單件包貨耗時(分/件)",
                                        markers=True,
                                        text="單件包貨耗時(分/件)",
                                        hover_data=["當日包貨件數", "當日包貨總分鐘"],
                                        title=f"👤 {st_name} - 個人每日包貨效率折線圖",
                                        labels={
                                            "回報日期": "日期 (天)",
                                            "單件包貨耗時(分/件)": "效率：單件耗時 (分/件)"
                                        }
                                    )
                                    fig_single.update_xaxes(type="category")
                                    fig_single.update_traces(line_color="#00B4D8", textposition="top center")
                                    c_list[j].plotly_chart(fig_single, use_container_width=True)
                    else:
                        single_df = daily_grouped[daily_grouped["姓名"] == target_staff]
                        fig_single = px.line(
                            single_df,
                            x="回報日期",
                            y="單件包貨耗時(分/件)",
                            markers=True,
                            text="單件包貨耗時(分/件)",
                            hover_data=["當日包貨件數", "當日包貨總分鐘"],
                            title=f"👤 {target_staff} - 個人每日包貨效率折線圖",
                            labels={
                                "回報日期": "日期 (天)",
                                "單件包貨耗時(分/件)": "效率：單件耗時 (分/件)"
                            }
                        )
                        fig_single.update_xaxes(type="category")
                        fig_single.update_traces(line_color="#00B4D8", textposition="top center")
                        st.plotly_chart(fig_single, use_container_width=True)
                else:
                    st.caption("目前篩選區間內尚無包貨數據可繪製效率折線圖。")

            st.divider()

            # 5. 匯出美化版分頁 Excel 報表
            st.subheader("📥 匯出美化版分頁 Excel 報表")
            col_dl1, col_dl2 = st.columns([2, 1])
            
            df_overview_exp = pd.DataFrame({
                "回報日期": filtered_df["回報日期"],
                "姓名": filtered_df["姓名"],
                "上下班時間": filtered_df["上班時間"] + " ~ " + filtered_df["下班時間"],
                "包貨時段": filtered_df["包貨開始"] + " ~ " + filtered_df["包貨結束"],
                "包貨件數": filtered_df["包貨件數"],
                "包貨總分鐘": filtered_df["包貨總分鐘"],
                "單件耗時(分)": filtered_df["單件包貨耗時(分)"],
                "寄貨件數": filtered_df["寄貨件數"],
                "客服時段": filtered_df["客服開始"] + " ~ " + filtered_df["客服結束"],
                "客服時數(小時)": (filtered_df["客服總分鐘"] / 60.0).round(2),
                "直播時段": filtered_df["直播開始"] + " ~ " + filtered_df["直播結束"],
                "直播時數(小時)": (filtered_df["直播總分鐘"] / 60.0).round(2),
                "有無客訴": filtered_df["有無客訴"]
            })
            
            df_tasks_exp = pd.DataFrame({
                "回報日期": filtered_df["回報日期"],
                "姓名": filtered_df["姓名"],
                "今日處理事項明細": filtered_df["今日處理事項"],
                "提案改善": filtered_df["提案改善"],
                "對其他同仁想法": filtered_df["對同仁想法"]
            })
            
            comp_only = filtered_df[(filtered_df["有無客訴"] == "有") | (filtered_df["有無8D"] == "是")]
            df_comp_exp = pd.DataFrame({
                "回報日期": comp_only["回報日期"],
                "姓名": comp_only["姓名"],
                "1.客人(Who)": comp_only["客訴_Who"],
                "2.管道(Where)": comp_only["客訴_Where"],
                "3.訂單與日期(When)": comp_only["客訴_When"],
                "4.問題詳情(What)": comp_only["客訴_What"],
                "5.原因分析(Why)": comp_only["客訴_Why"],
                "6.解決方案(How)": comp_only["客訴_How"],
                "有無8D": comp_only["有無8D"],
                "8D檔名": comp_only["8D檔名"]
            })

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                if not filtered_df.empty:
                    summary_df.to_excel(writer, index=False, sheet_name="1_人員統計彙整")
                df_overview_exp.to_excel(writer, index=False, sheet_name="2_每日作業總覽")
                df_tasks_exp.to_excel(writer, index=False, sheet_name="3_事項與提案明細")
                df_comp_exp.to_excel(writer, index=False, sheet_name="4_客訴異常紀錄")
                style_excel_File(writer)
            
            safe_period_str = sel_period if ("全部" not in sel_period) else "ALL"
            with col_dl1:
                st.download_button(
                    label="📊 點擊下載【美化分頁版】Excel 報表 (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"Gold_Report_{safe_period_str}_{get_tw_today().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col_dl2:
                if st.button("✨ 手動重新美化並同步 Google 試算表", use_container_width=True):
                    sync_pretty_sheets(ws_reports, ws_compact, ws_comp, apply_style=True)
                    clear_all_caches()
                    st.success("✅ 已將線上 Google 試算表重新排版上色並同步最新資料！")

            st.divider()

            # 6. 每日回報詳細內容卡片檢視區（含管理員單筆刪除二次確認防呆機制）
            st.subheader("📋 每日回報詳細內容檢視與刪除管理")
            st.caption("可於下方直接選擇想查看的月份、同仁或特定日期（預設自動與上方篩選器連動，亦可直接在此獨立切換）：")
            
            det_c1, det_c2, det_c3, det_c4 = st.columns(4)
            with det_c1:
                all_months_det = ["全部月份"] + sorted(df["月份"].unique().tolist(), reverse=True)
                default_m_idx = all_months_det.index(sel_period) if (filter_mode == "📅 依月份檢視" and sel_period in all_months_det) else 0
                det_month = st.selectbox("📅 選擇檢視月份", all_months_det, index=default_m_idx, key="det_m_sel")
            with det_c2:
                all_names_det = ["全部同仁"] + sorted(df["姓名"].unique().tolist())
                default_n_idx = all_names_det.index(sel_name) if sel_name in all_names_det else 0
                det_name = st.selectbox("👤 選擇檢視同仁", all_names_det, index=default_n_idx, key="det_n_sel")
                
            detail_df = df.copy()
            if det_month != "全部月份":
                detail_df = detail_df[detail_df["月份"] == det_month]
            if det_name != "全部同仁":
                detail_df = detail_df[detail_df["姓名"] == det_name]
                
            with det_c3:
                avail_dates = ["全部日期"] + sorted(detail_df["回報日期"].unique().tolist(), reverse=True)
                det_date = st.selectbox("📆 選擇特定日期", avail_dates, key="det_d_sel")
            with det_c4:
                det_status = st.selectbox("🚨 客訴 / 異常快篩", ["全部回報", "🚨 僅看有客訴", "📎 僅看有附8D報告"], key="det_s_sel")
                
            if det_date != "全部日期":
                detail_df = detail_df[detail_df["回報日期"] == det_date]
            if det_status == "🚨 僅看有客訴":
                detail_df = detail_df[detail_df["有無客訴"] == "有"]
            elif det_status == "📎 僅看有附8D報告":
                detail_df = detail_df[detail_df["有無8D"] == "是"]

            if detail_df.empty:
                st.info("🔍 此篩選條件下無符合的回報紀錄。")
            else:
                sorted_df = detail_df.sort_values(by=["回報日期", "送出時間"], ascending=[False, False]).reset_index(drop=True)
                st.caption(f"共找到 **{len(sorted_df)}** 筆符合條件的回報單（已按日期由新到舊排列）：")
                
                for idx, row in sorted_df.iterrows():
                    r_id = row.get("回報單號", f"idx_{idx}")
                    alert_badge = " 🚨【有客訴】" if row.get("有無客訴") == "有" else ""
                    badge_8d = " 📎【附8D報告】" if row.get("有無8D") == "是" else ""
                    exp_title = f"📅 {row.get('回報日期', '')} ｜ 👤 {row.get('姓名', '')} ｜ 🕒 上下班：{row.get('上班時間', '')}~{row.get('下班時間', '')} ｜ 📦 包貨：{int(row.get('包貨件數', 0))}件 / 🚚 寄貨：{int(row.get('寄貨件數', 0))}件{alert_badge}{badge_8d}"
                    
                    with st.expander(exp_title, expanded=(idx == 0)):
                        c1, c2, c3, c4 = st.columns(4)
                        c1.markdown(f"**🕒 出勤時間**\n\n`{row.get('上班時間', '')} ~ {row.get('下班時間', '')}`")
                        if int(row.get("包貨件數", 0)) > 0:
                            c2.markdown(f"**📦 包貨紀錄**\n\n`{row.get('包貨開始', '')} ~ {row.get('包貨結束', '')}`\n共 **{int(row.get('包貨件數', 0))}** 件 (`{row.get('單件包貨耗時(分)', 0)}` 分/件)")
                        else:
                            c2.markdown("**📦 包貨紀錄**\n\n`無`")
                            
                        cs_hr_row = round(float(row.get("客服總分鐘", 0)) / 60.0, 2)
                        c3.markdown(f"**🚚 寄貨 / 💬 客服**\n\n寄貨：**{int(row.get('寄貨件數', 0))}** 件\n客服：`{row.get('客服開始', '無')} ~ {row.get('客服結束', '無')}` (**{cs_hr_row}** 小時)")
                        
                        lv_hr_row = round(float(row.get("直播總分鐘", 0)) / 60.0, 2)
                        c4.markdown(f"**🎥 直播紀錄**\n\n時段：`{row.get('直播開始', '無')} ~ {row.get('直播結束', '無')}`\n時長：**{lv_hr_row}** 小時")
                        
                        st.markdown("---")
                        st.markdown("#### 📝 今日處理事項明細")
                        tasks_text = str(row.get("今日處理事項", "")).strip()
                        if tasks_text:
                            for line in tasks_text.split("\n"):
                                st.markdown(f"- {line}")
                        else:
                            st.caption("無填寫內容")
                            
                        col_p1, col_p2 = st.columns(2)
                        with col_p1:
                            st.markdown("#### 💡 提案改善")
                            prop_text = str(row.get("提案改善", "")).strip()
                            if prop_text:
                                st.info(prop_text)
                            else:
                                st.caption("無提案")
                        with col_p2:
                            st.markdown("#### 🤝 針對其他同仁想法")
                            peer_text = str(row.get("對同仁想法", "")).strip()
                            if peer_text:
                                st.warning(peer_text)
                            else:
                                st.caption("無填寫")
                                
                        if row.get("有無客訴") == "有":
                            st.markdown("---")
                            st.error("🚨 **客訴案件 5W1H 詳細報告**")
                            st.markdown(f"- **1. 【Who】客人本名 + TK名：** {row.get('客訴_Who', '')}")
                            st.markdown(f"- **2. 【Where】發生管道：** {row.get('客訴_Where', '')}")
                            st.markdown(f"- **3. 【When】訂單與客訴日期：** {row.get('客訴_When', '')}")
                            st.markdown(f"- **4. 【What】發生什麼事 / 商品問題：**\n  > {row.get('客訴_What', '')}")
                            st.markdown(f"- **5. 【Why】客訴背後原因：**\n  > {row.get('客訴_Why', '')}")
                            st.markdown(f"- **6. 【How】解決方案與未來改善：**\n  > {row.get('客訴_How', '')}")
                            
                        if row.get("有無8D") == "是" and row.get("8D檔案編碼"):
                            st.markdown("---")
                            b64_bytes = base64.b64decode(row["8D檔案編碼"])
                            st.download_button(
                                label=f"📎 點擊下載 8D 報告檔案 ({row.get('8D檔名', '8D_report.xlsx')})",
                                data=b64_bytes,
                                file_name=f"8D_{r_id}_{row.get('8D檔名', 'report.xlsx')}",
                                key=f"dl_card_{r_id}"
                            )
                            
                        st.markdown("---")
                        foot_c1, foot_c2 = st.columns([3, 2])
                        foot_c1.caption(f"系統回報單號：{r_id} ｜ 實際送出時間 (台灣時間)：{row.get('送出時間', '')}")
                        
                        with foot_c2:
                            if st.session_state.confirm_del_id == f"adm_{r_id}":
                                st.warning(f"⚠️ 確定要永久刪除「{row.get('回報日期', '')} - {row.get('姓名', '')}」的回報嗎？（刪除後無法復原）")
                                dc1, dc2 = st.columns(2)
                                if dc1.button("✅ 確定永久刪除", key=f"yes_adm_{r_id}", type="primary", use_container_width=True):
                                    if delete_report_by_id(ws_reports, ws_compact, ws_comp, r_id):
                                        st.session_state.confirm_del_id = None
                                        st.success(f"✅ 已徹底刪除 {row.get('回報日期', '')} - {row.get('姓名', '')} 的回報紀錄！")
                                        st.rerun()
                                if dc2.button("❌ 取消", key=f"no_adm_{r_id}", use_container_width=True):
                                    st.session_state.confirm_del_id = None
                                    st.rerun()
                            else:
                                if st.button("🗑️ 刪除此筆回報紀錄", key=f"adm_del_{r_id}", use_container_width=True):
                                    st.session_state.confirm_del_id = f"adm_{r_id}"
                                    st.rerun()
