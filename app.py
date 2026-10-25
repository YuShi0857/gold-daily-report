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
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Alignment

# 頁面基本設定
st.set_page_config(page_title="黃金部門工作回報系統", layout="wide", page_icon="✨")

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
    "今日處理事項", "提案改善", "對同仁想法", "有無客訴"
]

COMPLAINT_HEADERS = [
    "回報日期", "填表同仁", "1.客人(Who)", "2.管道(Where)", 
    "3.訂單與客訴日(When)", "4.問題詳情(What)", "5.原因分析(Why)", 
    "6.解決與改善(How)", "有無附8D報告", "8D檔案名稱"
]

# ================= 1. 資料庫連線與自動中文化升級 =================
@st.cache_resource
def init_connection():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    client = gspread.authorize(creds)
    return client.open("黃金部門工作回報系統")

def get_worksheets():
    sh = init_connection()
    ws_users = sh.worksheet("users")
    ws_reports = sh.worksheet("reports")
    
    # 1. 初始化 users 分頁 (全中文標題)
    u_headers = ws_users.row_values(1)
    if not u_headers:
        ws_users.append_row(["登入帳號", "密碼雜湊", "真實姓名", "綁定信箱", "系統權限", "建立時間"])
        admin_pass = hashlib.sha256("admin888".encode()).hexdigest()
        ws_users.append_row(["admin", admin_pass, "系統管理員", st.secrets["smtp"]["sender_email"], "admin", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    elif u_headers[0] == "username":
        ws_users.update("A1:F1", [["登入帳號", "密碼雜湊", "真實姓名", "綁定信箱", "系統權限", "建立時間"]])

    # 2. 初始化或升級 reports 主資料庫分頁 (全中文標題 + 擴充客服/直播分鐘數)
    r_vals = ws_reports.get_all_values()
    if not r_vals:
        ws_reports.append_row(REPORT_HEADERS)
    elif r_vals[0][0] == "report_id" or len(r_vals[0]) < len(REPORT_HEADERS):
        # 自動將舊版英文表格無痛升級為新版中文結構
        new_rows = [REPORT_HEADERS]
        for r in r_vals[1:]:
            if not any(str(c).strip() for c in r):
                continue
            r_pad = r + [""] * max(0, 32 - len(r))
            if len(r) <= 30 and r_vals[0][0] == "report_id":
                cs_m = calc_minutes_str(r_pad[12], r_pad[13])
                lv_m = calc_minutes_str(r_pad[15], r_pad[16])
                upgraded = r_pad[:14] + [cs_m, r_pad[14], r_pad[15], r_pad[16], lv_m] + r_pad[17:30]
                new_rows.append(upgraded)
            else:
                new_rows.append(r_pad[:len(REPORT_HEADERS)])
        ws_reports.clear()
        ws_reports.update(new_rows)

    # 3. 自動建立「每日作業總覽」與「客訴與異常紀錄」兩個精簡中文分頁（解決試算表太長要一直往右滑的問題）
    existing_titles = [w.title for w in sh.worksheets()]
    if "每日作業總覽" not in existing_titles:
        ws_compact = sh.add_worksheet(title="每日作業總覽", rows="1000", cols="12")
        ws_compact.append_row(COMPACT_HEADERS)
    else:
        ws_compact = sh.worksheet("每日作業總覽")
        
    if "客訴與異常紀錄" not in existing_titles:
        ws_comp = sh.add_worksheet(title="客訴與異常紀錄", rows="500", cols="11")
        ws_comp.append_row(COMPLAINT_HEADERS)
    else:
        ws_comp = sh.worksheet("客訴與異常紀錄")

    return ws_users, ws_reports, ws_compact, ws_comp

# ================= 2. 輔助與安全函數 =================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def calc_minutes(t_start, t_end):
    dt1 = datetime.datetime.combine(datetime.date.today(), t_start)
    dt2 = datetime.datetime.combine(datetime.date.today(), t_end)
    if dt2 < dt1:
        dt2 += datetime.timedelta(days=1)
    return round((dt2 - dt1).total_seconds() / 60.0, 2)

def calc_minutes_str(s1, s2):
    try:
        if not s1 or not s2 or s1 == "無" or s2 == "無" or s1 == "00:00" and s2 == "00:00":
            return 0
        t1 = datetime.datetime.strptime(s1.strip(), "%H:%M").time()
        t2 = datetime.datetime.strptime(s2.strip(), "%H:%M").time()
        return calc_minutes(t1, t2)
    except:
        return 0

def safe_get_df(ws):
    all_vals = ws.get_all_values()
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

def find_user_row_idx(ws_users, target_username):
    col_vals = ws_users.col_values(1)
    for idx, val in enumerate(col_vals):
        if idx > 0 and val.strip() == str(target_username).strip():
            return idx + 1
    return None

def get_quarter_label(date_str):
    try:
        dt = pd.to_datetime(date_str)
        q = (dt.month - 1) // 3 + 1
        return f"{dt.year}-Q{q}"
    except:
        return "未知季度"

def style_excel_File(writer):
    # 美化匯出的 Excel：加上深藍底白字標題、自動欄寬、文字自動換行，不用一直往右滑
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

# ================= 3. 狀態管理初始化 =================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_info = {}
if "task_rows" not in st.session_state:
    st.session_state.task_rows = 1

try:
    ws_users, ws_reports, ws_compact, ws_comp = get_worksheets()
except Exception as e:
    st.error(f"⚠️ 無法連線至 Google 試算表，請確認設定：{e}")
    st.stop()

# ================= 4. 未登入畫面（登入 / 註冊 / 忘記密碼） =================
if not st.session_state.logged_in:
    st.title("✨ 黃金部門工作回報系統")
    tab1, tab2, tab3 = st.tabs(["🔑 登入系統", "📝 新人註冊", "❓ 忘記密碼"])
    
    with tab1:
        st.subheader("員工/管理員登入")
        login_user = st.text_input("帳號", key="l_user")
        login_pass = st.text_input("密碼", type="password", key="l_pass")
        if st.button("登入", type="primary"):
            users_df = safe_get_df(ws_users)
            if not users_df.empty and login_user in users_df["登入帳號"].values:
                user_row = users_df[users_df["登入帳號"] == login_user].iloc[0]
                if user_row["密碼雜湊"] == hash_password(login_pass):
                    st.session_state.logged_in = True
                    st.session_state.user_info = {
                        "username": user_row["登入帳號"],
                        "password_hash": user_row["密碼雜湊"],
                        "name": user_row["真實姓名"],
                        "email": user_row["綁定信箱"],
                        "role": user_row["系統權限"]
                    }
                    st.rerun()
                else:
                    st.error("❌ 密碼錯誤！")
            else:
                st.error("❌ 找不到此帳號，請先註冊！")

    with tab2:
        st.subheader("註冊新帳號")
        reg_name = st.text_input("真實姓名 (將自動綁定未來的回報單)")
        reg_email = st.text_input("常用信箱 (忘記密碼寄信用，請務必填寫正確)")
        reg_user = st.text_input("設定登入帳號")
        reg_pass1 = st.text_input("設定密碼", type="password")
        reg_pass2 = st.text_input("再次確認密碼", type="password")
        
        if st.button("註冊帳號"):
            if not all([reg_name, reg_email, reg_user, reg_pass1, reg_pass2]):
                st.warning("⚠️ 所有欄位皆為必填！")
            elif reg_pass1 != reg_pass2:
                st.error("❌ 兩次輸入的密碼不一致，請重新確認！")
            else:
                users_df = safe_get_df(ws_users)
                if not users_df.empty and reg_user in users_df["登入帳號"].values:
                    st.error("❌ 此帳號已被註冊，請換一個！")
                else:
                    ws_users.append_row([
                        reg_user.strip(), hash_password(reg_pass1), reg_name.strip(), reg_email.strip(), "staff", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ])
                    st.success("✅ 註冊成功！請切換至「登入系統」分頁進行登入。")

    with tab3:
        st.subheader("忘記密碼 / 重新寄送臨時密碼")
        fp_user = st.text_input("請輸入您的帳號")
        fp_email = st.text_input("請輸入註冊時綁定的信箱")
        if st.button("寄送重設密碼信"):
            users_df = safe_get_df(ws_users)
            if not users_df.empty:
                match = users_df[(users_df["登入帳號"] == fp_user) & (users_df["綁定信箱"] == fp_email)]
                if not match.empty:
                    new_temp_pass = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
                    row_idx = find_user_row_idx(ws_users, fp_user)
                    if row_idx:
                        ws_users.update_cell(row_idx, 2, hash_password(new_temp_pass))
                        if send_reset_email(fp_email, new_temp_pass):
                            st.success("✅ 臨時新密碼已寄出！請至信箱收信並重新登入。")
                else:
                    st.error("❌ 帳號與綁定信箱不符合！")
            else:
                st.error("❌ 系統尚無使用者資料！")

# ================= 5. 已登入畫面 =================
else:
    user = st.session_state.user_info
    st.sidebar.title(f"👋 您好，{user['name']}")
    st.sidebar.caption(f"身份：{'👑 管理員' if user['role'] == 'admin' else '👤 部門同仁'}")
    
    menu_options = ["📝 每日工作回報", "🔒 修改密碼"]
    if user["role"] == "admin":
        menu_options = ["📊 主管後台數據中心", "👥 帳號與權限管理", "📝 每日工作回報", "🔒 修改密碼"]
        
    choice = st.sidebar.radio("功能選單", menu_options)
    
    if st.sidebar.button("🚪 登出系統"):
        st.session_state.logged_in = False
        st.session_state.user_info = {}
        st.rerun()

    # ---------------- 功能 A：每日工作回報 ----------------
    if choice == "📝 每日工作回報":
        st.title("📝 黃金部門 - 每日工作回報單")
        st.info(f"📌 填表人：**{user['name']}** （系統已自動綁定您的姓名）")
        
        report_date = st.date_input("📅 回報歸屬日期", datetime.date.today())
        
        st.markdown("### 一、 基礎出勤與作業數據")
        col1, col2 = st.columns(2)
        with col1:
            work_start = st.time_input("🕒 上班時間", datetime.time(9, 0), step=60)
        with col2:
            work_end = st.time_input("🕕 下班時間", datetime.time(18, 0), step=60)
            
        st.markdown("#### 📦 包貨與寄貨紀錄（有執行才需勾選填寫）")
        col_chk1, col_chk2 = st.columns(2)
        
        pack_start_str, pack_end_str = "無", "無"
        pack_count, pack_duration, min_per_pack = 0, 0, 0
        
        with col_chk1:
            has_pack = st.checkbox("📦 今日有包貨")
            if has_pack:
                c_p1, c_p2, c_p3 = st.columns(3)
                with c_p1:
                    pack_start = st.time_input("包貨開始時間", datetime.time(10, 0), step=60)
                with c_p2:
                    pack_end = st.time_input("包貨結束時間", datetime.time(12, 0), step=60)
                with c_p3:
                    pack_count = st.number_input("個人包貨件數", min_value=1, value=1, step=1)
                pack_start_str = pack_start.strftime("%H:%M")
                pack_end_str = pack_end.strftime("%H:%M")
                pack_duration = calc_minutes(pack_start, pack_end)
                min_per_pack = round(pack_duration / pack_count, 2) if pack_count > 0 else 0

        ship_count = 0
        with col_chk2:
            has_ship = st.checkbox("🚚 今日有寄貨")
            if has_ship:
                ship_count = st.number_input("今日寄貨件數", min_value=1, value=1, step=1)
            
        st.markdown("#### 💬 客服與直播紀錄（有執行才需勾選填寫）")
        col_chk3, col_chk4 = st.columns(2)
        
        cs_start_str, cs_end_str = "無", "無"
        cs_duration = 0
        with col_chk3:
            has_cs = st.checkbox("💬 今日有客服回覆")
            if has_cs:
                c_l1, c_l2 = st.columns(2)
                with c_l1:
                    cs_start = st.time_input("客服回覆開始", datetime.time(9, 0), step=60)
                with c_l2:
                    cs_end = st.time_input("客服回覆結束", datetime.time(18, 0), step=60)
                cs_start_str = cs_start.strftime("%H:%M")
                cs_end_str = cs_end.strftime("%H:%M")
                cs_duration = calc_minutes(cs_start, cs_end)

        live_start_str, live_end_str = "無", "無"
        live_duration = 0
        with col_chk4:
            has_live = st.checkbox("🎥 今日有直播")
            if has_live:
                c_l3, c_l4 = st.columns(2)
                with c_l3:
                    live_start = st.time_input("今日直播開始", datetime.time(19, 0), step=60)
                with c_l4:
                    live_end = st.time_input("今日直播結束", datetime.time(22, 0), step=60)
                live_start_str = live_start.strftime("%H:%M")
                live_end_str = live_end.strftime("%H:%M")
                live_duration = calc_minutes(live_start, live_end)

        st.divider()
        
        # 動態展開：客訴與 8D
        st.markdown("### 二、 客訴與重大異常回報")
        has_complaint = st.radio("🚨 今日是否有任何客訴？", ["無", "有"], horizontal=True)
        c_who = c_where = c_when = c_what = c_why = c_how = ""
        
        if has_complaint == "有":
            st.warning("⚠️ 請詳細填寫以下客訴 5W1H 資訊：")
            c_who = st.text_input("1. 【Who】客人本名 + TK名 / 帳號名稱")
            c_where = st.selectbox("2. 【Where】客訴發生管道", ["直播上", "LINE客服上", "LINE社群", "其他"])
            c_when = st.text_input("3. 【When】幾月幾號的訂單 及 客訴發生日期")
            c_what = st.text_area("4. 【What】發生什麼事 / 商品的問題 (請詳細打字說明)")
            c_why = st.text_area("5. 【Why】客訴背後原因分析")
            c_how = st.text_area("6. 【How】執行方法 / 解決方案 / 未來如何改善")

        has_8d = st.checkbox("📄 若有重大異常，需填寫並附上 8D 報告檔案")
        file_8d_name, file_8d_b64 = "", ""
        if has_8d:
            uploaded_8d = st.file_uploader("請上傳填妥之 8D 檔案 (支援 Excel, Word, PDF，建議小於 1MB)", type=["xlsx", "xls", "docx", "doc", "pdf"])
            if uploaded_8d is not None:
                file_8d_name = uploaded_8d.name
                file_8d_b64 = base64.b64encode(uploaded_8d.read()).decode("utf-8")

        st.divider()

        # 動態增減：今日處理事項
        st.markdown("### 三、 今日處理事項明細")
        st.caption("請清楚標註幾點幾分到幾點幾分做了哪些事（制式規格，不夠可按下方按鈕新增列）")
        
        col_btn1, col_btn2, _ = st.columns([1, 1, 4])
        with col_btn1:
            if st.button("➕ 新增時段項目"):
                st.session_state.task_rows += 1
        with col_btn2:
            if st.button("➖ 減少最後一列") and st.session_state.task_rows > 1:
                st.session_state.task_rows -= 1

        tasks_list = []
        for i in range(st.session_state.task_rows):
            t_col1, t_col2, t_col3 = st.columns([1.5, 1.5, 5])
            with t_col1:
                ts = st.time_input(f"開始時間 #{i+1}", datetime.time(9, 0), key=f"ts_{i}", step=60)
            with t_col2:
                te = st.time_input(f"結束時間 #{i+1}", datetime.time(10, 0), key=f"te_{i}", step=60)
            with t_col3:
                td = st.text_input(f"處理事項內容 #{i+1}", placeholder="請輸入該時段具體完成的事項...", key=f"td_{i}")
            if td.strip():
                tasks_list.append(f"【{ts.strftime('%H:%M')} - {te.strftime('%H:%M')}】{td.strip()}")

        st.divider()

        # 提案與回饋
        st.markdown("### 四、 提案改善與同仁想法回饋")
        improvement_proposal = st.text_area("💡 提案改善欄位 (針對作業流程、商品、工具有任何優化建議皆可提出)")
        peer_feedback = st.text_area("🤝 針對其他同仁有任何想法或建議 (僅主管後台可見)")

        if st.button("🚀 確認無誤，送出今日回報", type="primary", use_container_width=True):
            if has_complaint == "有" and not all([c_who, c_when, c_what, c_why, c_how]):
                st.error("❌ 您勾選了「有客訴」，請將 Who / When / What / Why / How 所有欄位填寫完整！")
            elif has_8d and not file_8d_name:
                st.error("❌ 您勾選了「需填寫8D」，請務必上傳 8D 檔案！")
            elif not tasks_list:
                st.error("❌ 請至少填寫一條「今日處理事項」！")
            else:
                report_id = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "_" + user["username"]
                # 1. 寫入主資料庫 (全中文標題)
                ws_reports.append_row([
                    report_id,
                    datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    str(report_date),
                    user["username"],
                    user["name"],
                    work_start.strftime("%H:%M"),
                    work_end.strftime("%H:%M"),
                    pack_start_str,
                    pack_end_str,
                    pack_count,
                    pack_duration,
                    min_per_pack,
                    cs_start_str,
                    cs_end_str,
                    cs_duration,
                    ship_count,
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
                
                # 2. 同步寫入「每日作業總覽」精簡分頁（讓你在 Google Sheet 直接看超清爽，不用往右滑！）
                pack_summary = f"{pack_start_str}~{pack_end_str} ({pack_count}件 / {min_per_pack}分)" if pack_count > 0 else "無"
                cs_summary = f"{cs_start_str}~{cs_end_str} ({round(cs_duration/60, 1)}小時)" if cs_duration > 0 else "無"
                live_summary = f"{live_start_str}~{live_end_str} ({round(live_duration/60, 1)}小時)" if live_duration > 0 else "無"
                ws_compact.append_row([
                    str(report_date),
                    user["name"],
                    f"{work_start.strftime('%H:%M')}~{work_end.strftime('%H:%M')}",
                    pack_summary,
                    ship_count,
                    cs_summary,
                    live_summary,
                    "\n".join(tasks_list),
                    improvement_proposal or "無",
                    peer_feedback or "無",
                    has_complaint
                ])
                
                # 3. 若有客訴或 8D，同步寫入「客訴與異常紀錄」專屬分頁
                if has_complaint == "有" or has_8d:
                    ws_comp.append_row([
                        str(report_date), user["name"], c_who, c_where, c_when, c_what, c_why, c_how,
                        "是" if has_8d else "否", file_8d_name
                    ])
                    
                st.success("🎉 今日工作回報已成功送出！辛苦了！")
                st.session_state.task_rows = 1

    # ---------------- 功能 B：修改密碼 ----------------
    elif choice == "🔒 修改密碼":
        st.title("🔒 修改個人密碼")
        old_p = st.text_input("請輸入舊密碼 (或信箱收到的臨時密碼)", type="password")
        new_p1 = st.text_input("請輸入新密碼", type="password")
        new_p2 = st.text_input("請再次確認新密碼", type="password")
        if st.button("更新密碼", type="primary"):
            if new_p1 != new_p2:
                st.error("❌ 兩次新密碼輸入不一致！")
            elif hash_password(old_p) != user["password_hash"]:
                st.error("❌ 舊密碼輸入錯誤！")
            else:
                row_idx = find_user_row_idx(ws_users, user["username"])
                if row_idx:
                    new_hash = hash_password(new_p1)
                    ws_users.update_cell(row_idx, 2, new_hash)
                    st.session_state.user_info["password_hash"] = new_hash
                    st.success("✅ 密碼更新成功！")

    # ---------------- 功能 C：帳號與權限管理 ----------------
    elif choice == "👥 帳號與權限管理" and user["role"] == "admin":
        st.title("👥 部門帳號與權限管理中心")
        users_df = safe_get_df(ws_users)
        
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
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ])
                    st.success(f"✅ 已成功建立 **{adm_name}** 的帳號！預設密碼為：`{adm_pass.strip()}`")
                    st.rerun()

    # ---------------- 功能 D：主管後台數據中心 ----------------
    elif choice == "📊 主管後台數據中心" and user["role"] == "admin":
        st.title("📊 黃金部門 - 主管後台數據中心")
        
        df = safe_get_df(ws_reports)
        if df.empty:
            st.info("目前尚無任何同仁提交回報資料。您可以先點選左側「📝 每日工作回報」測試送出一筆！")
        else:
            # 確保數值與時間計算完整
            for col in ["包貨件數", "包貨總分鐘", "單件包貨耗時(分)", "寄貨件數", "客服總分鐘", "直播總分鐘"]:
                if col not in df.columns:
                    df[col] = 0
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                
            # 若舊資料尚未計算客服/直播分鐘數，即時自動補算
            df["客服總分鐘"] = df.apply(
                lambda r: r["客服總分鐘"] if r["客服總分鐘"] > 0 else calc_minutes_str(r.get("客服開始", "無"), r.get("客服結束", "無")), axis=1
            )
            df["直播總分鐘"] = df.apply(
                lambda r: r["直播總分鐘"] if r["直播總分鐘"] > 0 else calc_minutes_str(r.get("直播開始", "無"), r.get("直播結束", "無")), axis=1
            )
            
            df["月份"] = df["回報日期"].astype(str).str.slice(0, 7)
            df["季度"] = df["回報日期"].apply(get_quarter_label)
            
            # 1. 篩選器（支援按月份、按季度、按同仁篩選）
            st.markdown("### 🔍 第一區：區間與人員篩選器")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                filter_mode = st.radio("統計週期維度", ["📅 依月份檢視", "📆 依季度檢視 (Q1~Q4)"], horizontal=True)
            with col_f2:
                if filter_mode == "📅 依月份檢視":
                    periods = ["全部月份"] + sorted(df["月份"].unique().tolist(), reverse=True)
                    sel_period = st.selectbox("選擇月份", periods)
                else:
                    periods = ["全部季度"] + sorted(df["季度"].unique().tolist(), reverse=True)
                    sel_period = st.selectbox("選擇季度", periods)
            with col_f3:
                names = ["全部同仁"] + sorted(df["姓名"].unique().tolist())
                sel_name = st.selectbox("👤 選擇篩選同仁", names)
                
            filtered_df = df.copy()
            if filter_mode == "📅 依月份檢視" and sel_period != "全部月份":
                filtered_df = filtered_df[filtered_df["月份"] == sel_period]
            elif filter_mode != "📅 依月份檢視" and sel_period != "全部季度":
                filtered_df = filtered_df[filtered_df["季度"] == sel_period]
                
            if sel_name != "全部同仁":
                filtered_df = filtered_df[filtered_df["姓名"] == sel_name]

            # 2. 頂部關鍵指標大字卡 (KPI Cards)
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
                # 調整欄位順序讓閱讀最順暢
                summary_df = summary_df[[
                    "姓名", "出勤回報天數", "總包貨件數", "總包貨耗時_分", "平均每件包貨耗時(分/件)",
                    "總寄貨件數", "客服總時數_小時", "直播總時數_小時", "客訴發生次數"
                ]]
                st.dataframe(summary_df, use_container_width=True)

                # 4. 當月/當季 包貨效率與工作量視覺化分析圖表
                st.subheader("📊 當月 / 當季包貨效率與工作量分析圖")
                chart_col1, chart_col2 = st.columns(2)
                
                with chart_col1:
                    pack_staff_df = summary_df[summary_df["總包貨件數"] > 0]
                    if not pack_staff_df.empty:
                        fig_eff = px.bar(
                            pack_staff_df,
                            x="姓名",
                            y="平均每件包貨耗時(分/件)",
                            color="姓名",
                            text="平均每件包貨耗時(分/件)",
                            title="⏱️ 各同仁平均單件包貨耗時對比 (數值越低代表包越快)"
                        )
                        fig_eff.update_traces(textposition="outside")
                        st.plotly_chart(fig_eff, use_container_width=True)
                    else:
                        st.caption("此區間尚無包貨數據可繪製效率對比圖")

                with chart_col2:
                    daily_pack_df = filtered_df[filtered_df["包貨件數"] > 0].copy()
                    if not daily_pack_df.empty:
                        daily_pack_df = daily_pack_df.sort_values("回報日期")
                        fig_trend = px.line(
                            daily_pack_df,
                            x="回報日期",
                            y="單件包貨耗時(分)",
                            color="姓名",
                            markers=True,
                            hover_data=["包貨件數", "包貨總分鐘"],
                            title="📈 每日包貨效率趨勢圖 (單件耗時變化)"
                        )
                        st.plotly_chart(fig_trend, use_container_width=True)
                    else:
                        st.caption("此區間尚無每日包貨趨勢數據")

                # 第二排圖表：包貨與寄貨件數＆客服直播時數對比
                chart_col3, chart_col4 = st.columns(2)
                with chart_col3:
                    fig_vol = px.bar(
                        summary_df,
                        x="姓名",
                        y=["總包貨件數", "總寄貨件數"],
                        barmode="group",
                        title="📦 各同仁總包貨件數 vs 總寄貨件數"
                    )
                    st.plotly_chart(fig_vol, use_container_width=True)
                with chart_col4:
                    fig_hr = px.bar(
                        summary_df,
                        x="姓名",
                        y=["客服總時數_小時", "直播總時數_小時"],
                        barmode="group",
                        title="💬 各同仁客服總時數 vs 🎥 直播總時數 (小時)"
                    )
                    st.plotly_chart(fig_hr, use_container_width=True)

            st.divider()

            # 5. 匯出美化版分頁 Excel 報表（解決一長條難看的問題）
            st.subheader("📥 匯出美化版分頁 Excel 報表 (自動排版/自動換行/免往右滑)")
            st.caption("匯出的 Excel 已自動幫您拆分為 4 個精簡分頁：【1_人員統計彙整】、【2_每日作業總覽】、【3_事項與提案明細】、【4_客訴異常紀錄】，打開即可直接閱讀或列印成 PDF！")
            
            # 整理精簡版匯出表格
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
            
            st.download_button(
                label=f"📊 點擊下載【美化分頁版】Excel 報表 ({sel_period}_{sel_name}.xlsx)",
                data=output.getvalue(),
                file_name=f"黃金部門工作回報_{sel_period}_{sel_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            st.divider()

            # 6. 每日回報詳細內容卡片檢視區
            st.subheader("📋 每日回報詳細內容檢視（點擊卡片可展開/收合）")
            sorted_df = filtered_df.iloc[::-1].reset_index(drop=True)
            for idx, row in sorted_df.iterrows():
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
                            label=f"📎 點擊下載 8D 報告檔案 ({row.get('8D檔名', '8D_report')})",
                            data=b64_bytes,
                            file_name=row.get("8D檔名", "8D_report"),
                            key=f"dl_card_{row.get('回報單號', idx)}"
                        )
                    st.caption(f"系統回報單號：{row.get('回報單號', '')} ｜ 實際送出時間：{row.get('送出時間', '')}")
