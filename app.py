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

# 頁面基本設定
st.set_page_config(page_title="黃金部門工作回報系統", layout="wide", page_icon="✨")

# ================= 1. 資料庫連線與初始化 =================
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
    
    # 自動檢查並初始化標題列
    if not ws_users.row_values(1):
        ws_users.append_row(["username", "password_hash", "name", "email", "role", "created_at"])
        admin_pass = hashlib.sha256("admin888".encode()).hexdigest()
        ws_users.append_row(["admin", admin_pass, "系統管理員", st.secrets["smtp"]["sender_email"], "admin", str(datetime.datetime.now())])
        
    if not ws_reports.row_values(1):
        ws_reports.append_row([
            "report_id", "submit_time", "report_date", "username", "name",
            "work_start", "work_end", 
            "pack_start", "pack_end", "pack_count", "pack_duration_min", "min_per_pack",
            "cs_start", "cs_end", "ship_count", "live_start", "live_end",
            "has_complaint", "c_who", "c_where", "c_when", "c_what", "c_why", "c_how",
            "has_8d", "file_8d_name", "file_8d_b64",
            "daily_tasks", "improvement_proposal", "peer_feedback"
        ])
    return ws_users, ws_reports

# 防呆讀取函數（徹底解決 GSpreadException 標題列問題）
def safe_get_df(ws):
    all_vals = ws.get_all_values()
    if not all_vals or len(all_vals) <= 1:
        return pd.DataFrame()
    headers = [h.strip() for h in all_vals[0]]
    valid_cols = [i for i, h in enumerate(headers) if h != ""]
    clean_headers = [headers[i] for i in valid_cols]
    data = []
    for r in all_vals[1:]:
        if not any(str(cell).strip() for cell in r):
            continue
        row_padded = r + [""] * max(0, len(headers) - len(r))
        data.append([row_padded[i] for i in valid_cols])
    return pd.DataFrame(data, columns=clean_headers)

# ================= 2. 輔助與安全函數 =================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def calc_minutes(t_start, t_end):
    dt1 = datetime.datetime.combine(datetime.date.today(), t_start)
    dt2 = datetime.datetime.combine(datetime.date.today(), t_end)
    if dt2 < dt1:
        dt2 += datetime.timedelta(days=1)
    return round((dt2 - dt1).total_seconds() / 60.0, 2)

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
    ws_users, ws_reports = get_worksheets()
except Exception as e:
    st.error("⚠️ 無法連線至 Google 試算表，請確認 Streamlit Secrets 設定與試算表共用權限是否正確！")
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
            if not users_df.empty and login_user in users_df["username"].values:
                user_row = users_df[users_df["username"] == login_user].iloc[0]
                if user_row["password_hash"] == hash_password(login_pass):
                    st.session_state.logged_in = True
                    st.session_state.user_info = user_row.to_dict()
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
                if not users_df.empty and reg_user in users_df["username"].values:
                    st.error("❌ 此帳號已被註冊，請換一個！")
                else:
                    ws_users.append_row([
                        reg_user, hash_password(reg_pass1), reg_name, reg_email, "staff", str(datetime.datetime.now())
                    ])
                    st.success("✅ 註冊成功！請切換至「登入系統」分頁進行登入。")

    with tab3:
        st.subheader("忘記密碼 / 重新寄送臨時密碼")
        fp_user = st.text_input("請輸入您的帳號")
        fp_email = st.text_input("請輸入註冊時綁定的信箱")
        if st.button("寄送重設密碼信"):
            users_df = safe_get_df(ws_users)
            if not users_df.empty:
                match = users_df[(users_df["username"] == fp_user) & (users_df["email"] == fp_email)]
                if not match.empty:
                    new_temp_pass = ''.join(random.choices(string.ascii_letters + string.digits, k=8))
                    row_idx = int(match.index[0]) + 2
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
    st.sidebar.caption(f"身份：{'管理員' if user['role'] == 'admin' else '部門同仁'}")
    
    menu_options = ["📝 每日工作回報", "🔒 修改密碼"]
    if user["role"] == "admin":
        menu_options.insert(0, "📊 主管後台數據中心")
        
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
        
        pack_start_str = "無"
        pack_end_str = "無"
        pack_count = 0
        pack_duration = 0
        min_per_pack = 0
        
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
        
        cs_start_str = "無"
        cs_end_str = "無"
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

        live_start_str = "無"
        live_end_str = "無"
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
        file_8d_name = ""
        file_8d_b64 = ""
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
                    ship_count,
                    live_start_str,
                    live_end_str,
                    has_complaint,
                    c_who, c_where, c_when, c_what, c_why, c_how,
                    "是" if has_8d else "否",
                    file_8d_name,
                    file_8d_b64,
                    "\n".join(tasks_list),
                    improvement_proposal,
                    peer_feedback
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
                users_df = safe_get_df(ws_users)
                row_idx = int(users_df[users_df["username"] == user["username"]].index[0]) + 2
                new_hash = hash_password(new_p1)
                ws_users.update_cell(row_idx, 2, new_hash)
                st.session_state.user_info["password_hash"] = new_hash
                st.success("✅ 密碼更新成功！")

    # ---------------- 功能 C：主管後台數據中心 ----------------
    elif choice == "📊 主管後台數據中心" and user["role"] == "admin":
        st.title("📊 黃金部門 - 主管後台數據中心")
        
        df = safe_get_df(ws_reports)
        if df.empty:
            st.info("目前尚無任何同仁提交回報資料。您可以先點選左側「📝 每日工作回報」測試送出一筆！")
        else:
            # 確保數值欄位轉為數字型態，方便統計加總
            for col in ["pack_count", "pack_duration_min", "min_per_pack", "ship_count"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                    
            df["月份"] = df["report_date"].astype(str).str.slice(0, 7)
            
            # 篩選器
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                months = ["全部月份"] + sorted(df["月份"].unique().tolist(), reverse=True)
                sel_month = st.selectbox("📅 選擇篩選月份", months)
            with col_f2:
                names = ["全部同仁"] + sorted(df["name"].unique().tolist())
                sel_name = st.selectbox("👤 選擇篩選同仁", names)
                
            filtered_df = df.copy()
            if sel_month != "全部月份":
                filtered_df = filtered_df[filtered_df["月份"] == sel_month]
            if sel_name != "全部同仁":
                filtered_df = filtered_df[filtered_df["name"] == sel_name]

            # 1. 月底/區間 個人包貨與寄貨總計 & 包貨效率自動計算
            st.subheader("📈 同仁包貨效率與寄貨總計看板")
            if not filtered_df.empty:
                summary_df = filtered_df.groupby("name").agg(
                    出勤回報天數=("report_id", "count"),
                    總包貨件數=("pack_count", "sum"),
                    總包貨耗時_分鐘=("pack_duration_min", "sum"),
                    總寄貨件數=("ship_count", "sum"),
                    客訴發生次數=("has_complaint", lambda x: (x == "有").sum())
                ).reset_index()
                
                summary_df["平均每件包貨耗時(分鐘/件)"] = summary_df.apply(
                    lambda r: round(r["總包貨耗時_分鐘"] / r["總包貨件數"], 2) if r["總包貨件數"] > 0 else 0, axis=1
                )
                summary_df.rename(columns={"name": "同仁姓名"}, inplace=True)
                st.dataframe(summary_df, use_container_width=True)
            
            st.divider()
            
            # 2. 匯出報表功能 (Excel)
            st.subheader("📥 匯出明細報表 (Excel / 供轉存 PDF)")
            export_df = filtered_df.drop(columns=["file_8d_b64"], errors="ignore").rename(columns={
                "submit_time": "送出時間", "report_date": "回報日期", "name": "姓名",
                "work_start": "上班時間", "work_end": "下班時間",
                "pack_start": "包貨開始", "pack_end": "包貨結束", "pack_count": "包貨件數",
                "pack_duration_min": "包貨總分鐘", "min_per_pack": "單件包貨耗時(分)",
                "cs_start": "客服開始", "cs_end": "客服結束", "ship_count": "寄貨件數",
                "live_start": "直播開始", "live_end": "直播結束",
                "has_complaint": "有無客訴", "c_who": "客訴_Who", "c_where": "客訴_Where",
                "c_when": "客訴_When", "c_what": "客訴_What", "c_why": "客訴_Why", "c_how": "客訴_How",
                "has_8d": "有無8D", "file_8d_name": "8D檔名",
                "daily_tasks": "今日處理事項", "improvement_proposal": "提案改善", "peer_feedback": "對同仁想法"
            })
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                export_df.to_excel(writer, index=False, sheet_name="回報明細")
                if not filtered_df.empty:
                    summary_df.to_excel(writer, index=False, sheet_name="月度統計彙整")
            
            st.download_button(
                label="📊 點擊下載篩選結果 (Excel 報表)",
                data=output.getvalue(),
                file_name=f"黃金部門工作回報_{sel_month}_{sel_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            st.dataframe(export_df, use_container_width=True)
            
            # 3. 8D 檔案下載專區
            df_8d = filtered_df[filtered_df["has_8d"] == "是"]
            if not df_8d.empty:
                st.divider()
                st.subheader("📎 重大異常 8D 報告檔案下載區")
                for idx, row in df_8d.iterrows():
                    if row["file_8d_b64"]:
                        b64_bytes = base64.b64decode(row["file_8d_b64"])
                        st.download_button(
                            label=f"下載 {row['report_date']} - {row['name']} 的 8D 檔案 ({row['file_8d_name']})",
                            data=b64_bytes,
                            file_name=row["file_8d_name"],
                            key=f"dl_{row['report_id']}"
                        )
