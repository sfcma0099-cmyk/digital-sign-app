import base64
import csv
import hashlib
import hmac
import json
import secrets
import urllib.parse
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import gspread
import numpy as np
import streamlit as st
from google.oauth2.service_account import Credentials
from PIL import Image
from streamlit_drawable_canvas import st_canvas


# =========================================================
# 新豐製版：Token 版數位簽收系統 v1
# ---------------------------------------------------------
# 這支 app.py 同時包含：
# 1) 廠內端：建立簽收單、批量建立連結、查詢簽收狀態
# 2) 客戶端：透過 ?token=xxxx 開啟簽收頁面
#
# 資料儲存在 Google Sheets，不存放在 Streamlit 雲端硬碟。
# =========================================================


# ---------- 基本設定 ----------
TAIPEI_TZ = ZoneInfo("Asia/Taipei")

POPULAR_CLIENTS = ["禎曜", "金森", "三和", "合歷", "佳鑫", "紙城", "易昇", "榮星"]

HEADERS = [
    "token",
    "created_at",
    "status",
    "client_name",
    "product_name",
    "quantity",
    "delivery_date",
    "note",
    "sign_url",
    "signed_at",
    "signer_name",
    "signer_phone",
    "signer_note",
    "signature_png_base64",
    "signature_json",
    "signed_device_hint",
]

STATUS_PENDING = "未簽收"
STATUS_SIGNED = "已簽收"


# ---------- Streamlit 頁面設定 ----------
st.set_page_config(
    page_title="新豐製版｜數位簽收系統",
    page_icon="🖊️",
    layout="centered",
)


# ---------- Secrets 讀取 ----------
def get_secret(key: str, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def get_required_config():
    app_base_url = str(get_secret("APP_BASE_URL", "")).rstrip("/")
    admin_password = str(get_secret("ADMIN_PASSWORD", ""))
    sheet_id = str(get_secret("GOOGLE_SHEET_ID", ""))

    try:
        service_account_info = dict(st.secrets["gcp_service_account"])
    except Exception:
        service_account_info = None

    return app_base_url, admin_password, sheet_id, service_account_info


def show_config_error():
    st.error("系統尚未完成設定，請先設定 Streamlit Secrets。")
    with st.expander("需要設定哪些 Secrets？", expanded=True):
        st.code(
            """
APP_BASE_URL = "https://你的-app.streamlit.app"
ADMIN_PASSWORD = "請設定一組廠內管理密碼"
GOOGLE_SHEET_ID = "你的 Google 試算表 ID"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = \"\"\"-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----\\n\"\"\"
client_email = "你的-service-account@xxx.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
            """.strip(),
            language="toml",
        )


def config_is_ready() -> bool:
    app_base_url, admin_password, sheet_id, service_account_info = get_required_config()
    return bool(app_base_url and admin_password and sheet_id and service_account_info)


# ---------- Google Sheets ----------
@st.cache_resource(show_spinner=False)
def get_worksheet():
    app_base_url, admin_password, sheet_id, service_account_info = get_required_config()

    if not service_account_info:
        raise RuntimeError("找不到 gcp_service_account 設定。")

    # 有些貼上 Secrets 的 private_key 會把換行變成 \\n，這裡統一轉回真正換行。
    if "private_key" in service_account_info:
        service_account_info["private_key"] = service_account_info["private_key"].replace("\\n", "\n")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    try:
        credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
        client = gspread.authorize(credentials)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.sheet1
        ensure_headers(worksheet)
        return worksheet

    except PermissionError:
        st.error("Google Sheet 權限不足：請確認試算表已分享給 service account，且權限是「編輯者」。")
        st.info(f"目前程式使用的 service account：{service_account_info.get('client_email', '讀取不到 client_email')}")
        st.info(f"目前程式使用的 GOOGLE_SHEET_ID：{sheet_id}")
        st.stop()

    except Exception as exc:
        st.error("連接 Google Sheet 時發生錯誤。")
        st.info(f"錯誤類型：{type(exc).__name__}")
        st.info("請檢查 Streamlit Secrets、Google Sheets API、Google Drive API，以及試算表共用權限。")
        st.stop()


def ensure_headers(worksheet):
    current_first_row = worksheet.row_values(1)

    if not current_first_row:
        worksheet.append_row(HEADERS, value_input_option="USER_ENTERED")
        return

    # 若第一列不是新版欄位，直接更新第一列標題，不會刪掉舊資料。
    if current_first_row != HEADERS:
        end_col = column_letter(len(HEADERS))
        worksheet.update(range_name=f"A1:{end_col}1", values=[HEADERS])


def column_letter(n: int) -> str:
    """1 -> A, 2 -> B, 27 -> AA"""
    result = ""
    while n:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def now_text() -> str:
    return datetime.now(TAIPEI_TZ).strftime("%Y-%m-%d %H:%M:%S")


def today_text() -> str:
    return datetime.now(TAIPEI_TZ).strftime("%Y-%m-%d")


def get_all_records():
    ws = get_worksheet()
    ensure_headers(ws)
    return ws.get_all_records(default_blank="")


def find_record_by_token(token: str):
    records = get_all_records()
    for row_number, record in enumerate(records, start=2):
        if str(record.get("token", "")).strip() == token:
            return row_number, record
    return None, None


def append_record(record: dict):
    ws = get_worksheet()
    row_values = [record.get(header, "") for header in HEADERS]
    ws.append_row(row_values, value_input_option="USER_ENTERED")


def update_record(row_number: int, record: dict):
    ws = get_worksheet()
    row_values = [record.get(header, "") for header in HEADERS]
    end_col = column_letter(len(HEADERS))
    ws.update(range_name=f"A{row_number}:{end_col}{row_number}", values=[row_values])


# ---------- Token / URL ----------
def generate_token() -> str:
    # URL-safe、高熵亂數。客戶無法靠改網址猜到別張單。
    return secrets.token_urlsafe(24)


def build_sign_url(token: str) -> str:
    app_base_url, _, _, _ = get_required_config()
    return f"{app_base_url.rstrip('/')}/?token={urllib.parse.quote(token)}"


# ---------- 管理端密碼 ----------
def password_ok(input_password: str, actual_password: str) -> bool:
    return hmac.compare_digest(input_password, actual_password)


def admin_login():
    _, admin_password, _, _ = get_required_config()

    if st.session_state.get("admin_ok") is True:
        return True

    st.title("🔐 新豐製版｜廠內管理端")
    st.caption("請輸入廠內管理密碼。客戶簽收頁不需要密碼，只需要專屬 token 連結。")

    with st.form("admin_login_form"):
        input_password = st.text_input("管理密碼", type="password")
        submitted = st.form_submit_button("登入")

    if submitted:
        if password_ok(input_password, admin_password):
            st.session_state["admin_ok"] = True
            st.rerun()
        else:
            st.error("密碼錯誤。")

    return False


# ---------- 簽名處理 ----------
def canvas_to_png_base64(image_data) -> str:
    """
    將簽名畫布轉成 PNG base64。
    會自動裁掉空白區，避免 Google Sheets 內儲存太大的簽名資料。
    """
    if image_data is None:
        return ""

    arr = np.asarray(image_data).astype("uint8")

    if arr.ndim != 3 or arr.shape[2] < 3:
        return ""

    rgb = arr[:, :, :3]

    if arr.shape[2] >= 4:
        alpha = arr[:, :, 3]
    else:
        alpha = np.full(rgb.shape[:2], 255, dtype=np.uint8)

    # 背景是白色，簽名是黑色。只抓有墨跡的像素。
    ink_mask = (alpha > 0) & np.any(rgb < 245, axis=2)

    if not ink_mask.any():
        return ""

    ys, xs = np.where(ink_mask)
    pad = 12
    y1 = max(int(ys.min()) - pad, 0)
    y2 = min(int(ys.max()) + pad, arr.shape[0])
    x1 = max(int(xs.min()) - pad, 0)
    x2 = min(int(xs.max()) + pad, arr.shape[1])

    cropped = arr[y1:y2, x1:x2]
    img = Image.fromarray(cropped, mode="RGBA")
    img.thumbnail((600, 240))

    buffer = BytesIO()
    img.save(buffer, format="PNG", optimize=True)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def show_signature_from_base64(signature_b64: str):
    if not signature_b64:
        st.info("此簽收紀錄沒有簽名圖。")
        return

    try:
        image_bytes = base64.b64decode(signature_b64)
        st.image(image_bytes, caption="客戶簽名", use_container_width=True)
    except Exception:
        st.warning("簽名圖無法顯示，但原始資料仍保留在資料表中。")


# ---------- 廠內端：建立簽收單 ----------
def make_empty_record(
    client_name: str,
    product_name: str,
    quantity: str,
    delivery_date: str,
    note: str,
) -> dict:
    token = generate_token()
    sign_url = build_sign_url(token)

    return {
        "token": token,
        "created_at": now_text(),
        "status": STATUS_PENDING,
        "client_name": client_name.strip(),
        "product_name": product_name.strip(),
        "quantity": quantity.strip(),
        "delivery_date": delivery_date.strip(),
        "note": note.strip(),
        "sign_url": sign_url,
        "signed_at": "",
        "signer_name": "",
        "signer_phone": "",
        "signer_note": "",
        "signature_png_base64": "",
        "signature_json": "",
        "signed_device_hint": "",
    }


def admin_create_single_tab():
    st.subheader("📝 單筆建立簽收單")

    client_type = st.radio(
        "客戶輸入方式",
        ["從常用清單選擇", "手動自行輸入"],
        horizontal=True,
        key="single_client_type",
    )

    if client_type == "從常用清單選擇":
        client_name = st.selectbox("客戶名稱", POPULAR_CLIENTS, key="single_client_select")
    else:
        client_name = st.text_input("客戶名稱", placeholder="例如：禎曜", key="single_client_text")

    product_name = st.text_input("出貨品名 / 刀模編號", placeholder="例如：達創-001")
    quantity = st.text_input("數量", placeholder="例如：1片 / 2組，可留空")
    delivery_date = st.date_input("出貨日期").strftime("%Y-%m-%d")
    note = st.text_area("備註", placeholder="例如：早班配送、交給王先生，可留空")

    if st.button("建立簽收單並產生網址", type="primary"):
        if not client_name.strip() or not product_name.strip():
            st.warning("請至少填寫「客戶名稱」與「出貨品名 / 刀模編號」。")
            return

        record = make_empty_record(client_name, product_name, quantity, delivery_date, note)
        append_record(record)

        st.success("已建立簽收單。")
        st.write("### 給客戶的簽收連結")
        st.code(record["sign_url"], language="text")

        line_text = (
            f"【新豐製版簽收通知】\n"
            f"客戶：{record['client_name']}\n"
            f"品名/單號：{record['product_name']}\n"
            f"請您收到版後，開啟以下連結完成簽收：\n"
            f"{record['sign_url']}"
        )
        st.write("### 可複製到 LINE 的訊息")
        st.code(line_text, language="text")


# ---------- 廠內端：批量建立 ----------
def parse_batch_line(line: str):
    """
    支援：
    客戶, 品名
    客戶, 品名, 數量
    客戶, 品名, 數量, 備註
    也支援中文逗號。
    """
    line = line.strip().replace("，", ",")
    if not line:
        return None

    parts = [part.strip() for part in line.split(",")]
    if len(parts) < 2:
        raise ValueError("缺少逗號或欄位不足")

    client_name = parts[0]
    product_name = parts[1]
    quantity = parts[2] if len(parts) >= 3 else ""
    note = ",".join(parts[3:]).strip() if len(parts) >= 4 else ""

    if not client_name or not product_name:
        raise ValueError("客戶名稱或品名不可空白")

    return client_name, product_name, quantity, note


def admin_batch_tab():
    st.subheader("🚀 批量建立簽收單")
    st.info("每一行代表一張簽收單。格式：客戶名稱, 品名單號, 數量, 備註。後兩欄可省略，也支援中文逗號。")

    default_text = "禎曜, 達創-001, 1片, 早班配送\n金森, 刀模-002\n三和, 紙盒-003, 2組"
    batch_input = st.text_area("貼上今日派單資料", value=default_text, height=220)
    delivery_date = st.date_input("這批資料的出貨日期", key="batch_delivery_date").strftime("%Y-%m-%d")

    if st.button("一鍵建立全部簽收單", type="primary"):
        lines = batch_input.splitlines()
        created_records = []
        error_lines = []

        for line_number, line in enumerate(lines, start=1):
            if not line.strip():
                continue

            try:
                parsed = parse_batch_line(line)
                if parsed is None:
                    continue

                client_name, product_name, quantity, note = parsed
                record = make_empty_record(client_name, product_name, quantity, delivery_date, note)
                append_record(record)
                created_records.append(record)

            except Exception as exc:
                error_lines.append((line_number, line, str(exc)))

        if created_records:
            st.success(f"已建立 {len(created_records)} 張簽收單。")

            all_links_text = ""
            for record in created_records:
                all_links_text += (
                    f"【{record['client_name']}】 {record['product_name']}\n"
                    f"請收到版後開啟簽收：{record['sign_url']}\n\n"
                )

            st.write("### 📲 複製給外務或客戶的 LINE 訊息")
            st.code(all_links_text, language="text")

        if error_lines:
            st.error("以下資料沒有建立成功，請修正格式後再補建：")
            for line_number, line, reason in error_lines:
                st.write(f"第 {line_number} 行：`{line}`，原因：{reason}")


# ---------- 廠內端：查詢簽收紀錄 ----------
def admin_dashboard_tab():
    st.subheader("📋 簽收狀態查詢")

    records = get_all_records()
    if not records:
        st.info("目前還沒有簽收單。")
        return

    status_filter = st.selectbox("狀態篩選", ["全部", STATUS_PENDING, STATUS_SIGNED])
    keyword = st.text_input("搜尋客戶 / 品名 / 簽收人", placeholder="可留空")

    visible_records = []
    for record in records:
        if status_filter != "全部" and record.get("status") != status_filter:
            continue

        haystack = " ".join(
            [
                str(record.get("client_name", "")),
                str(record.get("product_name", "")),
                str(record.get("signer_name", "")),
                str(record.get("note", "")),
            ]
        )
        if keyword.strip() and keyword.strip() not in haystack:
            continue

        visible_records.append(record)

    summary_pending = sum(1 for r in records if r.get("status") == STATUS_PENDING)
    summary_signed = sum(1 for r in records if r.get("status") == STATUS_SIGNED)

    col1, col2, col3 = st.columns(3)
    col1.metric("總筆數", len(records))
    col2.metric("未簽收", summary_pending)
    col3.metric("已簽收", summary_signed)

    display_rows = []
    for r in visible_records:
        display_rows.append(
            {
                "狀態": r.get("status", ""),
                "建立時間": r.get("created_at", ""),
                "出貨日期": r.get("delivery_date", ""),
                "客戶": r.get("client_name", ""),
                "品名/單號": r.get("product_name", ""),
                "數量": r.get("quantity", ""),
                "簽收時間": r.get("signed_at", ""),
                "簽收人": r.get("signer_name", ""),
                "電話": r.get("signer_phone", ""),
                "備註": r.get("note", ""),
                "簽收連結": r.get("sign_url", ""),
            }
        )

    st.dataframe(display_rows, use_container_width=True, hide_index=True)

    st.write("---")
    st.write("### 查看單筆簽名")
    signed_records = [r for r in visible_records if r.get("status") == STATUS_SIGNED]

    if not signed_records:
        st.info("目前篩選結果沒有已簽收資料。")
        return

    options = [
        f"{r.get('signed_at','')}｜{r.get('client_name','')}｜{r.get('product_name','')}｜{r.get('signer_name','')}"
        for r in signed_records
    ]
    selected_index = st.selectbox("選擇簽收紀錄", range(len(options)), format_func=lambda i: options[i])
    selected = signed_records[selected_index]

    st.write(f"**客戶：** {selected.get('client_name', '')}")
    st.write(f"**品名/單號：** {selected.get('product_name', '')}")
    st.write(f"**簽收人：** {selected.get('signer_name', '')}")
    st.write(f"**簽收時間：** {selected.get('signed_at', '')}")
    show_signature_from_base64(str(selected.get("signature_png_base64", "")))


# ---------- 客戶端：簽收頁 ----------
def customer_signing_page(token: str):
    st.title("🖊️ 新豐製版｜數位簽收單")

    if not config_is_ready():
        show_config_error()
        st.stop()

    row_number, record = find_record_by_token(token)

    if not record:
        st.error("此簽收連結無效，請聯絡新豐製版確認。")
        st.stop()

    st.write("請確認以下資料，收到版後再完成簽收。")

    with st.container(border=True):
        st.write(f"**客戶名稱：** {record.get('client_name', '')}")
        st.write(f"**品名 / 刀模編號：** {record.get('product_name', '')}")
        if record.get("quantity", ""):
            st.write(f"**數量：** {record.get('quantity', '')}")
        if record.get("delivery_date", ""):
            st.write(f"**出貨日期：** {record.get('delivery_date', '')}")
        if record.get("note", ""):
            st.write(f"**備註：** {record.get('note', '')}")

    if record.get("status") == STATUS_SIGNED:
        st.success(f"此單已於 {record.get('signed_at', '')} 完成簽收。")
        st.write(f"**簽收人：** {record.get('signer_name', '')}")
        if record.get("signer_note", ""):
            st.write(f"**簽收備註：** {record.get('signer_note', '')}")
        show_signature_from_base64(str(record.get("signature_png_base64", "")))
        st.stop()

    st.warning("若尚未收到版，請先不要簽收。")

    signer_name = st.text_input("簽收人姓名", placeholder="請輸入簽收人姓名")
    signer_phone = st.text_input("簽收人電話 / 分機", placeholder="可留空")
    signer_note = st.text_area("簽收備註", placeholder="可留空，例如：已收到、數量確認無誤")

    st.write("### 請在下方簽名")
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 0)",
        stroke_width=4,
        stroke_color="#000000",
        background_color="#FFFFFF",
        height=220,
        width=520,
        drawing_mode="freedraw",
        display_toolbar=True,
        update_streamlit=True,
        key=f"signature_canvas_{token}",
    )

    confirm = st.checkbox("我確認已收到上述物品，並同意送出簽收紀錄。")

    if st.button("確認送出簽收", type="primary"):
        if not signer_name.strip():
            st.warning("請輸入簽收人姓名。")
            return

        if not confirm:
            st.warning("請先勾選確認收到物品。")
            return

        signature_png_base64 = canvas_to_png_base64(canvas_result.image_data)

        if not signature_png_base64:
            st.warning("請先在簽名區簽名。")
            return

        if len(signature_png_base64) > 45000:
            st.error("簽名資料過大，請按畫布工具列清除後，用較簡單的筆畫重新簽名。")
            return

        signature_json = json.dumps(canvas_result.json_data or {}, ensure_ascii=False)

        record["status"] = STATUS_SIGNED
        record["signed_at"] = now_text()
        record["signer_name"] = signer_name.strip()
        record["signer_phone"] = signer_phone.strip()
        record["signer_note"] = signer_note.strip()
        record["signature_png_base64"] = signature_png_base64
        record["signature_json"] = signature_json
        record["signed_device_hint"] = "Streamlit web signing"

        update_record(row_number, record)

        st.success("簽收完成，謝謝。")
        st.balloons()
        st.rerun()


# ---------- 廠內端主畫面 ----------
def admin_page():
    if not config_is_ready():
        show_config_error()
        st.stop()

    if not admin_login():
        st.stop()

    st.title("📦 新豐製版｜數位簽收管理")
    st.caption("廠內建立簽收單後，系統會產生 token 專屬連結；客戶只能透過該連結簽收。")

    tab1, tab2, tab3 = st.tabs(["📝 單筆建立", "🚀 批量建立", "📋 狀態查詢"])

    with tab1:
        admin_create_single_tab()

    with tab2:
        admin_batch_tab()

    with tab3:
        admin_dashboard_tab()

    st.write("---")
    if st.button("登出管理端"):
        st.session_state["admin_ok"] = False
        st.rerun()


# ---------- 入口 ----------
def main():
    token = str(st.query_params.get("token", "")).strip()

    if token:
        customer_signing_page(token)
    else:
        admin_page()


if __name__ == "__main__":
    main()
