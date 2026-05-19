import base64
import csv
import hashlib
import html
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
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from PIL import Image, ImageOps
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
    "signature_image_url",
    "receipt_file_url",
    "receipt_folder_url",
    "billing_month",
    "billing_status",
    "archive_note",
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



def get_drive_folder_id() -> str:
    """Google Drive 母資料夾 ID。這個資料夾要分享給 service account 並給編輯者權限。"""
    return str(get_secret("GOOGLE_DRIVE_FOLDER_ID", "")).strip()


def get_google_credentials():
    _, _, _, service_account_info = get_required_config()
    if not service_account_info:
        raise RuntimeError("找不到 gcp_service_account 設定。")

    service_account_info = dict(service_account_info)
    if "private_key" in service_account_info:
        service_account_info["private_key"] = service_account_info["private_key"].replace("\\n", "\n")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    return Credentials.from_service_account_info(service_account_info, scopes=scopes)


def show_config_error():
    st.error("系統尚未完成設定，請先設定 Streamlit Secrets。")
    with st.expander("需要設定哪些 Secrets？", expanded=True):
        st.code(
            """
APP_BASE_URL = "https://你的-app.streamlit.app"
ADMIN_PASSWORD = "請設定一組廠內管理密碼"
GOOGLE_SHEET_ID = "你的 Google 試算表 ID"
GOOGLE_DRIVE_FOLDER_ID = "你的 Google Drive 簽收憑證資料夾 ID"

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

    try:
        credentials = get_google_credentials()
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
    """
    讀取 Google Sheet 全部資料。
    不使用 gspread 的 get_all_records()，避免第一列標題有重複、空白或舊欄位時直接報錯。
    固定用程式內建 HEADERS 當欄位標準，比較適合 MVP 持續升級。
    """
    ws = get_worksheet()
    ensure_headers(ws)

    values = ws.get_all_values()
    if not values or len(values) <= 1:
        return []

    records = []
    for row in values[1:]:
        if not any(str(cell).strip() for cell in row):
            continue

        padded_row = list(row) + [""] * max(0, len(HEADERS) - len(row))
        record = {}
        for index, header in enumerate(HEADERS):
            record[header] = padded_row[index] if index < len(padded_row) else ""
        records.append(record)

    return records


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



# ---------- Google Drive 憑證保存 ----------
@st.cache_resource(show_spinner=False)
def get_drive_service():
    credentials = get_google_credentials()
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def drive_escape_query_value(text: str) -> str:
    return str(text).replace("'", "\\'")


def get_or_create_drive_folder(name: str, parent_id: str) -> str:
    """在指定 parent folder 下找同名資料夾；不存在就建立。"""
    service = get_drive_service()
    safe_name = drive_escape_query_value(name)
    safe_parent = drive_escape_query_value(parent_id)
    query = (
        "mimeType='application/vnd.google-apps.folder' "
        f"and name='{safe_name}' "
        f"and '{safe_parent}' in parents "
        "and trashed=false"
    )

    result = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name, webViewLink)",
        pageSize=1,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    files = result.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = service.files().create(
        body=metadata,
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return folder["id"]


def upload_bytes_to_drive(filename: str, data: bytes, mime_type: str, parent_id: str) -> dict:
    service = get_drive_service()
    metadata = {"name": filename, "parents": [parent_id]}
    media = MediaIoBaseUpload(BytesIO(data), mimetype=mime_type, resumable=False)
    created = service.files().create(
        body=metadata,
        media_body=media,
        fields="id, name, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return created


def month_from_record(record: dict) -> str:
    for key in ("delivery_date", "signed_at", "created_at"):
        value = str(record.get(key, "") or "").strip()
        if len(value) >= 7:
            return value[:7]
    return datetime.now(TAIPEI_TZ).strftime("%Y-%m")


def drive_folder_url(folder_id: str) -> str:
    return f"https://drive.google.com/drive/folders/{folder_id}" if folder_id else ""


def save_company_accounting_copy_to_drive(record: dict) -> dict:
    """
    數位化後的公司/帳務聯：
    - 簽收憑證 HTML 存 Google Drive
    - 簽名或簽章圖片也存 Google Drive
    - Google Sheet 只保存連結與台帳欄位
    """
    parent_folder_id = get_drive_folder_id()
    if not parent_folder_id:
        return {"archive_note": "未設定 GOOGLE_DRIVE_FOLDER_ID，尚未自動保存到 Google Drive。"}

    client_name = str(record.get("client_name", "") or "未命名客戶")
    product_name = str(record.get("product_name", "") or "未命名品名")
    month_text = month_from_record(record)

    month_folder_id = get_or_create_drive_folder(month_text, parent_folder_id)
    client_folder_id = get_or_create_drive_folder(client_name, month_folder_id)

    date_text = str(record.get("delivery_date", "") or datetime.now(TAIPEI_TZ).strftime("%Y-%m-%d"))
    client_safe = safe_filename_text(client_name)
    product_safe = safe_filename_text(product_name)
    token_safe = safe_filename_text(record.get("token", ""))[:12]
    file_prefix = f"{date_text}_{client_safe}_{product_safe}_{token_safe}"

    result = {
        "receipt_folder_url": drive_folder_url(client_folder_id),
        "billing_month": month_text,
        "billing_status": str(record.get("billing_status", "") or "未結帳"),
        "archive_note": "已自動保存公司/帳務聯到 Google Drive。",
    }

    signature_b64 = str(record.get("signature_png_base64", "") or "")
    signature_mime_type = get_signature_mime_type(record)
    if signature_b64:
        ext = "jpg" if signature_mime_type == "image/jpeg" else "png"
        signature_bytes = base64.b64decode(signature_b64)
        signature_file = upload_bytes_to_drive(
            filename=f"{file_prefix}_簽收圖像.{ext}",
            data=signature_bytes,
            mime_type=signature_mime_type,
            parent_id=client_folder_id,
        )
        result["signature_image_url"] = signature_file.get("webViewLink", "")

    receipt_html = build_receipt_html(record)
    receipt_file = upload_bytes_to_drive(
        filename=f"{file_prefix}_簽收憑證.html",
        data=receipt_html.encode("utf-8"),
        mime_type="text/html",
        parent_id=client_folder_id,
    )
    result["receipt_file_url"] = receipt_file.get("webViewLink", "")

    return result


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



def uploaded_image_to_signature_base64(uploaded_file):
    """
    將客戶上傳的簽章圖片壓縮成可存入 Google Sheet 的 base64。
    使用 JPEG 是因為手機照片通常很大，壓縮後比較不會超過 Google Sheet 單格限制。
    """
    if uploaded_file is None:
        return "", "image/jpeg"

    raw_bytes = uploaded_file.getvalue()
    img = Image.open(BytesIO(raw_bytes))
    img = ImageOps.exif_transpose(img)

    if img.mode in ("RGBA", "LA"):
        background = Image.new("RGB", img.size, "white")
        background.paste(img.convert("RGBA"), mask=img.convert("RGBA").split()[-1])
        img = background
    else:
        img = img.convert("RGB")

    # 多輪壓縮，確保 base64 長度低於 Google Sheet 單格上限。
    attempts = [
        ((640, 360), 82),
        ((520, 300), 76),
        ((420, 260), 70),
        ((340, 220), 64),
        ((280, 180), 58),
    ]

    last_b64 = ""
    for size, quality in attempts:
        work_img = img.copy()
        work_img.thumbnail(size)

        buffer = BytesIO()
        work_img.save(buffer, format="JPEG", quality=quality, optimize=True)
        b64 = base64.b64encode(buffer.getvalue()).decode("utf-8")
        last_b64 = b64

        if len(b64) <= 45000:
            return b64, "image/jpeg"

    return last_b64, "image/jpeg"


def get_signature_metadata(record: dict) -> dict:
    raw = str(record.get("signature_json", "") or "").strip()
    if not raw:
        return {}

    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            return data
    except Exception:
        return {}

    return {}


def get_signature_method_label(record: dict) -> str:
    metadata = get_signature_metadata(record)
    method = str(metadata.get("method", "") or record.get("signed_device_hint", ""))

    if method == "stamp_image_upload":
        return "簽章圖片"

    return "電子簽名"


def get_signature_mime_type(record: dict) -> str:
    metadata = get_signature_metadata(record)
    mime_type = str(metadata.get("mime_type", "") or "").strip()

    if mime_type in ("image/png", "image/jpeg", "image/webp"):
        return mime_type

    return "image/png"


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
        "signature_image_url": "",
        "receipt_file_url": "",
        "receipt_folder_url": "",
        "billing_month": delivery_date[:7] if delivery_date else "",
        "billing_status": "未結帳",
        "archive_note": "",
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
                "簽收方式": get_signature_method_label(r) if r.get("status") == STATUS_SIGNED else "",
                "月結狀態": r.get("billing_status", ""),
                "簽收憑證": r.get("receipt_file_url", ""),
                "簽收圖像": r.get("signature_image_url", ""),
                "備註": r.get("note", ""),
                "簽收連結": r.get("sign_url", ""),
            }
        )

    st.dataframe(
        display_rows,
        use_container_width=True,
        hide_index=True,
        column_config={
            "簽收憑證": st.column_config.LinkColumn("簽收憑證"),
            "簽收圖像": st.column_config.LinkColumn("簽收圖像"),
            "簽收連結": st.column_config.LinkColumn("簽收連結"),
        },
    )

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




# ---------- 簽收憑證 / 客戶留底 ----------
def safe_filename_text(text: str) -> str:
    text = str(text or "").strip()
    keep = []
    for ch in text:
        if ch.isalnum() or ch in ("-", "_"):
            keep.append(ch)
        elif ch in (" ", ".", "．"):
            keep.append("_")
    result = "".join(keep).strip("_")
    return result[:60] or "receipt"


def build_receipt_html(record: dict) -> str:
    """
    建立客戶可下載留底的 HTML 簽收憑證。
    HTML 比 PDF 更適合第一版：手機、電腦都能直接開，也能用瀏覽器另存 PDF。
    """
    company_name = "新豐製版"
    client_name = html.escape(str(record.get("client_name", "")))
    product_name = html.escape(str(record.get("product_name", "")))
    quantity = html.escape(str(record.get("quantity", "")))
    delivery_date = html.escape(str(record.get("delivery_date", "")))
    note = html.escape(str(record.get("note", "")))
    signed_at = html.escape(str(record.get("signed_at", "")))
    signer_name = html.escape(str(record.get("signer_name", "")))
    signer_phone = html.escape(str(record.get("signer_phone", "")))
    signer_note = html.escape(str(record.get("signer_note", "")))
    token = html.escape(str(record.get("token", "")))
    sign_url = html.escape(str(record.get("sign_url", "")))
    signature_b64 = str(record.get("signature_png_base64", "") or "")
    signature_mime_type = html.escape(get_signature_mime_type(record))
    signature_method = html.escape(get_signature_method_label(record))

    signature_html = ""
    if signature_b64:
        signature_html = f'<img class="signature" src="data:{signature_mime_type};base64,{signature_b64}" alt="簽收圖像">'

    return f'''<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{company_name}｜簽收憑證</title>
<style>
body {{
    font-family: -apple-system, BlinkMacSystemFont, "Noto Sans TC", "Microsoft JhengHei", Arial, sans-serif;
    background: #f6f7f9;
    color: #222;
    margin: 0;
    padding: 24px;
}}
.receipt {{
    max-width: 760px;
    margin: 0 auto;
    background: #fff;
    border: 1px solid #ddd;
    border-radius: 16px;
    padding: 28px;
}}
h1 {{
    margin: 0 0 8px;
    font-size: 28px;
}}
.subtitle {{
    color: #666;
    margin-bottom: 24px;
}}
.status {{
    display: inline-block;
    background: #e8f7ee;
    color: #126b34;
    border: 1px solid #b9e5c9;
    border-radius: 999px;
    padding: 6px 12px;
    font-weight: 700;
    margin-bottom: 18px;
}}
table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 12px;
}}
th, td {{
    text-align: left;
    border-bottom: 1px solid #eee;
    padding: 12px 8px;
    vertical-align: top;
}}
th {{
    width: 160px;
    color: #555;
    font-weight: 700;
}}
.signature-box {{
    margin-top: 24px;
    border: 1px dashed #aaa;
    border-radius: 12px;
    padding: 16px;
    min-height: 120px;
}}
.signature {{
    max-width: 100%;
    max-height: 220px;
}}
.footer {{
    margin-top: 24px;
    color: #777;
    font-size: 13px;
    line-height: 1.6;
}}
@media print {{
    body {{ background: #fff; padding: 0; }}
    .receipt {{ border: none; border-radius: 0; }}
}}
</style>
</head>
<body>
<div class="receipt">
    <h1>{company_name}｜簽收憑證</h1>
    <div class="subtitle">此憑證由數位簽收系統產生，供客戶與廠內雙方留存。</div>
    <div class="status">已簽收</div>

    <table>
        <tr><th>客戶名稱</th><td>{client_name}</td></tr>
        <tr><th>品名 / 刀模編號</th><td>{product_name}</td></tr>
        <tr><th>數量</th><td>{quantity}</td></tr>
        <tr><th>出貨日期</th><td>{delivery_date}</td></tr>
        <tr><th>廠內備註</th><td>{note}</td></tr>
        <tr><th>簽收時間</th><td>{signed_at}</td></tr>
        <tr><th>簽收方式</th><td>{signature_method}</td></tr>
        <tr><th>電話 / 分機</th><td>{signer_phone}</td></tr>
        <tr><th>簽收備註</th><td>{signer_note}</td></tr>
        <tr><th>簽收單號</th><td>{token}</td></tr>
        <tr><th>原簽收網址</th><td>{sign_url}</td></tr>
    </table>

    <div class="signature-box">
        <strong>簽名：</strong><br>
        {signature_html}
    </div>

    <div class="footer">
        建議客戶下載此 HTML 檔留存；也可用瀏覽器的列印功能另存為 PDF。<br>
        若內容有疑問，請聯絡新豐製版確認。
    </div>
</div>
</body>
</html>'''


def build_receipt_download_filename(record: dict) -> str:
    client = safe_filename_text(record.get("client_name", ""))
    product = safe_filename_text(record.get("product_name", ""))
    signed_at = safe_filename_text(str(record.get("signed_at", "")).replace(":", "").replace("-", ""))
    return f"新豐製版_簽收憑證_{client}_{product}_{signed_at}.html"


def show_customer_receipt(record: dict):
    st.success("此單已完成簽收。以下資料可供客戶留底。")

    with st.container(border=True):
        st.write("### 📄 簽收憑證")
        st.write(f"**客戶名稱：** {record.get('client_name', '')}")
        st.write(f"**品名 / 刀模編號：** {record.get('product_name', '')}")
        if record.get("quantity", ""):
            st.write(f"**數量：** {record.get('quantity', '')}")
        if record.get("delivery_date", ""):
            st.write(f"**出貨日期：** {record.get('delivery_date', '')}")
        st.write(f"**簽收時間：** {record.get('signed_at', '')}")
        st.write(f"**簽收方式：** {get_signature_method_label(record)}")
        if record.get("signer_phone", ""):
            st.write(f"**電話 / 分機：** {record.get('signer_phone', '')}")
        if record.get("signer_note", ""):
            st.write(f"**簽收備註：** {record.get('signer_note', '')}")

        show_signature_from_base64(str(record.get("signature_png_base64", "")))

    receipt_html = build_receipt_html(record)
    filename = build_receipt_download_filename(record)

    st.download_button(
        label="下載簽收憑證 HTML",
        data=receipt_html.encode("utf-8"),
        file_name=filename,
        mime="text/html",
        help="下載後可直接打開，也可用瀏覽器列印成 PDF。",
    )

    st.info("客戶也可以保留這個簽收網址；日後再次開啟同一連結，會看到已簽收紀錄，不會重複簽收。")

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

    st.write("請確認以下資料，收到版後以電子簽名或簽章圖片完成簽收。")

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
        show_customer_receipt(record)
        st.stop()

    st.warning("若尚未收到版，請先不要簽收。")

    signer_note = st.text_area("簽收備註", placeholder="可留空，例如：已收到、數量確認無誤")

    st.write("### 請選擇簽收方式")
    sign_method = st.radio(
        "簽收方式",
        ["在畫面簽名", "上傳簽章圖片"],
        horizontal=True,
        key=f"sign_method_{token}",
    )

    canvas_result = None
    stamp_uploaded_file = None

    if sign_method == "在畫面簽名":
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
    else:
        st.write("### 上傳簽章圖片")
        st.info("可上傳印章照片或簽章圖片。系統會壓縮後只作為本筆簽收紀錄保存。")
        stamp_uploaded_file = st.file_uploader(
            "選擇圖片檔",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=False,
            key=f"stamp_upload_{token}",
        )
        if stamp_uploaded_file is not None:
            st.image(stamp_uploaded_file, caption="簽章圖片預覽", use_container_width=True)

    confirm = st.checkbox("我確認已收到上述物品，並同意以本次電子簽名或簽章圖片作為簽收紀錄。")

    if st.button("確認送出簽收", type="primary"):
        if not confirm:
            st.warning("請先勾選確認收到物品。")
            return

        signature_mime_type = "image/png"
        signed_device_hint = "signature_canvas"

        if sign_method == "在畫面簽名":
            signature_base64 = canvas_to_png_base64(canvas_result.image_data if canvas_result else None)

            if not signature_base64:
                st.warning("請先在簽名區簽名。")
                return

            signature_json = json.dumps(
                {
                    "method": "signature_canvas",
                    "mime_type": "image/png",
                    "canvas_data": canvas_result.json_data or {},
                },
                ensure_ascii=False,
            )
            signer_name_value = "電子簽收"

        else:
            if stamp_uploaded_file is None:
                st.warning("請先上傳簽章圖片。")
                return

            try:
                signature_base64, signature_mime_type = uploaded_image_to_signature_base64(stamp_uploaded_file)
            except Exception:
                st.error("簽章圖片讀取失敗，請改用 PNG/JPG 圖片重新上傳。")
                return

            if not signature_base64:
                st.warning("請先上傳簽章圖片。")
                return

            signature_json = json.dumps(
                {
                    "method": "stamp_image_upload",
                    "mime_type": signature_mime_type,
                    "filename": stamp_uploaded_file.name,
                },
                ensure_ascii=False,
            )
            signer_name_value = "簽章圖片簽收"
            signed_device_hint = "stamp_image_upload"

        if len(signature_base64) > 45000:
            st.error("簽收圖片資料過大，請改用較小或較清楚的圖片重新上傳，或改用畫面簽名。")
            return

        record["status"] = STATUS_SIGNED
        record["signed_at"] = now_text()
        record["signer_name"] = signer_name_value
        record["signer_phone"] = ""
        record["signer_note"] = signer_note.strip()
        record["signature_png_base64"] = signature_base64
        record["signature_json"] = signature_json
        record["signed_device_hint"] = signed_device_hint
        record["billing_month"] = month_from_record(record)
        record["billing_status"] = record.get("billing_status", "") or "未結帳"

        with st.spinner("正在保存簽收紀錄..."):
            try:
                drive_result = save_company_accounting_copy_to_drive(record)
                record.update(drive_result)
            except Exception as exc:
                record["archive_note"] = f"Google Drive 自動保存失敗：{type(exc).__name__}"
                st.warning("簽收已完成，但公司/帳務聯自動保存到 Google Drive 時發生問題；請通知新豐製版確認後台紀錄。")

            update_record(row_number, record)

        st.balloons()
        show_customer_receipt(record)
        st.stop()


# ---------- 廠內端主畫面 ----------
def admin_page():
    if not config_is_ready():
        show_config_error()
        st.stop()

    if not admin_login():
        st.stop()

    st.title("📦 新豐製版｜數位簽收管理")
    st.caption("廠內建立簽收單後，系統會產生 token 專屬連結；客戶只能透過該連結簽收。")

    if not get_drive_folder_id():
        st.warning("尚未設定 GOOGLE_DRIVE_FOLDER_ID；客戶簽收仍可完成，但公司/帳務聯不會自動保存到 Google Drive。")

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
