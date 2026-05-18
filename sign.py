import streamlit as st
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import io
import base64
import requests
from datetime import datetime

# ==========================================
# ⚙️ 可修改的設計與參數區 (Configuration)
# ==========================================
# 1. 🌐 請貼上您剛剛在 Apps Script 部署後取得的「網頁應用程式網址 (URL)」
GOOGLE_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwjrTiGT00k6gZWn9Ky1le1O-QUCPPPxBTt_cYr9cWhJLF6BVhNVN2TbQz-QiB0CsFZ/exec"

# 2. 簽名板畫布設定
CANVAS_STROKE_WIDTH = 3              # 筆跡粗細
CANVAS_STROKE_COLOR = "#000000"      # 筆跡顏色 (黑色)
CANVAS_BACKGROUND_COLOR = "#FFFFFF"  # 背景顏色 (白色)
# ==========================================

def upload_to_drive_via_gas(byte_im, filename):
    """透過 Google Apps Script 將圖檔傳送至 Google Drive"""
    try:
        # 將圖檔轉換為 base64 編碼文字
        base64_img = base64.b64encode(byte_im).decode('utf-8')
        
        # 打包要傳送的資料
        payload = {
            "filename": filename,
            "base64": base64_img
        }
        
        # 傳送給您的 Google 接收站
        response = requests.post(GOOGLE_WEB_APP_URL, json=payload)
        
        if response.status_code == 200 and response.json().get("status") == "success":
            return True
        else:
            st.error(f"雲端接收失敗：{response.text}")
            return False
    except Exception as e:
        st.error(f"傳輸發生錯誤：{e}")
        return False

# ==========================================
# 🎯 URL 參數處理：自動帶入資訊
# ==========================================
params = st.query_params
prefill_client = params.get("client", "")
prefill_product = params.get("product", "")

# ==========================================
# 🎨 網頁介面 (UI)
# ==========================================
st.set_page_config(page_title="數位簽收系統", layout="centered")

st.title("📦 數位簽收確認單")
st.write("您好，請確認下方出貨資訊，確認無誤後請於欄位內簽名。")

client_name = st.text_input("客戶名稱", value=prefill_client)
product_name = st.text_input("出貨品名 / 單號", value=prefill_product)

st.write("---")
st.write("### ✍️ 請在此處簽名")

canvas_result = st_canvas(
    stroke_width=CANVAS_STROKE_WIDTH,
    stroke_color=CANVAS_STROKE_COLOR,
    background_color=CANVAS_BACKGROUND_COLOR,
    height=250,
    width=500,
    drawing_mode="freedraw",
    key="canvas",
)

if st.button("✅ 確認簽署並送出"):
    if not client_name or not product_name:
        st.error("❌ 請填寫客戶名稱與品名單號。")
    elif canvas_result.image_data is None:
        st.warning("⚠️ 請先在上方畫布完成簽名。")
    else:
        # 將畫布內容轉換為圖片
        img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        byte_im = buf.getvalue()

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"{timestamp}_{client_name}_{product_name}_簽名結案.png"

        # 顯示處理中的動畫
        with st.spinner("系統正在將簽收單同步至雲端，請稍候..."):
            success = upload_to_drive_via_gas(byte_im, filename)
            
            if success:
                st.balloons()
                st.success(f"🎉 簽收成功！檔案已安全存入您的 Google 雲端硬碟。")
                st.write("您可以直接關閉此網頁。")