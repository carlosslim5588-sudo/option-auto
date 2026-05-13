# =========================================================
# Get_Option_Data.py（GitHub Actions 完全対応版）
# =========================================================

import os
import time
import json
import pandas as pd
from io import StringIO
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager


# =========================================================
# Driveアップロード関数（★先に定義する）
# =========================================================
def upload_to_drive(file_path, file_name):

    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from google.oauth2 import service_account

    creds_info = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])

    creds = service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/drive.file"]
    )

    service = build("drive", "v3", credentials=creds)

    folder_id = "1lolYbNbQsfohDD_Mx-jz8rsC4s7kI2mN"

    file_metadata = {
        "name": file_name,
        "parents": [folder_id]
    }

    media = MediaFileUpload(file_path, resumable=True)

    service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    print("✅ Google Drive保存完了")


# =========================================================
# 出力ファイル
# =========================================================
now_jst = datetime.utcnow() + timedelta(hours=9)
timestamp_str = now_jst.strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = f"Get_Option_Data_{timestamp_str}.xlsx"


# =========================================================
# Chrome設定
# =========================================================
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")


print("⏳ Chrome起動中...")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

print("✅ Chrome起動成功")


# =========================================================
# メイン処理
# =========================================================
try:
    url = "https://www.jpx.co.jp/markets/derivatives/quotes/index.html"
    driver.get(url)

    try:
        cookie_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".cc-btn.cc-dismiss"))
        )
        cookie_btn.click()
    except:
        pass

    link = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "オプション価格情報"))
    )
    link.click()

    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(5)

    prices = WebDriverWait(driver, 15).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td.a-right.price-now"))
    )

    nikkei_value = prices[0].text if len(prices) > 0 else "N/A"
    futures_value = prices[1].text if len(prices) > 1 else "N/A"


    tables = driver.find_elements(By.TAG_NAME, "table")

    df_list = []

    for tbl in tables:
        cls = tbl.get_attribute("class")

        if cls and "price-table" in cls:
            df = pd.read_html(StringIO(tbl.get_attribute("outerHTML")))[0]

            if isinstance(df.columns, pd.MultiIndex):
                df.columns = [
                    " ".join([str(c) for c in col if c])
                    for col in df.columns.values
                ]

            df = df[~df.astype(str).apply(
                lambda row: row.str.contains("デルタ|ガンマ|セータ|ベガ").any(),
                axis=1
            )]

            df_list.append(df)


    # =====================================================
    # 保存
    # =====================================================
    if df_list:

        combined_df = pd.concat(df_list, ignore_index=True)

        header_row = pd.DataFrame([{
            "日経平均株価": nikkei_value,
            "日経225先物": futures_value,
            "取得日時": now_jst.strftime("%Y-%m-%d %H:%M:%S")
        }])

        final_df = pd.concat([header_row, combined_df], ignore_index=True)

        final_df.to_excel(OUTPUT_FILE, index=False)

        print("💾 保存完了:", OUTPUT_FILE)

        # ★ここで呼ぶ（OK）
        upload_to_drive(OUTPUT_FILE, OUTPUT_FILE)

    else:
        print("⚠ テーブルなし")


finally:
    print("ブラウザ終了")
    driver.quit()

print(OUTPUT_FILE)
