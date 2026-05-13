#Get_Option_Data.py ｽﾏﾎ対応  2026/5/13PC版と完全一致





# ======================================================================
# Get_Option_Data.py（GitHub Actions 完全対応版）
# ======================================================================

import os
import time
from io import StringIO
from datetime import datetime, timedelta
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager


# ======================================================================
# 出力ファイル
# ======================================================================
now_jst = datetime.utcnow() + timedelta(hours=9)
timestamp_str = now_jst.strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = f"Get_Option_Data_{timestamp_str}.xlsx"


# ======================================================================
# Chrome設定（GitHub Actions用）
# ======================================================================
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")


# ======================================================================
# Chrome起動
# ======================================================================
print("⏳ Chrome 起動中...")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

print("✅ Chrome 起動成功")


# ======================================================================
# メイン処理
# ======================================================================
try:
    url = "https://www.jpx.co.jp/markets/derivatives/quotes/index.html"
    driver.get(url)
    print(f"✅ アクセス: {url}")

    # Cookie処理
    try:
        cookie_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".cc-btn.cc-dismiss"))
        )
        cookie_btn.click()
        print("✅ Cookie閉じた")
    except:
        print("⚠ Cookieなし")

    # オプションページ
    link = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "オプション価格情報"))
    )
    link.click()

    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(5)

    # 株価取得
    try:
        prices = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td.a-right.price-now"))
        )

        nikkei_value = prices[0].text.strip() if len(prices) > 0 else "N/A"
        futures_value = prices[1].text.strip() if len(prices) > 1 else "N/A"

        print("日経:", nikkei_value, "先物:", futures_value)

    except Exception as e:
        print("株価取得失敗:", e)
        nikkei_value = "N/A"
        futures_value = "N/A"


    # テーブル取得
    tables = driver.find_elements(By.TAG_NAME, "table")

    df_list = []

    for tbl in tables:
        cls = tbl.get_attribute("class")

        if cls and "price-table" in cls:
            try:
                df = pd.read_html(StringIO(tbl.get_attribute("outerHTML")))[0]

                # MultiIndex対策
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = [
                        " ".join([str(c) for c in col if c])
                        for col in df.columns.values
                    ]

                # Greek削除
                df = df[~df.astype(str).apply(
                    lambda row: row.str.contains("デルタ|ガンマ|セータ|ベガ").any(),
                    axis=1
                )]

                df_list.append(df)

            except Exception as e:
                print("DF変換失敗:", e)


    # ==================================================================
    # 保存
    # ==================================================================
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

    else:
        print("⚠ テーブルなし")


finally:
    print("ブラウザ終了")
    driver.quit()

print(OUTPUT_FILE)
