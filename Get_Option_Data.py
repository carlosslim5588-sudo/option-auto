# --- 2026/8/13 ---

import time
from io import StringIO
from datetime import datetime, timedelta, timezone
import pandas as pd
import numpy as np
import sys
import io
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 出力ファイル名 ---
now_jst = datetime.now(timezone.utc) + timedelta(hours=9)
timestamp_str = now_jst.strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = f"Get_Option_Data_{timestamp_str}.xlsx"

# --- Chrome起動設定 ---
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")

print("⏳ Chrome起動中...")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

try:
    url = "https://www.jpx.co.jp/markets/derivatives/quotes/index.html"
    driver.get(url)
    
    # Cookieバナー対応
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".cc-btn.cc-dismiss"))).click()
    except:
        pass

    # ページ遷移
    link = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.LINK_TEXT, "オプション価格情報")))
    link.click()
    driver.switch_to.window(driver.window_handles[-1])

    # 待機処理
    WebDriverWait(driver, 90).until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "table.price-table tr")) > 30)

    # 株価情報取得
    try:
        prices = WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td.a-right.price-now")))
        nikkei_value = prices[0].text.strip()
        futures_value = prices[1].text.strip()
    except:
        nikkei_value = "N/A"
        futures_value = "N/A"

    # テーブル取得とフィルタリング
    tables = driver.find_elements(By.CSS_SELECTOR, "table.price-table")
    df_list = []
    for tbl in tables:
        try:
            html = tbl.get_attribute("outerHTML")
            df = pd.read_html(StringIO(html))[0]
            
            # MultiIndex flatten
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = [" ".join([str(c) for c in col if c]) for col in df.columns.values]
            
            # ★修正箇所1: フィルタリング処理（8/13版と完全一致）
            df = df[~df.astype(str).apply(lambda row: row.str.contains("デルタ|ガンマ|セータ|ベガ").any(), axis=1)]
            df_list.append(df)
        except Exception as e:
            print(f"❌ テーブル取得失敗: {e}")

    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        timestamp = now_jst.strftime("%Y-%m-%d %H:%M:%S")

        # ★修正箇所2: ヘッダー作成・結合（8/13版と完全一致）
        header_row = pd.DataFrame([{
            "日経平均株価": nikkei_value,
            "日経225先物": futures_value,
            "取得日時": timestamp
        }])
        
        final_df = pd.concat([header_row, combined_df], ignore_index=True)

        # Excel保存
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"💾 保存完了: {OUTPUT_FILE}")
    else:
        print("⚠️ データが見つかりませんでした。")

finally:
    print("ブラウザ終了")
    driver.quit()
