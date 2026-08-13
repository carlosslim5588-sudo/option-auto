# --- セル2：JPXオプションデータ取得＆自動ダウンロード ---
import time
from io import StringIO
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.colab import files  # Colab専用：ファイルダウンロード用

# --- 出力ファイル名 ---
now_jst = datetime.utcnow() + timedelta(hours=9)
timestamp_str = now_jst.strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = f"Get_Option_Data_{timestamp_str}.xlsx"

# --- Colab用 Chrome起動設定 ---
chrome_options = Options()
chrome_options.add_argument('--headless')                 # 画面を表示しない（Colab必須）
chrome_options.add_argument('--no-sandbox')               # セキュリティ制限解除（Colab必須）
chrome_options.add_argument('--disable-dev-shm-usage')    # メモリ不足対策（Colab必須）
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument(
    "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
)

print("⏳ Colab内でChromeを起動中...")

# ColabにインストールされたChromiumを使用
service = Service('/usr/bin/chromedriver')
driver = webdriver.Chrome(service=service, options=chrome_options)

print("✅ Chrome起動成功")

try:
    url = "https://www.jpx.co.jp/markets/derivatives/quotes/index.html"
    driver.get(url)
    print(f"✅ 開始ページ {url} に移動しました。")

    # ページ読み込み待機
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    # Cookieバナー対応
    try:
        cookie_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".cc-btn.cc-dismiss"))
        )
        cookie_btn.click()
        print("✅ Cookieバナーを閉じました。")
    except:
        pass # バナーがなければ無視

    # 「オプション価格情報」をクリック
    link = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "オプション価格情報"))
    )
    print("🖱️ 「オプション価格情報」をクリックします。")

    old_handles_count = len(driver.window_handles)
    link.click()

    # 新タブ待機と切り替え
    WebDriverWait(driver, 30).until(
        lambda d: len(d.window_handles) > old_handles_count
    )
    driver.switch_to.window(driver.window_handles[-1])
    print("✅ 新しいタブに切り替えました。")

    # 新ページ読み込み待機
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    print("⏳ オプションページテーブル描画待機中...")
    
    # テーブル生成待機
    WebDriverWait(driver, 90).until(
        lambda d: len(d.find_elements(By.CSS_SELECTOR, "table.price-table tr")) > 30
    )
    print("✅ table生成完了")
    time.sleep(3) # 描画安定化のための追加待機

    # 株価情報取得
    try:
        print("🔍 日経平均株価と先物を取得します...")
        prices = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td.a-right.price-now"))
        )
        if len(prices) >= 2:
            nikkei_value = prices[0].text.strip()
            futures_value = prices[1].text.strip()
            print(f"✅ 取得結果: 日経平均={nikkei_value}, 先物={futures_value}")
        else:
            raise Exception("価格要素が不足")
    except Exception as e:
        print(f"❌ 株価情報取得失敗: {e}")
        nikkei_value = "N/A"
        futures_value = "N/A"

    # table取得
    tables = WebDriverWait(driver, 60).until(
        lambda d: d.find_elements(By.CSS_SELECTOR, "table.price-table")
    )

    df_list = []
    for tbl in tables:
        try:
            html = tbl.get_attribute("outerHTML")
            df = pd.read_html(StringIO(html))[0]
            
            # マルチインデックスの平坦化
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = [' '.join([str(c) for c in col if c]) for col in df.columns.values]
            
            # ツール（DB作成前処理Tool.py）の仕様に合わせるため、Greeks行は削除せずにそのまま保持
            df_list.append(df)
            
        except Exception as e:
            print(f"❌ DataFrame変換失敗: {e}")

    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        timestamp = now_jst.strftime("%Y-%m-%d %H:%M:%S")

        # ツール仕様に合わせた1行目（ヘッダー行代わり）の作成
        col_count = combined_df.shape[1]
        session_row = [nikkei_value, futures_value, timestamp] + [np.nan] * (col_count - 3)
        header_row = pd.DataFrame([session_row], columns=combined_df.columns)

        final_df = pd.concat([header_row, combined_df], ignore_index=True)

        # Excel保存（ツールがヘッダーなしで読み込むため header=False に設定）
        final_df.to_excel(OUTPUT_FILE, index=False, header=False)
        print(f"💾 保存完了: {OUTPUT_FILE}")
        
        # 完成したファイルをローカルに自動ダウンロード
        print("📥 ファイルをダウンロードします...")
        files.download(OUTPUT_FILE)

    else:
        print("⚠️ テーブル取得失敗")

finally:
    print("ブラウザ終了")
    driver.quit()
