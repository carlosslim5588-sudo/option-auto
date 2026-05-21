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

# --- 出力ファイル名 ---
now_jst = datetime.utcnow() + timedelta(hours=9)
timestamp_str = now_jst.strftime("%Y%m%d_%H%M%S")
# GitHubの環境で動くように保存先をシンプルに変更
OUTPUT_FILE = f"Get_Option_Data_{timestamp_str}.xlsx"

# --- Chrome起動設定 ---
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

print("⏳ Chrome起動中...")
driver_path = ChromeDriverManager().install()
driver = webdriver.Chrome(
    service=Service(driver_path),
    options=chrome_options
)
print("✅ Chrome起動成功")

try:
    url = "https://www.jpx.co.jp/markets/derivatives/quotes/index.html"
    driver.get(url)
    print(f"✅ 開始ページ {url} に移動しました。")

    # Cookieバナー対応
    try:
        cookie_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".cc-btn.cc-dismiss"))
        )
        cookie_btn.click()
        print("✅ Cookieバナーを閉じました。")
    except:
        print("⚠️ Cookieバナーは見つかりませんでした。")

    # 「オプション価格情報」をクリック
    link = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "オプション価格情報"))
    )
    print("🖱️ 「オプション価格情報」のリンクをクリックします。")
    link.click()

    driver.switch_to.window(driver.window_handles[-1])
    print("✅ 新しいタブに切り替えました。")

    time.sleep(5)

    # 株価情報を取得
    try:
        print("🔍 日経平均株価と先物を取得します...")
        prices = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "td.a-right.price-now"))
        )
        if len(prices) >= 2:
            nikkei_value = prices[0].text.strip()
            futures_value = prices[1].text.strip()
            print(f"✅ 取得結果: 日経平均={nikkei_value}, 先物={futures_value}")
        else:
            raise Exception("価格要素が足りません")
    except Exception as e:
        print(f"❌ 株価情報の取得失敗: {e}")
        nikkei_value = "N/A"
        futures_value = "N/A"

    # オプション価格テーブルを取得
    tables = driver.find_elements(By.TAG_NAME, "table")
    df_list = []
    for tbl in tables:
        cls = tbl.get_attribute("class")
        if cls and "price-table" in cls:
            try:
                df = pd.read_html(StringIO(tbl.get_attribute("outerHTML")))[0]

                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = [' '.join([str(c) for c in col if c]) for col in df.columns.values]

                df = df[~df.astype(str).apply(lambda row: row.str.contains("デルタ|ガンマ|セータ|ベガ").any(), axis=1)]

                df_list.append(df)
            except Exception as e:
                print(f"❌ DataFrame変換失敗: {e}")

    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        timestamp = now_jst.strftime("%Y-%m-%d %H:%M:%S")
        header_row = pd.DataFrame([{
            "日経平均株価": nikkei_value,
            "日経225先物": futures_value,
            "取得日時": timestamp
        }])

        final_df = pd.concat([header_row, combined_df], ignore_index=True)

        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"💾 保存完了: {OUTPUT_FILE}")
    else:
        print("⚠ 保存対象のテーブルが見つかりませんでした。")

finally:
    print("ブラウザ終了")
    driver.quit()

# エラーの原因になっていた関数を、エラーが出ない「何もしない関数」に書き換え
def upload_to_drive(local_file, drive_file):
    print("Google Driveへのアップロードはスキップします（GitHubに保存されます）")
    pass

# 元のコードの162行目に残っている呼び出しを安全に処理
upload_to_drive(OUTPUT_FILE, OUTPUT_FILE)
