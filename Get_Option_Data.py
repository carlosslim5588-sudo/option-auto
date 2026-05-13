#Get_Option_Data.py ｽﾏﾎ対応  2026/5/13PC版と完全一致





# ======================================================================
# 【Colab 2025 対応】100％動作する Get_Option_Data.py（完全移植版）
# ======================================================================

# --- 破損プロセス停止 ---
!killall -q chrome chromedriver || true
!mkdir -p /content/drive/MyDrive/OptionData
# --- Chrome 最新版インストール ---
!wget -q https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
!apt-get install -y ./google-chrome-stable_current_amd64.deb > /dev/null

# --- Chrome のバージョン確認 ---
!google-chrome --version

# --- WebDriver Manager で一致する ChromeDriver を取得 ---
!pip install webdriver-manager selenium pandas openpyxl -q

from webdriver_manager.chrome import ChromeDriverManager
driver_path = ChromeDriverManager().install()
print("ChromeDriver Path:", driver_path)

# --- Google Drive ---


# --- ライブラリ ---
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

# --- 出力ファイル名 ---
now_jst = datetime.utcnow() + timedelta(hours=9)
timestamp_str = now_jst.strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = f"Get_Option_Data_{timestamp_str}.xlsx"


# --- Chrome options ---
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

# --- 起動（ここが最重要：正しい ChromeDriver を指定） ---
print("⏳ Chrome を起動しています...")
driver = webdriver.Chrome(
    service=Service(driver_path),
    options=chrome_options
)
print("✅ Chrome Driver 起動成功")

# ======================================================================
# ▼▼▼ ここから下は PC版ロジックを一切改変せずそのまま実行 ▼▼▼
# ======================================================================

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

    # Colab 安定化
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

        # header_row(3列) + combined_df を結合して最終 DataFrame に
        final_df = pd.concat([header_row, combined_df], ignore_index=True)

        # --- Excel 保存 ---
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"💾 Excelファイルを保存しました: {OUTPUT_FILE}")

    else:
        print("⚠ 保存対象のテーブルが見つかりませんでした。")



finally:
    print("✅ ブラウザを閉じました。")
    driver.quit()


print(OUTPUT_FILE)
