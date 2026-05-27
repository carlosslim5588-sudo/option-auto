# --- 2026/5/27 ---

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
OUTPUT_FILE = f"Get_Option_Data_{timestamp_str}.xlsx"

# --- Chrome起動設定 ---
chrome_options = Options()

chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

# PC版に近づける
chrome_options.add_argument(
    "--user-agent=Mozilla/5.0 "
    "(Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/136.0.0.0 Safari/537.36"
)

# Bot検知軽減
chrome_options.add_argument(
    "--disable-blink-features=AutomationControlled"
)

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

    # ページ読み込み待機
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script(
            "return document.readyState"
        ) == "complete"
    )

    # Cookieバナー対応
    try:
        cookie_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".cc-btn.cc-dismiss")
            )
        )

        cookie_btn.click()

        print("✅ Cookieバナーを閉じました。")

    except:
        print("⚠️ Cookieバナーは見つかりませんでした。")

    # 「オプション価格情報」をクリック
    link = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.LINK_TEXT, "オプション価格情報")
        )
    )

    print("🖱️ 「オプション価格情報」をクリックします。")

    # タブ数記録
    old_handles_count = len(driver.window_handles)

    # クリック
    link.click()

    # 新タブ待機
    WebDriverWait(driver, 20).until(
        lambda d: len(d.window_handles) > old_handles_count
    )

    # 新タブへ切替
    driver.switch_to.window(driver.window_handles[-1])

    print("✅ 新しいタブに切り替えました。")

    # 新ページ読み込み待機
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script(
            "return document.readyState"
        ) == "complete"
    )

    print("⏳ オプションページ描画待機中...")

    # table生成待機
    WebDriverWait(driver, 30).until(
        lambda d: len(
            d.find_elements(
                By.CSS_SELECTOR,
                "table.price-table tr"
            )
        ) > 30
    )

    # 念のため追加待機
    time.sleep(3)

    # デバッグ保存（完全一致確認用）
    with open("debug_page_source.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)

    # 株価情報取得
    try:

        print("🔍 日経平均株価と先物を取得します...")

        prices = WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "td.a-right.price-now")
            )
        )

        if len(prices) >= 2:

            nikkei_value = prices[0].text.strip()
            futures_value = prices[1].text.strip()

            print(
                f"✅ 取得結果: "
                f"日経平均={nikkei_value}, "
                f"先物={futures_value}"
            )

        else:
            raise Exception("価格要素が不足")

    except Exception as e:

        print(f"❌ 株価情報取得失敗: {e}")

        nikkei_value = "N/A"
        futures_value = "N/A"

    # table取得
    tables = WebDriverWait(driver, 30).until(
        lambda d: d.find_elements(
            By.CSS_SELECTOR,
            "table.price-table"
        )
    )

    df_list = []

    for tbl in tables:

        try:

            html = tbl.get_attribute("outerHTML")

            df = pd.read_html(
                StringIO(html)
            )[0]

            # MultiIndex flatten
            if isinstance(df.columns, pd.MultiIndex):

                df.columns = [
                    ' '.join(
                        [str(c) for c in col if c]
                    )
                    for col in df.columns.values
                ]

            # Greeks除外
            mask = df.astype(str).apply(
                lambda col: col.str.contains(
                    "デルタ|ガンマ|セータ|ベガ",
                    regex=True
                )
            ).any(axis=1)

            df = df[~mask]

            df_list.append(df)

        except Exception as e:

            print(f"❌ DataFrame変換失敗: {e}")

    if df_list:

        combined_df = pd.concat(
            df_list,
            ignore_index=True
        )

        timestamp = now_jst.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        header_row = pd.DataFrame([{
            "日経平均株価": nikkei_value,
            "日経225先物": futures_value,
            "取得日時": timestamp
        }])

        final_df = pd.concat(
            [header_row, combined_df],
            ignore_index=True
        )

        # 完全一致確認用CSV保存
        final_df.to_csv(
            "debug_compare.csv",
            index=False,
            encoding="utf-8-sig"
        )

        # Excel保存
        final_df.to_excel(
            OUTPUT_FILE,
            index=False
        )

        print(f"💾 保存完了: {OUTPUT_FILE}")

        print(f"行数: {final_df.shape[0]}")
        print(f"列数: {final_df.shape[1]}")

    else:

        print("⚠️ テーブル取得失敗")

        driver.save_screenshot(
            "table_error.png"
        )

        raise FileNotFoundError(
            "データテーブル取得失敗"
        )

finally:

    print("ブラウザ終了")

    driver.quit()


# GitHub Actions用ダミー
def upload_to_drive(local_file, drive_file):
    pass


upload_to_drive(
    OUTPUT_FILE,
    OUTPUT_FILE
)
