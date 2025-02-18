#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 25 16:15:10 2024

@author: gonzaresu
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# ChromeDriverのパス
chromedriver_path = "/Users/gonzaresu/Documents/chromedriver"
service = Service(chromedriver_path)

# WebDriverの初期化
driver = webdriver.Chrome(service=service)

try:
    # Indeedのトップページにアクセス
    driver.get("https://www.indeed.com/")
    time.sleep(2)  # ページが完全に読み込まれるまで待機

    # 検索ボックスを探してキーワードを入力
    what_input = driver.find_element(By.ID, "text-input-what")
    what_input.send_keys("Software Engineer")  # 職種を指定

    where_input = driver.find_element(By.ID, "text-input-where")
    where_input.clear()  # デフォルトの場所をクリア
    where_input.send_keys("Remote")  # リモートジョブを検索
    where_input.send_keys(Keys.RETURN)
    time.sleep(3)

    wait = WebDriverWait(driver, 10)  # 最大10秒待機
    # ジョブリストを取得
    job_listings = driver.find_elements(By.CLASS_NAME, "job_seen_beacon")
    print(f"見つかった求人の数: {len(job_listings)}")

    # 最初のジョブをクリック
    if job_listings:
        #job_listings[0].click()
        wait.until(EC.element_to_be_clickable(job_listings[0])).click()
        time.sleep(3)

        # 新しいタブに切り替え
        driver.switch_to.window(driver.window_handles[-1])

        # "Apply Now"ボタンを探してクリック
        try:
            apply_button = driver.find_element(By.CLASS_NAME, "ia-IndeedApplyButton")
            apply_button.click()
            time.sleep(2)

            # アプライフォームに情報を入力（カスタマイズが必要）
            print("アプライを進行中...")

        except Exception as e:
            print(f"アプライボタンが見つからないか、エラー: {e}")
        
        # タブを閉じて元に戻る
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

finally:
    # ブラウザを閉じる
    driver.quit()