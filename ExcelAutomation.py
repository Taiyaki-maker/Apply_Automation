#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Nov 15 13:14:25 2024

@author: gonzaresu
"""

import requests
import pandas as pd
from bs4 import BeautifulSoup
import re
import os

def search_places_with_text(query, api_key):
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {
        "query": query,
        "key": api_key
    }
    
    response = requests.get(url, params=params)
    places_data = []
    
    if response.status_code == 200:
        data = response.json()
        places = data.get("results", [])
        
        for place in places:
            name = place.get("name")
            address = place.get("formatted_address")
            place_id = place.get("place_id")
            website = get_place_website(place_id, api_key)
            opening_hours = get_opening_hours(place_id, api_key)
            
            email = get_email_from_website(website) if website != "No website available" else None
            places_data.append({
                "name": name,
                "address": address,
                "website": website,
                "email": email,
                "opening_hours": opening_hours,
                "execution_flag": False
            })
    else:
        print("Failed to retrieve data:", response.status_code)
        
    return places_data


def get_place_website(place_id, api_key):
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {
        "place_id": place_id,
        "fields": "website",
        "key": api_key
    }
    
    response = requests.get(url, params=params)
    if response.status_code == 200:
        data = response.json()
        return data.get("result", {}).get("website", "No website available")
    return "No website available"


def get_opening_hours(place_id, api_key):
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {
        "place_id": place_id,
        "fields": "opening_hours",
        "key": api_key
    }
    
    response = requests.get(url, params=params)
    if response.status_code == 200:
        data = response.json()
        opening_hours = data.get("result", {}).get("opening_hours", {}).get("weekday_text", [])
        return ", ".join(opening_hours) if opening_hours else "No hours available"
    return "No hours available"


def get_email_from_website(url):
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        
        html_content = response.text
        emails = set(re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?", html_content))
        
        if emails:
            return list(emails)[0]
    except requests.RequestException as e:
        print(f"Failed to access {url}: {e}")
    return None


combined.to_excel(filename, index=False)

def save_to_excel(data, filename="Resume/places_data_real.xlsx"):
    # 必須列
    columns = ["name", "address", "website", "email", "opening_hours", "execution_flag"]

    # ファイルが存在する場合、既存データを読み込み
    if os.path.exists(filename):
        try:
            df_existing = pd.read_excel(filename, nrows=1)
            if set(columns).issubset(df_existing.columns):
                df_existing = pd.read_excel(filename)
            else:
                df_existing = pd.DataFrame(columns=columns)
        except Exception as e:
            print(f"Failed to read the existing Excel file: {e}")
            df_existing = pd.DataFrame(columns=columns)
    else:
        # ファイルが存在しない場合、新規データフレーム作成
        df_existing = pd.DataFrame(columns=columns)

    # 渡されたデータをデータフレームに変換
    if not data:  # dataが空の場合の対策
        print("No data provided to save.")
        return
    
    try:
        # メールアドレスが存在するデータだけフィルタリング
        filtered_data = [entry for entry in data if entry.get("email")]
        df_new = pd.DataFrame(filtered_data)
    except ValueError as e:
        print(f"Failed to create DataFrame: {e}")
        return

    # 必要な列が含まれているかチェック
    missing_columns = set(columns) - set(df_new.columns)
    if missing_columns:
        print(f"Missing columns in the new data: {missing_columns}")
        return

    # 重複除外
    existing_names = set(df_existing["name"]) if "name" in df_existing.columns else set()
    df_new = df_new[~df_new["name"].isin(existing_names)]

    # 結合して保存
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined.to_excel(filename, index=False)



# 使用例
api_key = "GOOGLE MAPS API KEYS"
query = "Restaurant near CBD"

places_data = search_places_with_text(query, api_key)
save_to_excel(places_data)
