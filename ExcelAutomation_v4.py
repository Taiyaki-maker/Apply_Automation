#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Dec  7 15:42:56 2024

@author: gonzaresu
"""

import pandas as pd
import os
import requests
import re
import time


def normalize_name(name):
    """Normalize names by stripping whitespace and converting to lowercase"""
    return name.strip().lower() if name else ""


def search_places_with_text(query, api_key, max_results=300, existing_names=None):
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {"query": query, "key": api_key}
    places_data = []

    if existing_names is None:
        existing_names = set()

    while len(places_data) < max_results:
        response = requests.get(url, params=params)
        if response.status_code != 200:
            print(f"Error: {response.status_code}")
            break

        data = response.json()
        places = data.get("results", [])

        for place in places:
            name = normalize_name(place.get("name"))
            if name in existing_names:
                continue

            address = place.get("formatted_address")
            place_id = place.get("place_id")
            website = get_place_website(place_id, api_key)
            email = get_email_from_website(website) if website != "No website available" else None

            places_data.append({
                "name": place.get("name"),
                "address": address,
                "website": website,
                "email": email
            })

            existing_names.add(name)

        next_page_token = data.get("next_page_token")
        if not next_page_token:
            break
        params["pagetoken"] = next_page_token
        time.sleep(2)
        
    return places_data[:max_results]


def get_place_website(place_id, api_key):
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {"place_id": place_id, "fields": "website", "key": api_key}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json().get("result", {}).get("website", "No website available")
    return "No website available"


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

'''

def save_to_excel(data, filename="resume/places_data.xlsx"):
    success_data = [entry for entry in data if entry.get("email")]
    failure_data = [entry for entry in data if not entry.get("email")]

    # 1枚目のデータにのみ "execution_flag" 列を追加
    for entry in success_data:
        entry["execution_flag"] = False  # エクセルでは FALSE と表示される

    if os.path.exists(filename):
        with pd.ExcelWriter(filename, mode='a', engine="openpyxl", if_sheet_exists="overlay") as writer:
            success_df = pd.read_excel(filename, sheet_name=0)
            failure_df = pd.read_excel(filename, sheet_name=1)
            success_combined = pd.concat([success_df, pd.DataFrame(success_data)], ignore_index=True).drop_duplicates()
            failure_combined = pd.concat([failure_df, pd.DataFrame(failure_data)], ignore_index=True).drop_duplicates()
            success_combined.to_excel(writer, index=False, sheet_name="succeed")
            failure_combined.to_excel(writer, index=False, sheet_name="failed")
    else:
        with pd.ExcelWriter(filename) as writer:
            pd.DataFrame(success_data).to_excel(writer, index=False, sheet_name="succeed")
            pd.DataFrame(failure_data).to_excel(writer, index=False, sheet_name="failed")'''
    
def save_to_excel(data, filename="resume/places_data.xlsx"):
    success_data = [entry for entry in data if entry.get("email")]
    failure_data = [entry for entry in data if not entry.get("email")]

    # 1枚目のデータにのみ "execution_flag" 列を追加
    for entry in success_data:
        entry["execution_flag"] = False  # エクセルでは FALSE と表示される

    if os.path.exists(filename):
        with pd.ExcelWriter(filename, mode='a', engine="openpyxl", if_sheet_exists="overlay") as writer:
            # 既存のデータを読み込む
            try:
                success_df = pd.read_excel(filename, sheet_name="succeed")
                failure_df = pd.read_excel(filename, sheet_name="failed")
            except ValueError:
                success_df = pd.DataFrame()
                failure_df = pd.DataFrame()

            # 既存データの行数を取得
            #success_start_row = len(success_df) + 1 if not success_df.empty else 0
            #failure_start_row = len(failure_df) + 1 if not failure_df.empty else 0]
            success_start_row = 1
            failure_start_row = 1

            # 新しいデータを結合して重複を排除
            success_combined = pd.concat([success_df, pd.DataFrame(success_data)], ignore_index=True).drop_duplicates()
            failure_combined = pd.concat([failure_df, pd.DataFrame(failure_data)], ignore_index=True).drop_duplicates()

            # データをそれぞれのシートに追加
            success_combined.to_excel(writer, index=False, sheet_name="succeed", startrow=success_start_row)
            failure_combined.to_excel(writer, index=False, sheet_name="failed", startrow=failure_start_row)
    else:
        with pd.ExcelWriter(filename) as writer:
            pd.DataFrame(success_data).to_excel(writer, index=False, sheet_name="succeed")
            pd.DataFrame(failure_data).to_excel(writer, index=False, sheet_name="failed")


def load_existing_names(filename="resume/places_data.xlsx"):
    existing_names = set()
    if os.path.exists(filename):
        # 読み込み時にシートが存在するか確認し、初期化
        try:
            success_df = pd.read_excel(filename, sheet_name=0)
            failure_df = pd.read_excel(filename, sheet_name=1)
        except ValueError as e:
            print(f"Error reading sheets: {e}")
            success_df = pd.DataFrame(columns=["name"])
            failure_df = pd.DataFrame(columns=["name"])
        
        # データフレームが空の場合に対応
        if "name" in success_df.columns:
            existing_names.update(normalize_name(name) for name in success_df["name"].dropna())
        if "name" in failure_df.columns:
            existing_names.update(normalize_name(name) for name in failure_df["name"].dropna())
    return existing_names


# Example usage
api_key = "GOOGLE MAPS API KEYS"
query = "Cafe near Murrumbeena"

existing_names = load_existing_names()

#places_data = search_places_with_text(query, api_key, existing_names=existing_names)
#save_to_excel(places_data, filename="resume/places_data.xlsx")
