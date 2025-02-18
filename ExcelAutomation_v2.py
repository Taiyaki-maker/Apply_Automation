#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Dec  2 17:11:27 2024

@author: gonzaresu
"""

import pandas as pd
import os
import requests
import re
import time
from bs4 import BeautifulSoup

def search_places_with_text(query, api_key, max_results=200, existing_names=None):
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
            name = place.get("name")
            if name in existing_names:
                print(f"Skipping already existing place: {name}")
                continue

            address = place.get("formatted_address")
            place_id = place.get("place_id")
            website = get_place_website(place_id, api_key)
            hours, closed_days = get_opening_hours(place_id, api_key)
            email = get_email_from_website(website) if website != "No website available" else None

            places_data.append({
                "name": name,
                "address": address,
                "website": website,
                "email": email,
                "opening_hours": hours,
                "closed_days": closed_days,
                "execution_flag": False
            })

        # 次のページトークン
        next_page_token = data.get("next_page_token")
        if not next_page_token:
            break
        params["pagetoken"] = next_page_token
        time.sleep(2)  # 推奨の待機時間

    return places_data[:max_results]

def get_place_website(place_id, api_key):
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {"place_id": place_id, "fields": "website", "key": api_key}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json().get("result", {}).get("website", "No website available")
    return "No website available"

def get_opening_hours(place_id, api_key):
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {"place_id": place_id, "fields": "opening_hours", "key": api_key}
    response = requests.get(url, params=params)
    
    if response.status_code == 200:
        hours = response.json().get("result", {}).get("opening_hours", {}).get("weekday_text", [])
        if hours:
            hours_dict = {}
            closed_days = []

            for day_hours in hours:
                day, time = day_hours.split(": ", 1)
                if time.strip().lower() == "closed":
                    closed_days.append(day)
                else:
                    hours_dict.setdefault(time.strip(), []).append(day[:3])

            opening_hours = []
            for time, days in hours_dict.items():
                day_range = ", ".join(days) if len(days) == 1 else f"{days[0]}-{days[-1]}"
                opening_hours.append(f"{day_range}: {time}")

            formatted_hours = ", ".join(opening_hours)
            formatted_closed = ", ".join(closed_days) if closed_days else "None"
            return formatted_hours, formatted_closed

    return "No hours available", "No closed days"

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

def save_to_excel(data, filename="resume/places_data.xlsx"):
    columns = ["name", "address", "website", "email", "opening_hours", "closed_days", "execution_flag"]
    
    if os.path.exists(filename):
        try:
            df_existing = pd.read_excel(filename)
        except Exception as e:
            print(f"Error reading Excel: {e}")
            df_existing = pd.DataFrame(columns=columns)
    else:
        df_existing = pd.DataFrame(columns=columns)

    filtered_data = [entry for entry in data if entry.get("email")]
    df_new = pd.DataFrame(filtered_data)
    df_new = df_new[~df_new["name"].isin(df_existing["name"])]
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined.to_excel(filename, index=False)

def load_existing_names(filename="places_data.xlsx"):
    if os.path.exists(filename):
        try:
            df = pd.read_excel(filename)
            return set(df["name"].dropna())
        except Exception as e:
            print(f"Error reading Excel: {e}")
    return set()

# 使用例
api_key = "GOOGLE MAPS API KEYS"
query = "Restaurant near CBD"

existing_names = load_existing_names()
places_data = search_places_with_text(query, api_key, existing_names=existing_names)
save_to_excel(places_data)
