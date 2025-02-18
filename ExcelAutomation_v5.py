#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jan  4 15:12:54 2025

@author: gonzaresu
"""

import pandas as pd
import os
import requests
import re
import time
import my_gmail_account as gmail


def normalize_name(name):
    """Normalize names by stripping whitespace and converting to lowercase"""
    return name.strip().lower() if name else ""


def get_email_from_website(url):
    """Retrieve the first email address found on a website"""
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?", response.text)
        return emails[0] if emails else None
    except requests.RequestException as e:
        return f"Error: {e}"


def get_place_website(place_id, api_key):
    """Fetch the website URL for a place using its Place ID"""
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {"place_id": place_id, "fields": "website", "key": api_key}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json().get("result", {}).get("website", "No website available")
    return "No website available"


def search_places_with_text(query, api_key, max_results=300, existing_names=None):
    """Search for places using a text query and collect data"""
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {"query": query, "key": api_key}
    places_data = []
    existing_names = existing_names or set()

    while len(places_data) < max_results:
        response = requests.get(url, params=params)
        if response.status_code != 200:
            print(f"API Error: {response.status_code}")
            break

        data = response.json()
        for place in data.get("results", []):
            name = normalize_name(place.get("name"))
            if name in existing_names:
                continue

            website = get_place_website(place.get("place_id"), api_key)
            email = get_email_from_website(website) if website != "No website available" else None
            places_data.append({
                "name": place.get("name"),
                "address": place.get("formatted_address"),
                "website": website,
                "email": email or "No email found"
            })
            existing_names.add(name)

        next_page_token = data.get("next_page_token")
        if not next_page_token:
            break
        params["pagetoken"] = next_page_token
        time.sleep(2)
        
    return places_data[:max_results]


def save_to_excel(data, filename="resume/places_data.xlsx"):
    """Save the collected place data to an Excel file"""
    success_data = [entry for entry in data if "@" in entry["email"]]
    failure_data = [entry for entry in data if "@" not in entry["email"]]
    
    for entry in success_data:
        entry["execution_flag"] = False  # 明示的に False を設定
            
    if os.path.exists(filename):
        with pd.ExcelWriter(filename, mode='a', engine="openpyxl", if_sheet_exists="overlay") as writer:
            # Read existing sheets or initialize empty DataFrames
            existing_success = pd.read_excel(filename, sheet_name="succeed") if "succeed" in pd.ExcelFile(filename).sheet_names else pd.DataFrame()
            existing_failed = pd.read_excel(filename, sheet_name="failed") if "failed" in pd.ExcelFile(filename).sheet_names else pd.DataFrame()

            # Combine new and existing data, dropping duplicates
            success_combined = pd.concat([existing_success, pd.DataFrame(success_data)], ignore_index=True).drop_duplicates(subset=["name", "address"])
            failure_combined = pd.concat([existing_failed, pd.DataFrame(failure_data)], ignore_index=True).drop_duplicates(subset=["name", "address"])

            # Write combined data back to respective sheets
            success_combined.to_excel(writer, index=False, sheet_name="succeed")
            failure_combined.to_excel(writer, index=False, sheet_name="failed")
    else:
        with pd.ExcelWriter(filename) as writer:
            pd.DataFrame(success_data).to_excel(writer, index=False, sheet_name="succeed")
            pd.DataFrame(failure_data).to_excel(writer, index=False, sheet_name="failed")


def load_existing_names(filename="resume/places_data.xlsx"):
    """Load existing place names from the Excel file to avoid duplication"""
    existing_names = set()
    if os.path.exists(filename):
        try:
            success_df = pd.read_excel(filename, sheet_name="succeed")
            failed_df = pd.read_excel(filename, sheet_name="failed")
            existing_names.update(success_df["name"].apply(normalize_name))
            existing_names.update(failed_df["name"].apply(normalize_name))
        except ValueError:
            pass  # Handle case where sheets don't exist
    return existing_names


# Example usage
query = "Cafe near Dandenong"

existing_names = load_existing_names()

places_data = search_places_with_text(query, gmail.api_key, existing_names=existing_names)
save_to_excel(places_data, filename="resume/places_data.xlsx")


