import requests
import json
import urllib3
import re
import time
import pandas as pd

# Suppress SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- 1. SETUP ---
FILE_NAME = "Toyota parts.xlsx"
OUTPUT_FILE = "Toyota_parts_WITH_MODELS.xlsx"

with open("cookies.json", "r") as f:
    cookies_list = json.load(f)
cookies = {cookie['name']: cookie['value'] for cookie in cookies_list}

headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://epcwebserver.tmhg.jp/home"
}

api_url = "https://epcwebserver.tmhg.jp/partnumberapplicationtomodels/list"

def get_fresh_token():
    print("Getting fresh CSRF token...")
    init_url = "https://epcwebserver.tmhg.jp/partnumberapplicationtomodels?series=&pnc=&partnumber=&partname="
    res = requests.get(init_url, cookies=cookies, headers=headers, verify=False)
    match = re.search(r'name="csrf-token" content="(.*?)"', res.text)
    if match:
        return match.group(1)
    return None

def search_part(part_number, token):
    # NEW RULE: Strip any prefix that ends with a slash (e.g., 058/, 328/, D03/)
    search_number = part_number
    if "/" in search_number:
        # Split at the first slash and take the second part
        parts = search_number.split("/", 1)
        # Only strip if the second part looks like a part number (contains letters/numbers)
        if len(parts) > 1 and len(parts[1]) > 3:
            search_number = parts[1]

    payload = {
        "partnumber": search_number,
        "series": "",
        "pnc": "",
        "partname": "",
        "_token": token
    }
    try:
        res = requests.get(api_url, params=payload, cookies=cookies, headers=headers, verify=False)
        data = res.json()
        
        if data.get("no_data") or not data.get("seriesCodes"):
            return "Not Found"
        
        models = [series["Title"] for series in data["seriesCodes"]]
        return " | ".join(models)
    except Exception as e:
        return f"Error: {e}"

def main():
    token = get_fresh_token()
    if not token:
        print("Could not get CSRF token. Are cookies expired? Run grab_cookies.py again.")
        return

    print(f"Loading {FILE_NAME}...")
    sheets = pd.read_excel(FILE_NAME, sheet_name=None)
    
    for sheet_name, df in sheets.items():
        if "BT" in sheet_name.upper():
            print(f"\nSkipping sheet: {sheet_name} (BT parts are in a different catalog)")
            continue

        print(f"\n--- Processing Sheet: {sheet_name} ---")
        
        part_col = None
        for col in df.columns:
            if "Part No" in str(col):
                part_col = col
                break
                
        if not part_col:
            print(f"Skipping sheet {sheet_name}: Could not find 'Part No' column.")
            continue
            
        total_rows = len(df)
        print(f"Found {total_rows} parts to search in this sheet.")
        
        model_col_name = "Model"
        if model_col_name not in df.columns:
            df[model_col_name] = ""
        
        df[model_col_name] = df[model_col_name].astype(str)

        for index, row in df.iterrows():
            if pd.isna(row[part_col]):
                continue
                
            part_number = str(row[part_col]).strip()
            
            if pd.notna(row[model_col_name]) and row[model_col_name] != "" and row[model_col_name] != "nan":
                continue

            print(f"[{index+1}/{total_rows}] Searching: {part_number} ... ", end="", flush=True)
            
            models_found = search_part(part_number, token)
            df.at[index, model_col_name] = models_found
            
            if models_found == "Not Found":
                print("❌ Not Found")
            else:
                print(f"✅ Found")
                
            time.sleep(1.5) 
            
            if (index + 1) % 50 == 0:
                print("\nRefreshing token to keep session alive...")
                token = get_fresh_token()

        sheets[sheet_name] = df

    print(f"\nAll sheets processed! Saving to {OUTPUT_FILE}...")
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    print("Done! You can now open the new Excel file.")

if __name__ == "__main__":
    main()
