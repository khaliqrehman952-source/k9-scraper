#!/usr/bin/env python3
"""
k9_api_scraper.py
Flask wrapper for your existing k9_scraper logic (uses Selenium exactly as in your script).
Routes:
 - GET  /                -> serves control panel HTML (templates/index.html)
 - POST /manual_scrape   -> { "breed": "Golden"|"Labrador", "name": "<dog name>", "resume_from": "<optional>" }
 - POST /rescrape_range  -> { "start_row": int, "end_row": int }
 - POST /scrape_single   -> { "url": "<full dog profile URL>" }
"""

import os
import re
import json
import time
import random
import io
import sys
import threading
import traceback
from contextlib import redirect_stdout
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS

import pandas as pd
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import undetected_chromedriver as uc

# 🧠 Keep latest scraped dogs in memory (live view)
LIVE_DOGS = []

# -------------------------
# CONFIGURATION (copied exactly)
# -------------------------
BASE_URL = "http://www.k9data.com/"
SAVE_DIR = "k9_dog_data"
WAIT_TIME = 12

FIELDS = [
    "Call name", "Gender", "Honorifics", "Country of origin",
    "Country of residence", "Registration", "Breeder", "Owner",
    "Web site", "Hip clearance", "Eye clearance", "Heart clearance",
    "Elbow clearance", "Cause of death"
]

SAVE_DIR = "k9_dog_data"
csv_path = os.path.join(SAVE_DIR, "k9_data.csv")
xlsx_path = os.path.join(SAVE_DIR, "k9_data.xlsx")
json_path = os.path.join(SAVE_DIR, "k9_data.json")

os.makedirs(SAVE_DIR, exist_ok=True)


if os.path.exists(csv_path):
    import shutil

    # --- Backup before modifying anything ---
    backup_path = csv_path.replace(".csv", "_backup_before_fix.csv")
    if not os.path.exists(backup_path):
        shutil.copy(csv_path, backup_path)
        print(f"🧾 Backup created at: {backup_path}")

    # --- Load the CSV normally ---
    existing_df = pd.read_csv(csv_path, dtype=str, encoding="utf-8-sig").fillna("N/A")

    # --- Handle duplicate column names manually ---
    # Pandas may rename duplicate columns like 'URL.1', 'URL.2', etc.
    dupe_cols = [c for c in existing_df.columns if c.startswith("URL.")]

    if dupe_cols:
        print(f"⚙️ Found duplicate URL columns: {dupe_cols}")
        for col in dupe_cols:
            existing_df["URL"] = existing_df["URL"].combine_first(existing_df[col])
        existing_df.drop(columns=dupe_cols, inplace=True)

        # --- Save cleaned data back immediately ---
        existing_df = existing_df.fillna("N/A")
        existing_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        existing_df.to_excel(xlsx_path, index=False)
        existing_df.to_json(json_path, orient="records", indent=2)
        print("✅ Duplicate URL columns removed and data saved cleanly!")

else:
    existing_df = pd.DataFrame(
        columns=["URL", "U-ID", "Name", "Date of Birth", "Date of Death", "Sire", "Dam"] + FIELDS
    )


def clean_url(url: str) -> str:
    """Safely clean and normalize dog profile URLs without regex escape issues."""
    if not isinstance(url, str):
        return "N/A"

    url = url.replace("\\", "/")
    url = re.sub(r'(?<!:)//+', "//", url)
    url = url.replace("http:////", "http://").replace("https:////", "https://")
    url = url.strip()

    if not url.startswith("http"):
        if "k9data.com" in url:
            url = "http://www.k9data.com" + ("/" + url.lstrip("/") if not url.startswith("/") else url)

    return url

    # -------------------------
# ✅ OPTIONAL: Clean all existing URLs once across CSV/XLSX/JSON
# -------------------------
try:
    if not existing_df.empty:
        print("🧹 Cleaning existing URLs in all files...")
        existing_df["URL"] = existing_df["URL"].apply(clean_url)
        existing_df = existing_df.fillna("N/A")
        existing_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        existing_df.to_excel(xlsx_path, index=False)
        existing_df.to_json(json_path, orient="records", indent=2)
        print("✅ URL cleanup applied successfully!")
except Exception as e:
    print(f"⚠️ URL cleanup skipped: {e}")




for col in ["URL", "U-ID", "Name", "Date of Birth", "Date of Death", "Sire", "Dam"] + FIELDS:
    if col not in existing_df.columns:
        existing_df[col] = "N/A"

def build_index_map(df):
    df["U-ID"] = df["U-ID"].fillna("N/A").astype(str).str.strip()
    return {uid: idx for idx, uid in df["U-ID"].items() if uid and uid != "N/A"}

index_map = build_index_map(existing_df)

# -------------------------
# URL CLEANUP FIX
# -------------------------






# -------------------------
# APPEND / UPDATE FUNCTION (same as your file)
# -------------------------
def append_or_update(df, mapping, new_row, insert_after=None):
    uid = str(new_row.get("U-ID", "N/A")).strip()
    url = str(new_row.get("URL", "N/A")).strip().lower()

    if uid != "N/A" and uid in mapping:
        idx = mapping[uid]
        for k, v in new_row.items():
            if k not in df.columns:
                df[k] = "N/A"
            new_val = str(v).strip()
            old_val = str(df.at[idx, k]).strip()
            if new_val not in ["", "N/A"]:
                df.at[idx, k] = new_val
        return df, mapping

    for j, row_url in df["URL"].astype(str).items():
        if row_url.lower() == url:
            if uid != "N/A":
                mapping[uid] = j
            for k, v in new_row.items():
                if k not in df.columns:
                    df[k] = "N/A"
                if str(v).strip() not in ["", "N/A"]:
                    df.at[j, k] = str(v).strip()
            return df, mapping

    append_row = {c: "N/A" for c in df.columns}
    for k, v in new_row.items():
        append_row[k] = str(v).strip() if v else "N/A"

    df = pd.concat([df, pd.DataFrame([append_row])], ignore_index=True)
    new_index = len(df) - 1

    if uid != "N/A":
        mapping[uid] = new_index
    return df.reset_index(drop=True), mapping

# -------------------------
# Pedigree extraction (same)
# -------------------------
def extract_sire_dam_from_pedigree(soup):
    sire, dam, sire_url, dam_url = "N/A", "N/A", "N/A", "N/A"
    try:
        ped_label = soup.find("strong", string=re.compile(r"Pedigree", re.I))
        if not ped_label:
            return sire, sire_url, dam, dam_url

        ped_table = ped_label.find_next("table")
        if not ped_table:
            return sire, sire_url, dam, dam_url

        top_rows = ped_table.find_all("tr", recursive=False)

        if len(top_rows) >= 1:
            sire_row = top_rows[0]
            sire_link = sire_row.find("a", href=re.compile(r"pedigree\.asp", re.I))
            if sire_link:
                sire = sire_link.get_text(strip=True)
                sire_url = BASE_URL + sire_link["href"]

        if len(top_rows) >= 2:
            dam_row = top_rows[1]
            dam_link = dam_row.find("a", href=re.compile(r"pedigree\.asp", re.I))
            if dam_link:
                dam = dam_link.get_text(strip=True)
                dam_url = BASE_URL + dam_link["href"]

        if sire == "N/A" or dam == "N/A":
            all_links = ped_table.find_all("a", href=re.compile(r"pedigree\.asp", re.I))
            names = [a.get_text(strip=True) for a in all_links if a.get_text(strip=True)]
            hrefs = [a["href"] for a in all_links]
            if sire == "N/A" and names:
                sire = names[0]
                sire_url = BASE_URL + hrefs[0]
            if dam == "N/A" and len(names) > 7:
                dam = names[-7]
                dam_url = BASE_URL + hrefs[-7]
    except Exception as e:
        print("⚠️ Pedigree parse error:", e)

    return sire, sire_url, dam, dam_url

# -------------------------
# Scraper functions (adapted to be called from Flask)
# All logic is preserved exactly from your script.
# -------------------------
def _create_driver():
    options = uc.ChromeOptions()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = uc.Chrome(options=options, use_subprocess=True)
    return driver

def save_progress():
    global existing_df
    existing_df = existing_df.drop_duplicates(subset=["U-ID"], keep="last").reset_index(drop=True)
    existing_df = existing_df.fillna("N/A")
    existing_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    existing_df.to_excel(xlsx_path, index=False)
    existing_df.to_json(json_path, orient="records", indent=2)

def scrape_single_dog(dog_url):
    global existing_df, index_map

    if not dog_url.startswith("http"):
        print("❌ Invalid URL!")
        return {"status": "error", "message": "Invalid URL"}

    print(f"\n➡️ Scraping single dog: {dog_url}")

    driver = _create_driver()
    wait = WebDriverWait(driver, WAIT_TIME)

    try:
        driver.get(dog_url)
        wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")

        try:
            raw_name = driver.find_element(By.XPATH, "//strong/font[@size='4']").text.strip()
        except Exception:
            raw_name = "Unknown"

        dog_data = {"URL": clean_url(dog_url), "U-ID": re.search(r"ID=(\d+)", dog_url).group(1) if "ID=" in dog_url else "N/A"}

        if "(" in raw_name and ")" in raw_name:
            name_part, date_part = raw_name.split("(", 1)
            dog_data["Name"] = name_part.strip()
            date_part = date_part.strip(")")
            if "-" in date_part:
                dob, dod = map(str.strip, date_part.split("-", 1))
                dog_data["Date of Birth"], dog_data["Date of Death"] = dob, dod
            else:
                dog_data["Date of Birth"], dog_data["Date of Death"] = date_part, "N/A"
        else:
            dog_data["Name"], dog_data["Date of Birth"], dog_data["Date of Death"] = raw_name, "N/A", "N/A"

        sire, sire_url, dam, dam_url = extract_sire_dam_from_pedigree(soup)
        dog_data["Sire"], dog_data["Dam"] = sire, dam

        for field in FIELDS:
            try:
                val = driver.find_element(By.XPATH, f"//td[normalize-space(text())='{field}:']/following-sibling::td").text.strip()
            except NoSuchElementException:
                val = "N/A"
            dog_data[field] = val

        existing_df, index_map = append_or_update(existing_df, index_map, dog_data)

        def make_placeholder(url, name):
            if url == "N/A":
                return None
            return {
                "U-ID": re.search(r"ID=(\d+)", url).group(1),
                "URL": url,
                "Name": name,
                "Date of Birth": "N/A",
                "Date of Death": "N/A",
                "Sire": "N/A",
                "Dam": "N/A",
                **{f: "N/A" for f in FIELDS}
            }

        for entry, label in [(make_placeholder(sire_url, sire), "sire"), (make_placeholder(dam_url, dam), "dam")]:
            if entry and entry["U-ID"] not in index_map:
                existing_df, index_map = append_or_update(existing_df, index_map, entry)
                print(f"✅ Placeholder created for {label}: {entry['Name']}")
            elif entry:
                existing_df, index_map = append_or_update(existing_df, index_map, entry)
                print(f"🔄 Updated existing {label}: {entry['Name']}")

            save_progress()
            print("💾 Data saved successfully!")

            dog_data["URL"] = clean_url(dog_data["URL"])


            # 🧠 Add scraped dog to live memory (for control panel)
            LIVE_DOGS.insert(0, dog_data)
            if len(LIVE_DOGS) > 20:
                LIVE_DOGS.pop() 


        return {"status": "ok", "name": dog_data.get("Name", "Unknown"), "U-ID": dog_data.get("U-ID", "N/A")}
    except Exception as e:
        print(f"⚠️ Error scraping {dog_url}: {e}")
        traceback.print_exc()
        return {"status": "error", "message": str(e)}
    finally:
        try:
            driver.quit()
        except:
            pass

def rescrape_range(start_row, end_row):
    global existing_df, index_map
    print(f"\n🔁 Starting rescrape from row {start_row} to {end_row}...")

    driver = _create_driver()
    wait = WebDriverWait(driver, WAIT_TIME)

    subset = existing_df.iloc[start_row:end_row].copy()
    print(f"🧾 Found {len(subset)} dogs to rescrape.")

    processed = 0
    errors = 0

    for idx, row in subset.iterrows():
        url = str(row.get("URL", "N/A")).strip()
        uid = str(row.get("U-ID", "N/A")).strip()
        if url == "N/A" or not url.startswith("http"):
            print(f"⚠️ Skipping row {idx} (invalid URL).")
            continue

        print(f"\n➡️ Rescraping [{idx}] {row.get('Name', 'Unknown')} — {url}")

        try:
            driver.get(url)
            wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
            html = driver.page_source
            soup = BeautifulSoup(html, "html.parser")

            try:
                raw_name = driver.find_element(By.XPATH, "//strong/font[@size='4']").text.strip()
            except Exception:
                raw_name = "Unknown"

            dog_data = {"URL": clean_url(url), "U-ID": uid}


            if "(" in raw_name and ")" in raw_name:
                name_part, date_part = raw_name.split("(", 1)
                dog_data["Name"] = name_part.strip()
                date_part = date_part.strip(")")
                if "-" in date_part:
                    dob, dod = map(str.strip, date_part.split("-", 1))
                    dog_data["Date of Birth"], dog_data["Date of Death"] = dob, dod
                else:
                    dog_data["Date of Birth"], dog_data["Date of Death"] = date_part, "N/A"
            else:
                dog_data["Name"], dog_data["Date of Birth"], dog_data["Date of Death"] = raw_name, "N/A", "N/A"

            sire, sire_url, dam, dam_url = extract_sire_dam_from_pedigree(soup)
            dog_data["Sire"], dog_data["Dam"] = sire, dam

            for field in FIELDS:
                try:
                    val = driver.find_element(By.XPATH, f"//td[normalize-space(text())='{field}:']/following-sibling::td").text.strip()
                except NoSuchElementException:
                    val = "N/A"
                dog_data[field] = val

            existing_df, index_map = append_or_update(existing_df, index_map, dog_data)

            def make_placeholder(url, name):
                if url == "N/A":
                    return None
                return {
                    "U-ID": re.search(r"ID=(\d+)", url).group(1),
                    "URL": url,
                    "Name": name,
                    "Date of Birth": "N/A",
                    "Date of Death": "N/A",
                    "Sire": "N/A",
                    "Dam": "N/A",
                    **{f: "N/A" for f in FIELDS}
                }

            for entry, label in [(make_placeholder(sire_url, sire), "sire"), (make_placeholder(dam_url, dam), "dam")]:
                if entry and entry["U-ID"] not in index_map:
                    existing_df, index_map = append_or_update(existing_df, index_map, entry)
                    print(f"✅ Placeholder created for {label}: {entry['Name']}")
                elif entry:
                    existing_df, index_map = append_or_update(existing_df, index_map, entry)
                    print(f"🔄 Updated existing {label}: {entry['Name']}")

            existing_df = existing_df.drop_duplicates(subset=["U-ID"], keep="last").reset_index(drop=True)
            save_progress()
            processed += 1
            print("💾 Progress saved!\n")

            dog_data["URL"] = clean_url(dog_data["URL"])


             # 🧠 Add scraped dog to live memory (for control panel)
            LIVE_DOGS.insert(0, dog_data)
            if len(LIVE_DOGS) > 20:
                LIVE_DOGS.pop()

            delay = random.randint(10, 14)
            print(f"⏳ Waiting {delay}s before next dog...")
            time.sleep(delay)

        except Exception as e:
            errors += 1
            print(f"⚠️ Error rescraping {url}: {e}")
            traceback.print_exc()
            continue

    try:
        driver.quit()
    except:
        pass

    print("\n✅ RESCRAPE COMPLETE.")
    return {"status": "ok", "processed": processed, "errors": errors}

def manual_scrape_entry(breed_choice, SEARCH_TERM, resume_from):
    """
    The manual new scraping flow (from your script's 'else' branch).
    """
    global existing_df, index_map

    if not SEARCH_TERM:
        print("❌ Dog name cannot be empty!")
        return {"status": "error", "message": "Dog name cannot be empty!"}

    print(f"🐶 Manual scrape -> breed: {breed_choice}, name: {SEARCH_TERM}, resume_from: {resume_from}")

    options = uc.ChromeOptions()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = uc.Chrome(options=options, use_subprocess=True)
    wait = WebDriverWait(driver, WAIT_TIME)

    print("🚀 Launching Chrome...")
    driver.get(BASE_URL)

    try:
        wait.until(EC.presence_of_element_located((By.ID, "breed")))
        breed_select = Select(driver.find_element(By.ID, "breed"))
        breed_select.select_by_value("1" if (str(breed_choice) != "2" and str(breed_choice).lower().startswith("g")) else "2")

        search_box = wait.until(EC.presence_of_element_located((By.NAME, "name")))
        search_box.clear()
        search_box.send_keys(SEARCH_TERM)
        driver.find_element(By.XPATH, "//input[@type='submit' and @value='Search']").click()
        time.sleep(2)
    except Exception as e:
        print("❌ Search failed:", e)
        try:
            driver.quit()
        except:
            pass
        return {"status": "error", "message": f"Search failed: {e}"}

    dog_links, dog_names = [], []
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//p/a[contains(@href, 'pedigree.asp?ID=')]")))
        anchors = driver.find_elements(By.XPATH, "//p/a[contains(@href, 'pedigree.asp?ID=')]")
        dog_links = [a.get_attribute("href") for a in anchors if a.get_attribute("href")]
        dog_names = [a.text.strip() for a in anchors if a.text.strip()]
        print(f"✅ Found {len(dog_links)} profile links.")
    except Exception as e:
        print("⚠️ No results found:", e)

    start_index = 0
    resume_found = not bool(resume_from)

    if resume_from:
        resume_from_lower = resume_from.lower().strip()
        for i, (link, name) in enumerate(zip(dog_links, dog_names)):
            if (resume_from_lower in link.lower()) or (resume_from_lower in name.lower()):
                start_index = i
                resume_found = True
                print(f"⏩ Resuming scraping from dog #{i+1}: {name} — {link}")
                break
        if not resume_found:
            print(f"⚠️ Resume point '{resume_from}' not found — starting from beginning.")

    total = len(dog_links)
    processed = 0
    errors = 0

    for idx_link, link in enumerate(dog_links[start_index:], start=start_index + 1):
        try:
            driver.get(link)
            wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
            html = driver.page_source
            soup = BeautifulSoup(html, "html.parser")

            raw_name = "Unknown"
            try:
                raw_name = driver.find_element(By.XPATH, "//strong/font[@size='4']").text.strip()
            except Exception:
                pass

            dog_name_clean = raw_name.split("(")[0].strip() if "(" in raw_name else raw_name.strip()

            print(f"\n➡️ ({idx_link}/{len(dog_links)}) {dog_name_clean}")

            sire, sire_url, dam, dam_url = extract_sire_dam_from_pedigree(soup)

            dog_data = {"URL": clean_url(link), "U-ID": re.search(r"ID=(\d+)", link).group(1) if "ID=" in link else "N/A"}

            if "(" in raw_name and ")" in raw_name:
                name_part, date_part = raw_name.split("(", 1)
                dog_data["Name"] = name_part.strip()
                date_part = date_part.strip(")")
                if "-" in date_part:
                    dob, dod = map(str.strip, date_part.split("-", 1))
                    dog_data["Date of Birth"], dog_data["Date of Death"] = dob, dod
                else:
                    dog_data["Date of Birth"], dog_data["Date of Death"] = date_part, "N/A"
            else:
                dog_data["Name"], dog_data["Date of Birth"], dog_data["Date of Death"] = raw_name, "N/A", "N/A"

            dog_data["Sire"], dog_data["Dam"] = sire, dam

            for field in FIELDS:
                try:
                    val = driver.find_element(By.XPATH, f"//td[normalize-space(text())='{field}:']/following-sibling::td").text.strip()
                except NoSuchElementException:
                    val = "N/A"
                dog_data[field] = val

            existing_df, index_map = append_or_update(existing_df, index_map, dog_data)

            def make_placeholder(url, name):
                if url == "N/A": return None
                return {
                    "U-ID": re.search(r"ID=(\d+)", url).group(1),
                    "URL": url,
                    "Name": name,
                    "Date of Birth": "N/A",
                    "Date of Death": "N/A",
                    "Sire": "N/A",
                    "Dam": "N/A",
                    **{f: "N/A" for f in FIELDS}
                }

            for entry, label in [(make_placeholder(sire_url, sire), "sire"), (make_placeholder(dam_url, dam), "dam")]:
                if entry and entry["U-ID"] not in index_map:
                    existing_df, index_map = append_or_update(existing_df, index_map, entry)
                    print(f"✅ Placeholder created for {label}: {entry['Name']}")
                elif entry:
                    existing_df, index_map = append_or_update(existing_df, index_map, entry)
                    print(f"🔄 Updated existing {label}: {entry['Name']}")

            existing_df = existing_df.drop_duplicates(subset=["U-ID"], keep="last").reset_index(drop=True)

            save_progress()
            processed += 1
            print("💾 Progress saved!\n")

            dog_data["URL"] = clean_url(dog_data["URL"])



            # 🧠 Add scraped dog to live memory (for control panel)
            LIVE_DOGS.insert(0, dog_data)
            if len(LIVE_DOGS) > 20:
               LIVE_DOGS.pop()

            delay = random.randint(10, 14)
            print(f"⏳ Waiting {delay}s before next dog...")
            time.sleep(delay)

        except Exception as e:
            errors += 1
            print(f"⚠️ Error scraping {link}: {e}")
            traceback.print_exc()
            continue

    try:
        driver.quit()
    except:
        pass

    print("\n✅ SCRAPING COMPLETE — Data saved in:", SAVE_DIR)
    return {"status": "ok", "processed": processed, "errors": errors, "total": total}

# -------------------------
# Flask app + endpoints
# -------------------------
app = Flask(__name__, template_folder="templates")
CORS(app)

# ensure only one scrape runs at a time
scrape_lock = threading.Lock()

def run_and_capture(func, *args, **kwargs):
    """
    Runs func(*args, **kwargs) while capturing stdout prints.
    Returns (result_obj, stdout_text, exception_string_or_None)
    """
    buf = io.StringIO()
    try:
        with redirect_stdout(buf):
            result = func(*args, **kwargs)
    except Exception as e:
        tb = traceback.format_exc()
        return None, buf.getvalue(), tb
    return result, buf.getvalue(), None

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/scrape_single", methods=["POST"])
def api_scrape_single():
    payload = request.get_json(force=True)
    url = payload.get("url", "").strip()
    if not url:
        return jsonify({"status": "error", "message": "Missing url"}), 400

    if not scrape_lock.acquire(blocking=False):
        return jsonify({"status": "busy", "message": "Another scrape is running"}), 409

    try:
        res, logtxt, err = run_and_capture(scrape_single_dog, url)
        if err:
            return jsonify({"status": "error", "log": logtxt, "error": err}), 500
        return jsonify({"status": "ok", "log": logtxt, "result": res})
    finally:
        scrape_lock.release()

@app.route("/rescrape_range", methods=["POST"])
def api_rescrape_range():
    payload = request.get_json(force=True)
    start = payload.get("start_row")
    end = payload.get("end_row")
    if start is None or end is None:
        return jsonify({"status": "error", "message": "start_row and end_row required"}), 400

    try:
        start = int(start)
        end = int(end)
    except:
        return jsonify({"status": "error", "message": "start_row and end_row must be integers"}), 400

    if not scrape_lock.acquire(blocking=False):
        return jsonify({"status": "busy", "message": "Another scrape is running"}), 409

    try:
        res, logtxt, err = run_and_capture(rescrape_range, start, end)
        if err:
            return jsonify({"status": "error", "log": logtxt, "error": err}), 500
        return jsonify({"status": "ok", "log": logtxt, "result": res})
    finally:
        scrape_lock.release()

@app.route("/manual_scrape", methods=["POST"])
def api_manual_scrape():
    payload = request.get_json(force=True)
    breed = payload.get("breed", "Golden")
    name = payload.get("name", "").strip()
    resume = payload.get("resume_from", "").strip() or ""

    if not name:
        return jsonify({"status": "error", "message": "name is required"}), 400

    if not scrape_lock.acquire(blocking=False):
        return jsonify({"status": "busy", "message": "Another scrape is running"}), 409

    try:
        res, logtxt, err = run_and_capture(manual_scrape_entry, breed, name, resume)
        if err:
            return jsonify({"status": "error", "log": logtxt, "error": err}), 500
        return jsonify({"status": "ok", "log": logtxt, "result": res})
    finally:
        scrape_lock.release()

        # -------------------------
# NEW: Open CSV/XLSX/JSON in system default software
# -------------------------
from flask import send_file

@app.route("/open_file/<file_type>", methods=["GET"])
def open_file(file_type):
    SAVE_DIR = "k9_dog_data"
    mapping = {
        "csv": os.path.join(SAVE_DIR, "k9_data.csv"),
        "xlsx": os.path.join(SAVE_DIR, "k9_data.xlsx"),
        "json": os.path.join(SAVE_DIR, "k9_data.json")
    }

    file_path = mapping.get(file_type)
    if not file_path or not os.path.exists(file_path):
        return jsonify({"message": f"{os.path.basename(file_path)} file not found", "status": "error"})

    os.startfile(file_path)
    return jsonify({"message": f"Opened {file_type.upper()} file successfully", "status": "ok"})


# -------------------------
# NEW: Return full dog data as JSON
# -------------------------
@app.route("/dog_data_json", methods=["GET"])
def dog_data_json():
    try:
        return jsonify(json.loads(existing_df.to_json(orient="records")))
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/last_json", methods=["GET"])
def get_last_json():
    SAVE_DIR = "k9_dog_data"
    json_path = os.path.join(SAVE_DIR, "k9_data.json")

    if not os.path.exists(json_path):
        return jsonify({"error": "k9_data.json file not found"}), 404

    try:
        with open(json_path, "r", encoding="utf-8-sig") as f:
            data = json.load(f)

        # Ensure it's a list (some JSON files may contain dicts)
        if isinstance(data, dict):
            data = list(data.values())
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
LIVE_DOGS = []

@app.route("/api/live_dogs")
def api_live_dogs():
    return jsonify({
        "status": "ok",
        "count": len(LIVE_DOGS),
        "dogs": LIVE_DOGS[:20]
    })


from datetime import datetime
import os

@app.route("/data_dashboard")
def data_dashboard():
    SAVE_DIR = "k9_dog_data"  # ✅ Define inside the function

    files = []
    for fname in ["k9_data.csv", "k9_data.xlsx", "k9_data.json"]:
        path = os.path.join(SAVE_DIR, fname)
        if os.path.exists(path):
            stat = os.stat(path)
            ext = fname.rsplit(".", 1)[-1] if "." in fname else ""
            files.append({
                "name": fname,
                "ext": ext,
                "size": f"{os.path.getsize(path)/1024:.1f} KB",
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            })

    # ✅ This part must be OUTSIDE the for-loop
    # Get rows count from query parameter
    rows_to_show = int(request.args.get("rows", 10))

    # ✅ Safely handle dataframe preview
    try:
        preview_df = existing_df.head(rows_to_show).fillna("N/A")
        cols = list(preview_df.columns)
        preview = preview_df.to_dict(orient="records")
    except Exception as e:
        cols, preview = [], []

    return render_template("data_dashboard.html", files=files, columns=cols, preview=preview)


import pandas as pd
import json
import os
from flask import render_template, jsonify, request

@app.route("/view_file/<file_type>")
def view_file(file_type):
    SAVE_DIR = "k9_dog_data"
    mapping = {
        "csv": os.path.join(SAVE_DIR, "k9_data.csv"),
        "xlsx": os.path.join(SAVE_DIR, "k9_data.xlsx"),
        "json": os.path.join(SAVE_DIR, "k9_data.json")
    }

    file_path = mapping.get(file_type)
    if not file_path or not os.path.exists(file_path):
        return f"<h3>❌ File not found: {file_type.upper()}</h3>"

    try:
        # ✅ Handle all file types correctly
        if file_type == "csv":
            df = pd.read_csv(file_path, encoding="utf-8-sig")
        elif file_type == "xlsx":
            df = pd.read_excel(file_path)
        elif file_type == "json":
            with open(file_path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            # In case JSON is a dict instead of list
            if isinstance(data, dict):
                data = list(data.values())
            df = pd.DataFrame(data)
        else:
            return f"<h3>⚠️ Unsupported file type: {file_type}</h3>"

        df = df.fillna("N/A")
        cols = list(df.columns)
        data = df.to_dict(orient="records")

        # ✅ Render Admin File Viewer with data injected
        return render_template(
            "admin_view.html",
            file_type=file_type,
            columns=cols,
            data=data
        )

    except Exception as e:
        return f"<h3>❌ Error loading {file_type.upper()} file: {e}</h3>"



if __name__ == "__main__":
    print("Starting k9_api_scraper Flask server on http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=True)

