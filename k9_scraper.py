# =========================
# k9_scraper
# =========================
import os
import re
import json
import time
import random
import pandas as pd
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import undetected_chromedriver as uc

# =========================
# CONFIGURATION
# =========================
BASE_URL = "http://www.k9data.com/"
SAVE_DIR = "k9_dog_data"
WAIT_TIME = 12

FIELDS = [
    "Call name", "Gender", "Honorifics", "Country of origin",
    "Country of residence", "Registration", "Breeder", "Owner",
    "Web site", "Hip clearance", "Eye clearance", "Heart clearance",
    "Elbow clearance", "Cause of death"
]

# =========================
# FILE SETUP
# =========================
os.makedirs(SAVE_DIR, exist_ok=True)
csv_path = os.path.join(SAVE_DIR, "k9_data.csv")
xlsx_path = os.path.join(SAVE_DIR, "k9_data.xlsx")
json_path = os.path.join(SAVE_DIR, "k9_data.json")

if os.path.exists(csv_path):
    existing_df = pd.read_csv(csv_path, dtype=str).fillna("N/A")
else:
    existing_df = pd.DataFrame(columns=["URL", "U-ID", "Name", "Date of Birth", "Date of Death", "Sire", "Dam"] + FIELDS)

for col in ["URL", "U-ID", "Name", "Date of Birth", "Date of Death", "Sire", "Dam"] + FIELDS:
    if col not in existing_df.columns:
        existing_df[col] = "N/A"

def build_index_map(df):
    df["U-ID"] = df["U-ID"].fillna("N/A").astype(str).str.strip()
    return {uid: idx for idx, uid in df["U-ID"].items() if uid and uid != "N/A"}

index_map = build_index_map(existing_df)

# =========================
# APPEND / UPDATE FUNCTION
# =========================
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


# =========================
# ✅ FIXED SIRE/DAM EXTRACTION
# =========================
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


# =========================
# 🐕 SCRAPE SINGLE DOG FUNCTION (Option 3)
# =========================
def scrape_single_dog(dog_url):
    global existing_df, index_map

    if not dog_url.startswith("http"):
        print("❌ Invalid URL!")
        return

    print(f"\n➡️ Scraping single dog: {dog_url}")

    options = uc.ChromeOptions()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = uc.Chrome(options=options, use_subprocess=True)
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

        dog_data = {"URL": dog_url, "U-ID": re.search(r"ID=(\d+)", dog_url).group(1) if "ID=" in dog_url else "N/A"}

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
        existing_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        existing_df.to_excel(xlsx_path, index=False)
        existing_df.to_json(json_path, orient="records", indent=2)
        print("💾 Data saved successfully!")

    except Exception as e:
        print(f"⚠️ Error scraping {dog_url}: {e}")
    finally:
        driver.quit()


# =========================
# 🔁 RESCRAPE RANGE FUNCTION
# =========================
def rescrape_range(start_row, end_row):
    print(f"\n🔁 Starting rescrape from row {start_row} to {end_row}...")

    options = uc.ChromeOptions()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")

    driver = uc.Chrome(options=options, use_subprocess=True)
    wait = WebDriverWait(driver, WAIT_TIME)

    global existing_df, index_map

    subset = existing_df.iloc[start_row:end_row].copy()
    print(f"🧾 Found {len(subset)} dogs to rescrape.")

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

            dog_data = {"URL": url, "U-ID": uid}

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
            existing_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            existing_df.to_excel(xlsx_path, index=False)
            existing_df.to_json(json_path, orient="records", indent=2)
            print("💾 Progress saved!\n")

            delay = random.randint(10, 14)
            print(f"⏳ Waiting {delay}s before next dog...")
            time.sleep(delay)

        except Exception as e:
            print(f"⚠️ Error rescraping {url}: {e}")
            continue

    driver.quit()
    print("\n✅ RESCRAPE COMPLETE.")


# =========================
# MAIN MENU
# =========================
print("🐾 Choose mode:")
print("  1 = 🆕 Manual new scraping")
print("  2 = 🔁 Re-scrape existing data")
print("  3 = 🐕 Scrape single dog by URL")
mode = input("👉 Enter choice (1, 2, or 3): ").strip()

if mode == "2":
    print(f"\n📘 Existing rows: {len(existing_df)}")
    start = int(input("➡️ Enter start row: "))
    end = int(input("➡️ Enter end row: "))
    rescrape_range(start, end)

elif mode == "3":
    dog_url = input("🐕 Enter full dog profile URL: ").strip()
    scrape_single_dog(dog_url)

# ✅ FIXED RESUME FEATURE BELOW
else:
    print("🐶 Choose a breed:")
    print("  1 = Golden Retriever")
    print("  2 = Labrador Retriever")
    breed_choice = input("👉 Enter breed number (1 or 2): ").strip()

    SEARCH_TERM = input("🔍 Enter dog name to search: ").strip()
    if not SEARCH_TERM:
        print("❌ Dog name cannot be empty!")
        raise SystemExit

    resume_from = input("➡️ Enter dog name or ID to resume from (press Enter to start from beginning): ").strip()

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
        breed_select.select_by_value("1" if breed_choice != "2" else "2")

        search_box = wait.until(EC.presence_of_element_located((By.NAME, "name")))
        search_box.clear()
        search_box.send_keys(SEARCH_TERM)
        driver.find_element(By.XPATH, "//input[@type='submit' and @value='Search']").click()
        time.sleep(2)
    except Exception as e:
        print("❌ Search failed:", e)
        driver.quit()
        raise SystemExit

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

            dog_data = {"URL": link, "U-ID": re.search(r"ID=(\d+)", link).group(1) if "ID=" in link else "N/A"}
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
            existing_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            existing_df.to_excel(xlsx_path, index=False)
            existing_df.to_json(json_path, orient="records", indent=2)
            print("💾 Progress saved!\n")

            delay = random.randint(10, 14)
            print(f"⏳ Waiting {delay}s before next dog...")
            time.sleep(delay)

        except Exception as e:
            print(f"⚠️ Error scraping {link}: {e}")
            continue

    driver.quit()
    print("\n✅ SCRAPING COMPLETE — Data saved in:", SAVE_DIR)
