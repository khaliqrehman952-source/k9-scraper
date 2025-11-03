

```markdown
# 🐾 K9 Scraper - Flask + Selenium Control Panel

A Python-based web scraper for K9 dog profiles with a modern Flask-powered dashboard.  
This tool allows you to scrape, re-scrape, and view live dog data directly from a browser-based control panel.

---

## 🚀 Features

✅ Flask web interface for easy control  
✅ Manual scraping and re-scraping options  
✅ Real-time JSON preview of scraped data  
✅ Data export to CSV, Excel (XLSX), and JSON  
✅ Built-in logs and status display  
✅ Uses **undetected-chromedriver** to bypass bot detection  
✅ Organized folder structure and modular codebase  

---

## 🧱 Project Structure

```

k9_scraper/
│
├── k9_api_server.py        # Flask backend server
├── k9_scraper.py           # Main scraping logic
├── templates/              # HTML templates for dashboard
│   └── dashboard.html
│
├── k9_dog_data/            # Saved CSV, JSON, and Excel data (ignored in Git)
├── chromedriver-win64/     # Chrome WebDriver folder
├── venv/                   # Virtual environment (ignored in Git)
│
├── .gitignore
├── requirements.txt
└── README.md

````

---

## ⚙️ Setup Instructions

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/YOUR_USERNAME/k9_scraper.git
cd k9_scraper
````

### 2️⃣ Create & Activate a Virtual Environment

**Windows (PowerShell):**

```bash
python -m venv venv
venv\Scripts\activate
```

**Mac/Linux (bash):**

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3️⃣ Install Required Packages

```bash
pip install -r requirements.txt
```

---

## 🖥️ Running the Project

To start the Flask control panel:

```bash
python k9_api_server.py
```

Then open your browser and visit:

```
http://127.0.0.1:5000
```

You’ll see the **K9 Scraper Control Panel**, where you can:

* Start manual scraping
* Re-scrape specific ranges
* Scrape a single dog profile
* View live JSON output
* Open your exported data files

---

## 📁 Output Files

All data files are saved inside the `k9_dog_data/` folder:

* `dogs_data.csv`
* `dogs_data.xlsx`
* `dogs_data.json`

> These files are automatically ignored in `.gitignore` to keep your GitHub repo clean.

---

## 🧠 Tech Stack

* **Python 3.10+**
* **Flask** (web framework)
* **Selenium** (web automation)
* **undetected-chromedriver**
* **BeautifulSoup4**
* **pandas**
* **openpyxl**
* **HTML/CSS/JS** (for frontend)

---

## 🧩 Requirements File

All dependencies are listed in `requirements.txt`:

```
Flask>=2.0
flask-cors
pandas
beautifulsoup4
selenium
undetected-chromedriver
openpyxl
xlrd
lxml
```

Install them anytime with:

```bash
pip install -r requirements.txt
```

---

## 🧰 Useful Tips

* Ensure **Google Chrome** is installed on your machine.
* Keep your **chromedriver** version matching your Chrome browser.
* To stop the scraper safely, press `CTRL + C` in the terminal.
* For best results, avoid running multiple scraping sessions simultaneously.

---

## 📄 License

This project is released under the **MIT License** — free to use, modify, and distribute with attribution.

---

## 💡 Author

**Developed by:** *Wajid Khanzada*
**GitHub:** [https://github.com/YOUR_USERNAME](https://github.com/YOUR_USERNAME)

---

### ⭐ If you like this project, consider giving it a star on GitHub!


"# k9-scraper" 
"# k9-scraper" 
