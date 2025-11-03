# 🐾 K9 Scraper – Flask + Selenium Control Panel

A modern, browser-based **Dog Profile Scraper** built with **Flask**, **Selenium**, and **BeautifulSoup**.  
This tool lets you scrape, re-scrape, and visualize live dog profile data in real-time through an elegant web control panel.  

---

## 🚀 Key Features

- 🧠 **Flask Web Interface** – intuitive dashboard for full scraper control  
- 🐶 **Manual Scraping & Re-scraping** – scrape specific breeds or data ranges  
- 📊 **Real-time JSON Viewer** – see live scraped data updates instantly  
- 💾 **Multi-format Export** – automatically saves data as CSV, Excel, and JSON  
- 🕹️ **Simple Control Buttons** – start, re-scrape, or scrape single dogs easily  
- ⚙️ **Powerful Engine** – built using `undetected-chromedriver` to bypass site detection  
- 📋 **Integrated Logs** – see progress and debug messages right in the panel  

---

## 🧱 Project Structure

```
k9_scraper/
│
├── k9_api_server.py        # Flask backend server (dashboard)
├── k9_scraper.py           # Core scraping logic using Selenium
│
├── templates/              # HTML templates for dashboard UI
│   └── dashboard.html
│
├── k9_dog_data/            # Folder where CSV, JSON & Excel files are saved
├── chromedriver-win64/     # Chrome WebDriver (for Selenium)
├── venv/                   # Python virtual environment (ignored in Git)
│
├── .gitignore              # Files/folders ignored by GitHub
├── requirements.txt        # Required dependencies
└── README.md               # Project documentation
```

---

## ⚙️ Setup Instructions

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/YOUR_USERNAME/k9_scraper.git
cd k9_scraper
```

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

### 3️⃣ Install Dependencies
```bash
pip install -r requirements.txt
```

---

## 🖥️ Running the Web Dashboard

Once setup is complete, start the Flask app:

```bash
python k9_api_server.py
```

Then open your browser and go to:

👉 **http://127.0.0.1:5000**

You’ll see the **K9 Scraper Control Panel**, where you can:

- 🐕 Start or resume scraping  
- 🔁 Re-scrape a data range  
- 🔍 Scrape a single dog profile  
- 📂 Open your data files (CSV, Excel, JSON)  
- 🧠 View live JSON data from the scraper  

---

## 📁 Output Files

All scraped data is automatically saved inside the `k9_dog_data/` folder as:

- `dogs_data.csv`  
- `dogs_data.xlsx`  
- `dogs_data.json`

> These files are **ignored by GitHub** (via `.gitignore`) to keep your repo lightweight.

---

## 🧠 Tech Stack

- **Python 3.10+**  
- **Flask** – backend framework  
- **Selenium** – web automation  
- **undetected-chromedriver** – bypasses bot detection  
- **BeautifulSoup4** – HTML parsing  
- **pandas** – data management  
- **openpyxl** – Excel file handling  
- **HTML / CSS / JS** – frontend design  

---

## 📦 Requirements

List of dependencies inside `requirements.txt`:

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

To install them:
```bash
pip install -r requirements.txt
```

---

## 💡 Useful Tips

- Make sure you have **Google Chrome** installed.  
- Your **chromedriver** version should match your Chrome version.  
- To stop the scraper safely, press `CTRL + C` in the terminal.  
- Avoid running multiple scraping sessions at once.  
- Keep your data folder clean for better organization.  

---

## 📄 License

This project is released under the **MIT License** — free to use, modify, and share with credit.  

---

## 👨‍💻 Author

**Developed by:** *Wajid Khanzada*  
**GitHub:** [https://github.com/YOUR_USERNAME](https://github.com/YOUR_USERNAME)

---

### ⭐ If you found this project helpful, please give it a star on GitHub!
