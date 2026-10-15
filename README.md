# 📊 Web Scraper → Formatted Excel Exporter

A robust Python script that crawls the full [books.toscrape.com](http://books.toscrape.com/) catalogue (1,000 books across 50 pages), cleans the extracted data, and generates a professionally formatted Excel report.

> **Note:** `books.toscrape.com` is a practice website designed for web scraping. This project demonstrates an end-to-end data pipeline (extraction → cleaning → reporting) that can easily be adapted for other e-commerce platforms and business automation needs.

## ✨ Features

- **Full Catalogue Crawling:** Handles automated pagination by following the "next" page links.
- **Data Cleaning & Normalization:** Parses price strings to numeric float values (`£51.77` → `51.77`) and converts star ratings to integers (1–5).
- **Fault-Tolerant & Polite:** Implements HTTP request timeouts, automatic retry strategies with backoff, and custom request delays to respect target servers.
- **Professional Excel Output:** Custom header styling, currency formatting, auto-filters, frozen top rows, auto-adjusted column widths, and clickable links to product pages.

## 📋 Report Structure

| Column | Description |
|---|---|
| Title | Full product title |
| Price (£) | Formatted numeric price |
| Rating (1-5) | Star rating converted to integer |
| Availability | Stock availability status |
| URL | Direct link to the book product page |

## 🛠️ Tech Stack

Python 3.9+ · Requests · BeautifulSoup4 · Pandas · OpenPyXL

## 🚀 Installation & Usage

```bash
# 1. Clone the repository
git clone [https://github.com/Martin-Enriquez/python-web-scraper-excel.git](https://github.com/Martin-Enriquez/python-web-scraper-excel.git)
cd python-web-scraper-excel

# 2. (Optional) Create a virtual environment
python3 -m venv venv
source venv/bin/activate        # On Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the scraper
python3 main.py
```
<img width="1280" height="563" alt="Captura de pantalla 2026-10-14 a la(s) 9 10 39 p  m" src="https://github.com/user-attachments/assets/7489b6e8-22e7-43a3-a3f0-0016ca94ebf5" />
