# 📚 Web Scraper to Excel (USD)

Python command-line tool that scrapes a product catalog and exports a clean, formatted Excel report with prices converted to US dollars.

## Features

- **Full catalog crawl:** follows pagination automatically, with retries on server errors and a polite delay between requests.
- **Clean data:** title, price, rating (1-5), availability, and a clickable product link.
- **USD conversion:** prices are converted from GBP using a configurable exchange rate and formatted as `$`.
- **Styled Excel report:** colored headers, frozen header row, auto-filters, auto-fitted columns.
- **Summary sheet:** run date, source, exchange rate, number of books, and price/rating statistics.

## Quick Start

```bash
git clone https://github.com/<your-user>/<your-repo>.git
cd <your-repo>
pip install -r requirements.txt
python main.py
```

## Options

| Option | Default | Description |
|---|---|---|
| `--url` | books.toscrape.com, page 1 | First catalog page to scrape. |
| `--max-pages` | all pages | Limit the number of pages (e.g. `--max-pages 5`). |
| `--delay` | `0.5` | Seconds to wait between pages. |
| `--rate` | `1.34` (or env `GBP_USD_RATE`) | GBP → USD exchange rate. Update it before real use. |
| `--include-gbp` | off | Adds the original GBP price as an extra column. |
| `--output` | `books_report.xlsx` | Excel output file name. |

Example:

```bash
python main.py --max-pages 5 --rate 1.34 --include-gbp --output books.xlsx
```

## Output

**`books_report.xlsx`**

| Sheet | Contents |
|---|---|
| `Books` | Title, price ($), rating, availability, and product URL (optionally the source GBP price). |
| `Summary` | Run date, source, currencies, exchange rate, books extracted, average / lowest / highest price, and average rating. |

## Scope & Limitations

- The data source is [books.toscrape.com](http://books.toscrape.com), a sandbox site built for scraping practice.
- Prices are converted at a fixed rate that you provide; they are not live exchange rates.
- Prices exclude taxes and shipping.
- Built for sites with static HTML. Sites with login, heavy JavaScript, or anti-bot protection are not supported.
- Always check a site's terms of use and `robots.txt` before scraping it.

## Tech Stack

Python 3.9+ · Requests · BeautifulSoup4 · Pandas · OpenPyXL

## License

See the [LICENSE](LICENSE) file for details.
