"""Web scraper de books.toscrape.com con exportación a Excel con formato.

Uso básico:
    python main.py
    python main.py --max-pages 5 --output libros.xlsx
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

START_URL = "http://books.toscrape.com/catalogue/page-1.html"
TIMEOUT = 10  # segundos
SHEET_NAME = "Libros Extraídos"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("scraper")


# --------------------------------------------------------------------------
# Extracción
# --------------------------------------------------------------------------
def build_session() -> requests.Session:
    """Sesión HTTP con User-Agent propio y reintentos automáticos."""
    session = requests.Session()
    session.headers.update(
        {"User-Agent": "Mozilla/5.0 (compatible; python-web-scraper-excel/1.0)"}
    )
    retry = Retry(total=3, backoff_factor=1, status_forcelist=(429, 500, 502, 503, 504))
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def parse_price(text: str) -> float:
    """Convierte '£51.77' en 51.77."""
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        raise ValueError(f"No se pudo interpretar el precio: {text!r}")
    return float(match.group())


def parse_books(soup: BeautifulSoup, page_url: str) -> list[dict]:
    """Extrae los libros de una página del catálogo."""
    books = []
    for article in soup.select("article.product_pod"):
        rating_tag = article.select_one("p.star-rating")
        rating_word = next(
            (c for c in (rating_tag.get("class", []) if rating_tag else []) if c in RATINGS),
            None,
        )
        books.append(
            {
                "Título": article.h3.a["title"],
                "Precio (£)": parse_price(article.select_one("p.price_color").text),
                "Valoración (1-5)": RATINGS.get(rating_word),
                "Disponibilidad": article.select_one("p.instock.availability").text.strip(),
                "URL": urljoin(page_url, article.h3.a["href"]),
            }
        )
    return books


def scrape_books(
    start_url: str = START_URL, max_pages: int | None = None, delay: float = 0.5
) -> list[dict]:
    """Recorre el catálogo siguiendo el enlace 'next' hasta terminar."""
    session = build_session()
    data: list[dict] = []
    url: str | None = start_url
    pages = 0

    while url and (max_pages is None or pages < max_pages):
        log.info("Leyendo %s", url)
        try:
            response = session.get(url, timeout=TIMEOUT)
            response.raise_for_status()
        except requests.RequestException as exc:
            log.error("Error al acceder a %s: %s", url, exc)
            break

        # El sitio no declara la codificación; sin esto '£' llega como 'Â£'.
        response.encoding = "utf-8"
        soup = BeautifulSoup(response.text, "html.parser")

        data.extend(parse_books(soup, response.url))
        pages += 1

        next_link = soup.select_one("li.next > a")
        url = urljoin(response.url, next_link["href"]) if next_link else None
        if url:
            time.sleep(delay)  # cortesía con el servidor

    log.info("Páginas leídas: %d | Libros extraídos: %d", pages, len(data))
    return data


# --------------------------------------------------------------------------
# Exportación
# --------------------------------------------------------------------------
def export_to_styled_excel(data: list[dict], output_filename: str) -> None:
    """Exporta los datos a un .xlsx con encabezados estilizados."""
    df = pd.DataFrame(data)

    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=SHEET_NAME)
        ws = writer.sheets[SHEET_NAME]

        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        center = Alignment(horizontal="center", vertical="center")

        # Encabezados
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center

        columns = list(df.columns)
        price_col = columns.index("Precio (£)") + 1
        rating_col = columns.index("Valoración (1-5)") + 1
        url_col = columns.index("URL") + 1
        link_font = Font(color="0563C1", underline="single")

        # Formato de datos
        for row in ws.iter_rows(min_row=2):
            row[price_col - 1].number_format = '"£"#,##0.00'
            row[rating_col - 1].alignment = center
            url_cell = row[url_col - 1]
            url_cell.hyperlink = url_cell.value
            url_cell.font = link_font

        # Comodidad de uso
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Ancho de columnas automático (con tope)
        for col in ws.columns:
            longest = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(
                max(longest + 3, 12), 60
            )

    log.info("Datos exportados a %s", output_filename)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scraper de libros con exportación a Excel.")
    parser.add_argument("--url", default=START_URL, help="URL de la primera página del catálogo.")
    parser.add_argument("--max-pages", type=int, default=None, help="Límite de páginas (por defecto, todas).")
    parser.add_argument("--delay", type=float, default=0.5, help="Segundos de espera entre páginas.")
    parser.add_argument("--output", default="reporte_libros.xlsx", help="Nombre del archivo Excel.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    results = scrape_books(args.url, args.max_pages, args.delay)
    if not results:
        log.error("No se extrajeron datos; no se generó el Excel.")
        return 1

    export_to_styled_excel(results, args.output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
