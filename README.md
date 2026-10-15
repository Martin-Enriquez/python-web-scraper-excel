# 📊 Web Scraper → Excel con formato

Script en Python que extrae el catálogo completo de [books.toscrape.com](http://books.toscrape.com/) (1000 libros, 50 páginas), limpia los datos y genera un archivo Excel listo para compartir.

> **Nota:** `books.toscrape.com` es un sitio de práctica creado para aprender scraping. El proyecto demuestra el flujo completo (extracción → limpieza → reporte) y la misma estructura se puede adaptar a otros sitios, respetando sus términos de uso.

*English summary: Python scraper that crawls the full books.toscrape.com catalogue (with pagination), cleans the data with pandas and exports a styled `.xlsx` report with openpyxl. Usage instructions below.*

## ✨ Qué hace

- **Recorre todas las páginas** del catálogo siguiendo el enlace "next".
- **Limpia los datos:** precios como número (`£51.77` → `51.77`), valoración en estrellas como entero (1–5).
- **Es tolerante a fallos:** `timeout`, reintentos automáticos y espera entre peticiones.
- **Genera un Excel profesional:** encabezados con color, formato de moneda, filtros, fila fija, anchos de columna automáticos y enlaces clicables a cada libro.

## 📋 Columnas del reporte

| Columna | Descripción |
|---|---|
| Título | Nombre completo del libro |
| Precio (£) | Precio como número, con formato de moneda en Excel |
| Valoración (1-5) | Estrellas del libro |
| Disponibilidad | Estado de stock |
| URL | Enlace a la ficha del libro |

## 🛠️ Tecnologías

Python 3.9+ · Requests · BeautifulSoup4 · pandas · openpyxl

## 🚀 Instalación y uso

```bash
# 1. Clonar el repositorio
git clone https://github.com/Martin-Enriquez/python-web-scraper-excel.git
cd python-web-scraper-excel

# 2. (Opcional) Crear un entorno virtual
python -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar
python main.py
```

Se generará `reporte_libros.xlsx` en la carpeta del proyecto.

### Opciones

```bash
python main.py --max-pages 5              # solo las primeras 5 páginas
python main.py --output mis_libros.xlsx   # nombre del archivo de salida
python main.py --delay 1                  # 1 segundo entre peticiones
python main.py --help                     # ver todas las opciones
```

## 🗂️ Estructura

```
main.py            # scraper + exportación a Excel
requirements.txt   # dependencias
LICENSE            # MIT
```

Funciones principales en `main.py`: `scrape_books()` (extracción con paginación), `parse_books()` (parseo de una página) y `export_to_styled_excel()` (generación del reporte).

## 🗺️ Ideas para seguir

- Añadir la categoría de cada libro (requiere visitar cada ficha).
- Exportar también a CSV o a una base de datos SQLite.
- Pruebas automáticas para `parse_price` y `parse_books`.

## 📄 Licencia

MIT — ver [LICENSE](LICENSE).
