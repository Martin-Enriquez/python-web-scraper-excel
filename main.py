import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

def scrape_books():
    url = "http://books.toscrape.com/"
    response = requests.get(url)
    
    if response.status_code != 200:
        print(f"Error al acceder a la página: {response.status_code}")
        return []

    soup = BeautifulSoup(response.text, 'html.parser')
    books = soup.find_all('article', class_='product_pod')
    
    data = []
    for book in books:
        title = book.h3.a['title']
        # Limpieza de precio: de "£51.77" a flotante 51.77
        price_text = book.find('p', class_='price_color').text
        price = float(price_text.replace('£', '').replace('Â', ''))
        
        # Extracción de disponibilidad
        availability = book.find('p', class_='instock availability').text.strip()
        
        data.append({
            'Título': title,
            'Precio (£)': price,
            'Estado': availability
        })
    
    return data

def export_to_styled_excel(data, output_filename="reporte_libros.xlsx"):
    df = pd.DataFrame(data)
    
    # Crear writer de Pandas con motor OpenPyXL
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Libros Extraídos')
        workbook = writer.book
        worksheet = writer.sheets['Libros Extraídos']

        # Estilos visuales para el Excel
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        align_center = Alignment(horizontal="center", vertical="center")

        # Formatear Encabezados
        for col_num in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center

        # Ajustar ancho de columnas automáticamente
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    print(f"✅ Datos exportados exitosamente a {output_filename}")

if __name__ == "__main__":
    print("Iniciando Web Scraping...")
    results = scrape_books()
    if results:
        export_to_styled_excel(results)