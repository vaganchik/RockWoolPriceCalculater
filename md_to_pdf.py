import markdown
import pdfkit
import os

# ПУТЬ К ИСПОЛНЯЕМОМУ ФАЙЛУ wkhtmltopdf (нужно установить отдельно)
PATH_TO_WKHTMLTOPDF = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

def convert_md_to_pdf(input_md_path, output_pdf_path):
    """Конвертирует Markdown в PDF с применением стилей оформления."""
    try:
        # 1. Чтение исходного файла
        with open(input_md_path, 'r', encoding='utf-8') as f:
            md_content = f.read()

        # 2. Преобразование Markdown в HTML (с поддержкой таблиц)
        html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])

        # 3. Оформление (CSS) для PDF
        css_style = """
        <style>
            body { font-family: 'Arial', sans-serif; line-height: 1.6; color: #333; margin: 40px; }
            h1 { color: #2c3e50; text-align: center; border-bottom: 2px solid #2c3e50; padding-bottom: 10px; }
            h2 { color: #2980b9; border-bottom: 1px solid #eee; padding-bottom: 5px; margin-top: 30px; }
            table { border-collapse: collapse; width: 100%; margin: 20px 0; }
            th, td { border: 1px solid #bdc3c7; padding: 10px; text-align: left; }
            th { background-color: #ecf0f1; font-weight: bold; }
            code { background-color: #f4f4f4; padding: 2px 4px; border-radius: 3px; font-family: 'Courier New', monospace; }
            pre { background-color: #f4f4f4; padding: 15px; border-radius: 5px; border: 1px solid #ddd; }
            blockquote { border-left: 5px solid #3498db; padding-left: 15px; color: #555; font-style: italic; }
        </style>
        """
        
        # Сборка полного HTML-документа
        full_html = f"<html><head><meta charset='utf-8'>{css_style}</head><body>{html_content}</body></html>"

        # 4. Настройка и генерация PDF
        config = pdfkit.configuration(wkhtmltopdf=PATH_TO_WKHTMLTOPDF)
        options = {
            'encoding': "UTF-8",
            'enable-local-file-access': None,
            'margin-top': '20mm',
            'margin-bottom': '20mm',
            'margin-left': '20mm',
            'margin-right': '20mm',
        }

        pdfkit.from_string(full_html, output_pdf_path, configuration=config, options=options)
        print(f"Успешно! PDF-файл создан: {output_pdf_path}")

    except Exception as e:
        print(f"Ошибка: {e}")

if __name__ == "__main__":
    # Конвертируем вашу документацию по расчетам
    base_path = os.path.dirname(os.path.abspath(__file__))
    convert_md_to_pdf(
        os.path.join(base_path, 'CALCULATION_LOGIC.md'),
        os.path.join(base_path, 'CALCULATION_LOGIC.pdf')
    )