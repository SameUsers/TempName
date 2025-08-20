import os
import re
import math
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from PIL import Image, ImageEnhance
import pytesseract
from docx import Document
from difflib import get_close_matches
from openpyxl.styles import numbers


class Pdf_Worker:
    """
    Класс для извлечения таблиц и данных из PDF-файлов и сохранения их в XLSX.
    Поддерживает:
    - Табличные PDF-файлы (Specification, PL)
    - PDF с линиями Invoice (Invoice_purchase)
    """

    def __init__(self):
        pass

    def pdf_to_xlsx(self, pdf_path: str, xlsx_path: str,
                    filter_spec: bool = False,
                    remove_edges: bool = False,
                    invoice_lines: bool = False) -> bool:
        """
        Конвертирует PDF в XLSX.
        
        :param pdf_path: Путь к PDF-файлу
        :param xlsx_path: Путь для сохранения XLSX
        :param filter_spec: Фильтровать строки для Specification
        :param remove_edges: Удалять первую и последнюю строку таблиц (PL)
        :param invoice_lines: Парсить как Invoice_purchase с регулярными выражениями
        :return: True, если удалось сохранить XLSX, иначе False
        """

        tables_standard = []
        tables_invoice = []

        with pdfplumber.open(pdf_path) as pdf:
            if invoice_lines:
                pattern = re.compile(
                    r'^(\d+)\s+(\d+)\s+(.*?)\s+(\d+)\s+'
                    r'(Stk\.?\s*/\s*\d+\s*\w*)\s+'
                    r'([\d,.]+(?:\s*/?\s*[\d,.]*)?)\s+'
                    r'([\d,.]+)$'
                )

                all_lines = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_lines.extend([ln.strip() for ln in text.split("\n") if ln.strip()])

                skip_patterns = [
                    r'^Rechnung', r'^Kundennummer', r'^Rechnungsdatum', r'^Lieferdatum',
                    r'^IBAN', r'^BIC', r'^UST-Id', r'^Zwischensumme', r'^Vortrag',
                    r'^Seite', r'^Amtsgericht', r'^HRB', r'^GF:', r'^Gläubiger-ID',
                    r'^www\.', r'^http', r'^Eissing GmbH',
                ]

                allowed_patterns = [
                    r'^EAN', r'^Art', r'^Hersteller'
                ]

                customs_pattern = re.compile(r'^Zolltarif-Nr\.?:\s*(\d+)')

                i = 0
                while i < len(all_lines):
                    line = all_lines[i]
                    match = pattern.match(line)
                    if match:
                        groups = list(match.groups())
                        description = groups[2]
                        customs_code = None
                        j = i + 1
                        while j < len(all_lines):
                            next_line = all_lines[j]
                            if re.match(r'^\d+\s+\d+\s+', next_line):
                                break
                            if any(re.match(pat, next_line) for pat in skip_patterns):
                                j += 1
                                continue
                            customs_match = customs_pattern.match(next_line)
                            if customs_match:
                                customs_code = customs_match.group(1)
                                j += 1
                                continue
                            if any(re.match(pat, next_line) for pat in allowed_patterns):
                                description += " " + next_line
                                j += 1
                                continue
                            break
                        groups[2] = description.strip()
                        df = pd.DataFrame([groups + [customs_code]], columns=[
                            "№", "Code", "Description", "Quantity",
                            "Unit/Volume", "PricePerUnit", "TotalPrice",
                            "CustomsCode"
                        ])
                        tables_invoice.append(df)
                        i = j
                    else:
                        i += 1

            else:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        df = pd.DataFrame(table)
                        if filter_spec:
                            df[0] = pd.to_numeric(df[0], errors='coerce')
                            max_val = df[0].max()
                            df = df[(df[0] >= 1) & (df[0] <= max_val)].reset_index(drop=True)
                        if remove_edges and len(df) > 2:
                            df = df.iloc[1:-1].reset_index(drop=True)
                        tables_standard.append(df)

        all_tables = tables_standard + tables_invoice
        if all_tables:
            result_df = pd.concat(all_tables, axis=0, ignore_index=True, sort=False)
            result_df.to_excel(xlsx_path, index=False)
            return True
        return False


class File_Generate:
    """
    Класс для генерации итогового XLSX-файла на основе шаблона.
    Поддерживает:
    - Заполнение названий товаров из справочника по CustomsCode
    - Заполнение данных из PL.xlsx
    - Заполнение данных из Specifications_sell.xlsx
    - Заполнение колонки N значениями из invoice (колонка H)
    - Авто-заполнение колонки E через OCR из изображений
    - Математическое округление для определённых колонок
    """

    def __init__(self):
        pass

    def fill_invoice(self,
                     template_filename: str,
                     invoice_filename: str,
                     ref_filename: str,
                     pl_filename: str,
                     spec_filename: str,
                     output_filename: str,
                     photos_folder: str = "PHOTOS") -> bool:
        import os, re, math
        from difflib import get_close_matches
        import pandas as pd
        from openpyxl import load_workbook
        from PIL import Image, ImageEnhance
        import pytesseract

        # Пути к файлам
        template_path = os.path.join("examples", template_filename)
        invoice_path = os.path.join("xlsx_files", invoice_filename)
        ref_path = os.path.join("examples", ref_filename)
        pl_path = os.path.join("xlsx_files", pl_filename)
        spec_path = os.path.join("xlsx_files", spec_filename)
        output_path = os.path.join("examples", output_filename)

        # --- Загружаем файлы ---
        invoice_df = pd.read_excel(invoice_path)
        ref_df = pd.read_excel(ref_path, header=None)
        pl_df = pd.read_excel(pl_path)
        spec_df = pd.read_excel(spec_path)

        ref_map = dict(zip(ref_df[3], ref_df[2]))

        def safe_column(df, col_index, numeric=False, replace_kan=False):
            values = df.iloc[:, col_index]
            if replace_kan:
                values = values.astype(str).str.replace("Kan", "").str.strip()
            if numeric:
                return [
                    float(str(x).replace(",", ".").strip()) if str(x).strip() not in ["", "nan", "None"] else 0
                    for x in values
                ]
            return values.fillna("").tolist()

        # --- Данные из PL.xlsx ---
        pl_col_B = safe_column(pl_df, 1, replace_kan=True)   # D
        pl_col_E = safe_column(pl_df, 4, numeric=True)       # G
        pl_col_D = safe_column(pl_df, 3, numeric=True)       # H
        pl_col_H = safe_column(pl_df, 7, numeric=True)       # J (округлять по твоему правилу или как было)
        pl_col_I = safe_column(pl_df, 8, numeric=True)       # K

        # --- Данные из Specifications_sell.xlsx ---
        import re
        spec_values = [
            float(re.sub(r'[^0-9,]', '', str(x)).replace(',', '.')) if re.search(r'\d', str(x)) else 0
            for x in spec_df.iloc[:, 4]
        ]

        # --- Данные из invoice, колонка H ---
        invoice_col_H = safe_column(invoice_df, 7)

        # --- Загружаем шаблон ---
        wb = load_workbook(template_path)
        ws = wb.active
        row_start = 18

        # --- CustomsCode -> название (в C) ---
        for idx, code in enumerate(invoice_df["CustomsCode"].fillna("").tolist()):
            ws.cell(row=row_start + idx, column=3, value=ref_map.get(code, ""))

        # --- PL.xlsx колонки ---
        for idx, val in enumerate(pl_col_B):
            ws.cell(row=row_start + idx, column=4, value=val)  # D

        for idx, val in enumerate(pl_col_E):
            ws.cell(row=row_start + idx, column=7, value=val)  # G

        for idx, val in enumerate(pl_col_D):
            ws.cell(row=row_start + idx, column=8, value=val)  # H

        for idx, val in enumerate(pl_col_H):
            # Округление как раньше (ceil для >=0 и floor для <0). Если нужно "0.5 вверх, 0.49 вниз" — замени тут.
            ws.cell(row=row_start + idx, column=10, value=math.ceil(val) if val >= 0 else math.floor(val))  # J

        for idx, val in enumerate(pl_col_I):
            ws.cell(row=row_start + idx, column=11, value=math.ceil(val) if val >= 0 else math.floor(val))  # K# --- Specifications_sell.xlsx в L ---
        for idx, val in enumerate(spec_values):
            ws.cell(row=row_start + idx, column=12, value=val)

        # --- Invoice, колонка H в N ---
        for idx, val in enumerate(invoice_col_H):
            ws.cell(row=row_start + idx, column=14, value=val)

        # --- Заполнение колонки E через OCR ---
        from difflib import get_close_matches
        for idx, product_name in enumerate(pl_col_B):
            try:
                candidate_folders = os.listdir(photos_folder)
            except FileNotFoundError:
                candidate_folders = []
            best_match = get_close_matches(product_name, candidate_folders, n=1, cutoff=0.5)
            if not best_match:
                continue
            folder_path = os.path.join(photos_folder, best_match[0])
            if not os.path.isdir(folder_path):
                continue
            jpg_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".jpg")][:2]
            made_in_value = "EU"  # дефолт
            for jpg in jpg_files:
                image_path = os.path.join(folder_path, jpg)
                try:
                    img = Image.open(image_path)
                    # Улучшаем картинку для OCR
                    img = img.resize((img.width * 2, img.height * 2))
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(3.0)
                    text = pytesseract.image_to_string(img)
                    match = re.search(r"made\s*in\s*([A-Za-z\s]+)", text, re.IGNORECASE)
                    if match:
                        country = match.group(1).strip()
                        made_in_value = country if country.upper() != "EU" else "EU"
                        break
                except Exception:
                    continue
            ws.cell(row=row_start + idx, column=5, value=made_in_value)  # E

        # --- F = G + H (численно, без формул) ---
        def _to_float(x):
            if x is None:
                return None
            if isinstance(x, (int, float)):
                return float(x)
            s = str(x).strip()
            if s == "" or s.lower() in ("nan", "none"):
                return None
            s = s.replace(" ", "").replace(",", ".")
            try:
                return float(s)
            except Exception:
                return None

        # Сколько строк считать: ориентируемся на G/H
        max_len = max(len(pl_col_E), len(pl_col_D))
        for idx in range(max_len):
            r = row_start + idx
            g_val = _to_float(ws.cell(row=r, column=7).value)  # G
            h_val = _to_float(ws.cell(row=r, column=8).value)  # H
            if g_val is None and h_val is None:
                continue
            total = (g_val or 0.0) + (h_val or 0.0)
            cell_f = ws.cell(row=r, column=6, value=total)     # F
            cell_f.number_format = "0.00"  # два знака после запятой

        wb.save(output_path)
        return True

    
class Docx_Filler:
    """
    Класс для заполнения таблицы в Word (.docx) данными из Excel через openpyxl.
    """

    def __init__(self):
        pass

    def fill_table_from_excel(self, template_path: str, excel_path: str, output_path: str, table_index: int = 0) -> bool:
        """
        Берет значения из Excel и заполняет таблицу Word:
        - D18-D47 → 3-я колонка
        - E18-E47 → 4-я колонка
        - F18-F47 → 5-я колонка
        - N18-N47 → 6-я колонка
        Начало вставки со второй строки Word таблицы.

        :param template_path: путь к шаблону Word
        :param excel_path: путь к Excel-файлу
        :param output_path: путь для сохранения Word
        :param table_index: индекс таблицы в Word
        :return: True если успешно
        """
        if not os.path.exists(template_path):
            print(f"[ERROR] Шаблон Word {template_path} не найден")
            return False
        if not os.path.exists(excel_path):
            print(f"[ERROR] Excel-файл {excel_path} не найден")
            return False

        # --- Загружаем Excel ---
        wb = load_workbook(excel_path, data_only=True)
        ws = wb.active

        # --- Читаем диапазоны ---
        values_D = [str(ws[f"D{row}"].value or "") for row in range(18, 48)]
        values_E = [str(ws[f"E{row}"].value or "") for row in range(18, 48)]
        values_F = [str(ws[f"F{row}"].value or "Error") for row in range(18, 48)]
        values_N = [str(ws[f"N{row}"].value or "") for row in range(18, 48)]
        # --- Загружаем Word ---
        doc = Document(template_path)

        try:
            table = doc.tables[table_index]
        except IndexError:
            print(f"[ERROR] Таблица с индексом {table_index} не найдена в шаблоне Word")
            return False

        # --- Вставляем значения по колонкам ---
        for i in range(len(values_D)):
            row_index = i + 1  # начинаем со второй строки
            if row_index >= len(table.rows):
                print(f"[WARNING] Недостаточно строк в таблице для строки {row_index+1}")
                continue
            table.rows[row_index].cells[2].text = values_D[i]  # 3-я колонка
            table.rows[row_index].cells[3].text = values_E[i]  # 4-я колонка
            table.rows[row_index].cells[4].text = values_F[i]  # 5-я колонка
            table.rows[row_index].cells[5].text = values_N[i]  # 6-я колонка

        # --- Сохраняем результат ---
        doc.save(output_path)
        return True