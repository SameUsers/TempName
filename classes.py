import os
import re
import math
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from PIL import Image, ImageEnhance
import pytesseract
from difflib import get_close_matches


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
        """
        Генерирует заполненный XLSX файл на основе шаблона.

        :param template_filename: Шаблон (examples/)
        :param invoice_filename: Файл invoice (xlsx_files/)
        :param ref_filename: Справочник (examples/)
        :param pl_filename: Файл PL.xlsx (xlsx_files/)
        :param spec_filename: Specifications_sell.xlsx (xlsx_files/)
        :param output_filename: Итоговый файл (examples/)
        :param photos_folder: Папка с фото товаров для OCR
        :return: True если успешно сохранено
        """

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
        pl_col_B = safe_column(pl_df, 1, replace_kan=True)
        pl_col_E = safe_column(pl_df, 4, numeric=True)
        pl_col_D = safe_column(pl_df, 3, numeric=True)
        pl_col_H = safe_column(pl_df, 7, numeric=True)
        pl_col_I = safe_column(pl_df, 8, numeric=True)

        # --- Данные из Specifications_sell.xlsx ---
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

        # --- CustomsCode -> название ---
        for idx, code in enumerate(invoice_df["CustomsCode"].fillna("").tolist()):
            ws.cell(row=row_start + idx, column=3, value=ref_map.get(code, ""))

        # --- PL.xlsx колонки ---
        for idx, val in enumerate(pl_col_B):
            ws.cell(row=row_start + idx, column=4, value=val)
        for idx, val in enumerate(pl_col_E):
            ws.cell(row=row_start + idx, column=7, value=val)
        for idx, val in enumerate(pl_col_D):
            ws.cell(row=row_start + idx, column=8, value=val)
        for idx, val in enumerate(pl_col_H):
            ws.cell(row=row_start + idx, column=10, value=math.ceil(val) if val >= 0 else math.floor(val))
        for idx, val in enumerate(pl_col_I):
            ws.cell(row=row_start + idx, column=11, value=math.ceil(val) if val >= 0 else math.floor(val))

        # --- Specifications_sell.xlsx ---
        for idx, val in enumerate(spec_values):
            ws.cell(row=row_start + idx, column=12, value=val)

        # --- Invoice, колонка H ---
        for idx, val in enumerate(invoice_col_H):
            ws.cell(row=row_start + idx, column=14, value=val)

        # --- Заполнение колонки E через OCR ---
        for idx, product_name in enumerate(pl_col_B):
            candidate_folders = os.listdir(photos_folder)
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
                    # Увеличиваем и повышаем контраст для OCR
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
            ws.cell(row=row_start + idx, column=5, value=made_in_value)

        wb.save(output_path)
        return True