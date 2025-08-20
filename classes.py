import os
import re
import math
import pdfplumber
import pandas as pd
from openpyxl import load_workbook


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
                     output_filename: str) -> bool:
        """
        Генерирует заполненный XLSX файл на основе шаблона.

        :param template_filename: Шаблон (examples/)
        :param invoice_filename: Файл invoice (xlsx_files/)
        :param ref_filename: Справочник (examples/)
        :param pl_filename: Файл PL.xlsx (xlsx_files/)
        :param spec_filename: Specifications_sell.xlsx (xlsx_files/)
        :param output_filename: Итоговый файл (examples/)
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

        # --- Справочник CustomsCode -> Название ---
        ref_map = dict(zip(ref_df[3], ref_df[2]))

        # --- Универсальная функция для получения колонки ---
        def safe_column(df, col_index, numeric=False, replace_kan=False):
            """Возвращает все значения колонки, начиная с первой строки данных, корректно обрабатывая пустые/NaN"""
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

        # C18: CustomsCode -> Название
        for idx, code in enumerate(invoice_df["CustomsCode"].fillna("").tolist()):
            ws.cell(row=row_start + idx, column=3, value=ref_map.get(code, ""))

        # D18: PL B
        for idx, val in enumerate(pl_col_B):
            ws.cell(row=row_start + idx, column=4, value=val)

        # G18: PL E
        for idx, val in enumerate(pl_col_E):
            ws.cell(row=row_start + idx, column=7, value=val)

        # H18: PL D
        for idx, val in enumerate(pl_col_D):
            ws.cell(row=row_start + idx, column=8, value=val)

        # J18: PL H с округлением
        for idx, val in enumerate(pl_col_H):
            ws.cell(row=row_start + idx, column=10, value=math.ceil(val) if val >= 0 else math.floor(val))

        # K18: PL I с округлением
        for idx, val in enumerate(pl_col_I):
            ws.cell(row=row_start + idx, column=11, value=math.ceil(val) if val >= 0 else math.floor(val))

        # L18: Specifications_sell.xlsx
        for idx, val in enumerate(spec_values):
            ws.cell(row=row_start + idx, column=12, value=val)

        # N18: Invoice, колонка H
        for idx, val in enumerate(invoice_col_H):
            ws.cell(row=row_start + idx, column=14, value=val)

        # --- Сохраняем итоговый файл ---
        wb.save(output_path)
        return True