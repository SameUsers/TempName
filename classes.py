import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import re
import os


class Pdf_Worker:
    def __init__(self):
        pass

    def pdf_to_xlsx(self, pdf_path, xlsx_path,
                    filter_spec=False,
                    remove_edges=False,
                    invoice_lines=False):
        tables_standard = []
        tables_invoice = []

        with pdfplumber.open(pdf_path) as pdf:
            if invoice_lines:
                # --- Регулярка для основной строки товара ---
                pattern = re.compile(
                    r'^(\d+)\s+'                  # №
                    r'(\d+)\s+'                   # Code
                    r'(.*?)\s+'                   # Начало Description
                    r'(\d+)\s+'                   # Quantity
                    r'(Stk\.?\s*/\s*\d+\s*\w*)\s+'  # Unit/Volume
                    r'([\d,.]+(?:\s*/?\s*[\d,.]*)?)\s+'  # Price per unit
                    r'([\d,.]+)$'                 # TotalPrice
                )

                # --- Список строк всего документа ---
                all_lines = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_lines.extend([ln.strip() for ln in text.split("\n") if ln.strip()])

                # --- Стоп-слова (не добавлять в Description) ---
                skip_patterns = [
                    r'^Rechnung', r'^Kundennummer', r'^Rechnungsdatum',
                    r'^Lieferdatum', r'^IBAN', r'^BIC', r'^UST-Id',
                    r'^Zwischensumme', r'^Vortrag', r'^Seite',
                    r'^Amtsgericht', r'^HRB', r'^GF:', r'^Gläubiger-ID',
                    r'^www\.', r'^http', r'^Eissing GmbH',
                ]

                # --- Допустимые хвосты, которые идут в Description ---
                allowed_patterns = [
                    r'^EAN', r'^Art', r'^Hersteller'
                ]

                # --- Спец-хвосты, которые сохраняем в отдельные колонки ---
                customs_pattern = re.compile(r'^Zolltarif-Nr\.?:\s*(\d+)')

                # --- Парсинг ---
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

                            # если новая позиция → стоп
                            if re.match(r'^\d+\s+\d+\s+', next_line):
                                break

                            # если стоп-слово → игнорируем
                            if any(re.match(pat, next_line) for pat in skip_patterns):
                                j += 1
                                continue

                            # если это Zolltarif-Nr
                            customs_match = customs_pattern.match(next_line)
                            if customs_match:
                                customs_code = customs_match.group(1)
                                j += 1
                                continue

                            # если разрешённые хвосты → в Description
                            if any(re.match(pat, next_line) for pat in allowed_patterns):
                                description += " " + next_line
                                j += 1
                                continue

                            # иначе считаем мусором и выходим
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
                # --- Табличные PDF ---
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        df = pd.DataFrame(table)

                        if filter_spec:
                            df[0] = pd.to_numeric(df[0], errors='coerce')
                            max_val = df[0].max()
                            df = df[(df[0] >= 1) & (df[0] <= max_val)]
                            df = df.reset_index(drop=True)

                        if remove_edges:
                            if len(df) > 2:
                                df = df.iloc[1:-1].reset_index(drop=True)
                            else:
                                continue

                        tables_standard.append(df)

        # --- Итог ---
        all_tables = tables_standard + tables_invoice
        if all_tables:
            result_df = pd.concat(all_tables, axis=0, ignore_index=True, sort=False)
            result_df.to_excel(xlsx_path, index=False)
            return True
        else:
            return False
        
class File_Generate:
    def __init__(self):
        pass

    def fill_invoice(self, template_filename, invoice_filename, ref_filename, pl_filename, output_filename):
        """
        template_filename : str  - имя шаблона (.xlsx) в examples/
        invoice_filename  : str  - имя invoice файла в xlsx_files/
        ref_filename      : str  - имя справочника в examples/
        pl_filename       : str  - имя файла PL.xlsx в examples/
        output_filename   : str  - имя файла для сохранения в examples/
        """

        # Пути к файлам
        template_path = os.path.join("examples", template_filename)
        invoice_path = os.path.join("xlsx_files", invoice_filename)
        ref_path = os.path.join("examples", ref_filename)
        pl_path = os.path.join("xlsx_files", pl_filename)
        output_path = os.path.join("examples", output_filename)

        # Загружаем invoice с CustomsCode
        invoice_df = pd.read_excel(invoice_path)

        # Загружаем справочник
        ref_df = pd.read_excel(ref_path, header=None)
        ref_map = dict(zip(ref_df[3], ref_df[2]))  # D → C

        # Загружаем PL.xlsx
        pl_df = pd.read_excel(pl_path)
        pl_values = pl_df.iloc[1:, 1].astype(str).tolist()  # колонка B начиная со второй строки
        # удаляем "Kan"
        pl_values = [val.replace("Kan", "").strip() for val in pl_values]

        # Загружаем шаблон
        wb = load_workbook(template_path)
        ws = wb.active

        # Заполняем колонку C начиная с C18 (названия из справочника)
        row_start = 18
        for idx, code in enumerate(invoice_df["CustomsCode"].dropna(), start=0):
            if code in ref_map:
                ws.cell(row=row_start + idx, column=3, value=ref_map[code])
            else:
                ws.cell(row=row_start + idx, column=3, value=f"UNKNOWN {code}")

        # Заполняем колонку D начиная с D18 (данные из PL.xlsx)
        for idx, val in enumerate(pl_values):
            ws.cell(row=row_start + idx, column=4, value=val)

        # Сохраняем итоговый файл
        wb.save(output_path)
        return True