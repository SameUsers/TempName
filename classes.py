import pdfplumber
import pandas as pd
import re

class Pdf_Worker:
    def __init__(self):
        pass

    def pdf_to_xlsx(self, pdf_path, xlsx_path, filter_spec=False, remove_edges=False, invoice_lines=False):
        tables_standard = []
        tables_invoice = []

        with pdfplumber.open(pdf_path) as pdf:
            if invoice_lines:
                # Паттерн строки товара
                pattern = re.compile(
                    r'^(\d+)\s+'                  # №
                    r'(\d+)\s+'                   # Code
                    r'(.*?)\s+'                   # Начало Description
                    r'(\d+)\s+'                   # Quantity
                    r'(Stk\.?\s*/\s*\d+\s*\w*)\s+'  # Unit/Volume
                    r'([\d,.]+(?:\s*/?\s*[\d,.]*)?)\s+'  # Price per unit
                    r'([\d,.]+)$'                 # TotalPrice
                )

                # 1. Собираем все строки документа в один список
                all_lines = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_lines.extend([ln.strip() for ln in text.split("\n") if ln.strip()])

                # 2. Проходим по всем строкам и склеиваем Description
                i = 0
                while i < len(all_lines):
                    line = all_lines[i]
                    match = pattern.match(line)
                    if match:
                        groups = list(match.groups())
                        description = groups[2]  # base description

                        # продолжаем смотреть вперёд
                        j = i + 1
                        while j < len(all_lines):
                            next_line = all_lines[j]
                            # если это новая позиция → стоп
                            if re.match(r'^\d+\s+\d+\s+', next_line):
                                break
                            # иначе это часть Description
                            description += " " + next_line
                            j += 1

                        groups[2] = description.strip()
                        df = pd.DataFrame([groups], columns=[
                            "№", "Code", "Description", "Quantity", "Unit/Volume", "PricePerUnit", "TotalPrice"
                        ])
                        tables_invoice.append(df)
                        i = j
                    else:
                        i += 1

            else:
                # Табличные PDF
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

        # итог
        all_tables = tables_standard + tables_invoice
        if all_tables:
            result_df = pd.concat(all_tables, axis=0, ignore_index=True, sort=False)
            result_df.to_excel(xlsx_path, index=False)
            return True
        else:
            return False