"""
 Num_1:
 
"""

import camelot
import pandas as pd

def extract_tables_to_excel(pdf_path, pages, output_excel):
    """
    Extracts tables from a PDF file and saves them into a single Excel file, 
    with each table in a separate sheet.

    Parameters:
    - pdf_path: str, path to the PDF file.
    - pages: str, range of pages to extract tables from (e.g., '1-5').
    - output_excel: str, path to the output Excel file.
    """

    tables = camelot.read_pdf(pdf_path, pages=pages, strip_text='\n', backend="poppler")

    print(f"Total tables extracted: {len(tables)}")

    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as excel_writer:

        for i, table in enumerate(tables):
            sheet_name = f'Table_{i+1}'
            table.df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
            print(f"Table {i+1} saved to sheet {sheet_name}")

    print(f"All tables have been saved to {output_excel}")

extract_tables_to_excel('Scholarship_Pamphlet.pdf', pages='9-32', output_excel='output_tables.xlsx')
