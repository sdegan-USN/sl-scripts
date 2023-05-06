import pandas as pd
import docx

# open the Word document
doc = docx.Document('output.docx')

# create an empty list to store the tables
tables = []

# loop through each table in the document
for i, table in enumerate(doc.tables):

    # get the first row of the table (the header)
    header_row = table.rows[0]

    # check if at least one cell in the header row has white text
    has_white_text = any(cell.text.strip() == '' or cell.paragraphs[0].runs[0].font.color.rgb == docx.shared.RGBColor(255, 255, 255) for cell in header_row.cells)

    # if the header row has at least one cell with white text, add the table to the list
    if has_white_text:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text)
            table_data.append(row_data)
        tables.append(pd.DataFrame(table_data[1:], columns=table_data[0]))

# create a new Excel file and write the tables to separate sheets
with pd.ExcelWriter('output.xlsx') as writer:
    for i, table in enumerate(tables):
        table.to_excel(writer, sheet_name=f'Table {i+1}', index=False)

