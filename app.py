import pandas as pd
from docx import Document

# Load the Excel file
excel_path = "data.xlsx"
df_journal = pd.read_excel(excel_path, sheet_name="Journal Publication", skiprows=5)

# Fix column names (remove spaces)
df_journal.columns = df_journal.columns.str.strip()

# Check actual column names
print("Excel Columns:", df_journal.columns.tolist())

# Filter data for a particular faculty member (Remove extra spaces if any)
name1 = "Dr. Thenmozhi T"
df_filtered = df_journal[df_journal["Name of the faculty"].str.strip() == name1]

# Load the Word document
doc_path = "template.docx"
doc = Document(doc_path)

# Verify table index
for i, table in enumerate(doc.tables):
    print(f"\nTable {i}:")
    for row in table.rows[:2]:  # Print first 2 rows only
        print([cell.text.strip() for cell in row.cells])

table_index = 3  # Update if needed
table = doc.tables[table_index]

start_row = 1
# Fill the table with Excel data
for i, (_, row) in enumerate(df_filtered.iterrows()):
    row_index = start_row+i
    if i+1 >= len(table.rows):  
        table.add_row()  # Add rows if needed

    table.cell(i+2, 0).text = str(i+1)  # Serial No.
    table.cell(i+2, 1).text = str(row.get("Paper Title", "N/A"))
    table.cell(i+2, 2).text = str(row.get("Journal Name", "N/A"))
    table.cell(i+2, 3).text = str(row.get("Year of Publication", "N/A"))  # Fixed Date Mapping
    table.cell(i+2, 4).text = str(row.get("ISSN", "N/A"))
    table.cell(i+2, 5).text = str(row.get("Citation", "N/A"))  # Citation was missing in mapping
    table.cell(i+2, 6).text = str(row.get("Impact Factor", "N/A"))

# Save the modified document
output_doc_path = "filled_template.docx"
doc.save(output_doc_path)
print(f"✅ Word document saved as {output_doc_path}")
