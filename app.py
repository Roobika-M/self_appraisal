import pandas as pd
from docx import Document

# Load the Excel file
excel_path = "data.xlsx"
doc_path = "template.docx"

# Load the Word document template
doc = Document(doc_path)

# Load and clean the "Journal Publication" sheet
df_journal = pd.read_excel(excel_path, sheet_name="Journal Publication", skiprows=4)

# Rename columns based on actual headers
df_journal.columns = [
    "Faculty Name", "S.No", "Journal Name", "Paper Title", "Author Name",
    "Volume Number", "Issue Number", "Page Number From", "Page Number To",
    "ISSN", "Citation", "Year of Publication", "Web Link", "Impact Factor"
]

# Drop any fully empty rows
df_journal = df_journal.dropna(how="all")

# Remove the first row (repeated headers)
df_journal_cleaned = df_journal.iloc[1:][[
    "S.No", "Journal Name", "Paper Title", "Author Name", 
    "ISSN", "Impact Factor"
]]

# Function to add a table to the Word document
def add_table_no_style(doc, data, section_title, headers):
    doc.add_paragraph(section_title)  # Add section title

    if not data.empty:
        table = doc.add_table(rows=1, cols=len(headers))  # Create table

        # Add headers
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            hdr_cells[i].text = header

        # Add rows
        for _, row in data.iterrows():
            cells = table.add_row().cells
            for i, col in enumerate(row):
                cells[i].text = str(col) if pd.notna(col) else "-"
        doc.add_paragraph("\n")
    else:
        doc.add_paragraph("No data available.\n")

# Add Journal Publications to the Word document
add_table_no_style(doc, df_journal_cleaned, "2.1 No. of Journal Publications", 
                   ["S.No", "Journal Name", "Paper Title", "Author Name", "ISSN", "Impact Factor"])

# Save the modified document
output_path = "filled_research_template.docx"
doc.save(output_path)

print(f"Template filled and saved as: {output_path}")
