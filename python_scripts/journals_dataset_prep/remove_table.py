from docx import Document

# Load the Word document
doc = Document('/Users/senthil/Downloads/manuscript package/Chapters_combined copy.docx')

# Iterate through the tables in the document
for table in doc.tables:
    # Remove each table from the document
    tbl = table._element
    tbl.getparent().remove(tbl)

# Save the modified document
doc.save('/Users/senthil/Downloads/manuscript package/Chapters_combined_modified_document.docx')
