### convert excel to pdf
This script will:
- Read all `.xls` and `.xlsx` files in the current directory.
- Convert each Excel file to a PDF file with the same base name.
- Save the PDF files in the same directory.

Explanation:
- `pandas` is used to read the Excel files.
- `reportlab` is used to create the PDF files.
- The `excel_to_pdf` function reads the Excel file and converts its content into a PDF table.
- The `convert_all_excels_to_pdfs` function iterates over all files in the specified directory, checks if they are Excel files, and converts them to PDFs.

Test it by:
```
py convertor.py
```
