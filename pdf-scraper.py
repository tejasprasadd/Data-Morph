import pdfplumber
with pdfplumber.open("All-Pdfs/001-IR-E-U-0456.pdf") as pdf:
    page=pdf.pages[0]
    text=page.extract_text()
    print(text)
