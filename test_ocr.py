from ocr import extract_text_from_pdf

# Load a PDF from disk (use any simple PDF you have)
with open("samples/sample1.pdf", "rb") as f:
    pdf_bytes = f.read()

text = extract_text_from_pdf(pdf_bytes)

print(text[:1000])  # print the first 1000 characters

