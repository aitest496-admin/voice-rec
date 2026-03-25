import fitz
import os

pdf_path = r'c:\Users\ti919\Documents\gemini\voice-rec\ai.pdf'
doc = fitz.open(pdf_path)

print(f"Total pages: {len(doc)}")
# Render up to 5 pages to avoid out-of-memory or too many files
for i in range(min(5, len(doc))):
    page = doc.load_page(i)
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x scale for better readability
    out_path = f"c:\\Users\\ti919\\Documents\\gemini\\voice-rec\\page_{i}.png"
    pix.save(out_path)
    print(f"Saved {out_path}")
