import fitz
doc = fitz.open(r'c:\Users\ti919\Documents\gemini\voice-rec\ai.pdf')
text = []
for page in doc:
    text.append(page.get_text())
with open('pdf_content2.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(text))
print("Extraction complete.")
