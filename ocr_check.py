from pathlib import Path
pdf = Path(r'user_inputs\\EIDP_Import_Docs\\SN 1111.pdf')
print('OCR sanity for:', pdf)
try:
    import fitz
    from PIL import Image
    import pytesseract
    doc = fitz.open(str(pdf))
    page = doc.load_page(0)
    for dpi in (300, 400, 600):
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
        for psm in (4, 6, 11):
            txt = pytesseract.image_to_string(img, lang='eng', config=f'--psm {psm}')
            print(f'--- OCR dpi={dpi} psm={psm} len={len(txt)} ---')
            print((txt or '').strip()[:600])
    doc.close()
except Exception as e:
    print('OCR failed:', e)
