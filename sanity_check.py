from pathlib import Path
pdf = Path(r'user_inputs\\EIDP_Import_Docs\\SN 1111.pdf')
print('PDF exists:', pdf.exists(), pdf)
try:
    import fitz
    doc = fitz.open(str(pdf))
    print('Pages:', doc.page_count)
    p = 0
    txt = doc.load_page(p).get_text('text')
    print('PyMuPDF text len:', len(txt))
    print('--- PyMuPDF sample ---')
    print((txt or '').strip()[:500])
    doc.close()
except Exception as e:
    print('PyMuPDF failed:', e)
try:
    from pdfminer.high_level import extract_text
    t = extract_text(str(pdf), page_numbers=[0])
    print('pdfminer text len:', len(t))
    print('--- pdfminer sample ---')
    print((t or '').strip()[:500])
except Exception as e:
    print('pdfminer failed:', e)
try:
    try:
        from pypdf import PdfReader
    except Exception:
        from PyPDF2 import PdfReader
    r = PdfReader(str(pdf))
    pg = r.pages[0]
    s = pg.extract_text() or ''
    print('pypdf text len:', len(s))
    print('--- pypdf sample ---')
    print(s.strip()[:500])
except Exception as e:
    print('pypdf failed:', e)
