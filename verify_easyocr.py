import sys
print('Python:', sys.version)
try:
    import easyocr
    print('easyocr version:', easyocr.__version__)
    r = easyocr.Reader(['en'], gpu=False, verbose=False)
    print('Reader OK; langs:', r.lang_list)
    print('OK')
except Exception as e:
    import traceback
    traceback.print_exc()
    sys.exit(1)
