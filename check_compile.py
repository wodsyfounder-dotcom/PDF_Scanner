import py_compile, traceback
try:
    py_compile.compile(r"scripts/easyocr_dump_spatial.py", doraise=True)
    print("OK")
except Exception as e:
    traceback.print_exc()
