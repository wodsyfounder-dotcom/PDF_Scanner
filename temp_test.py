import importlib.util
from pathlib import Path
spec = importlib.util.spec_from_file_location('core','Application/eidp_term_scanner.core.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
terms = module.load_terms(Path('user_inputs/terms.schema.smartsnap.xlsx'))
term_date = next(t for t in terms if t.term_label == 'Date Compiled')
print(module.scan_pdf_for_term_smart(Path('Data Packages/FakeProgram_SV1_SN0000.pdf'), 'SN0000', term_date, 200, False))
print(module.scan_pdf_for_term_smart(Path('Data Packages/FakeProgram_SV1_SN1111.pdf'), 'SN1111', term_date, 200, False))
