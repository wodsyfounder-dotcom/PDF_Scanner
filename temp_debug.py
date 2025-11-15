import importlib.util
from pathlib import Path
spec = importlib.util.spec_from_file_location('core','Application/eidp_term_scanner.core.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
terms = module.load_terms(Path('user_inputs/terms.schema.smartsnap.xlsx'))
term = next(t for t in terms if t.term_label == 'Date Compiled')
print('format', term.value_format)
res = module.scan_pdf_for_term_smart(Path('Data Packages/FakeProgram_SV1_SN0000.pdf'), 'SN0000', term, 200, False)
print(res)
