import tokenize
p='scripts/easyocr_dump_spatial.py'
with open(p,'rb') as f:
    for tok in tokenize.tokenize(f.readline):
        if 68 <= tok.start[0] <= 90:
            print(tok)
