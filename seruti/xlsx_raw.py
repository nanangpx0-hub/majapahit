import zipfile, re, os, sys

def read_xlsx(path, sheet_index=0):
    z = zipfile.ZipFile(path)
    # shared strings
    ss = []
    if 'xl/sharedStrings.xml' in z.namelist():
        xml = z.read('xl/sharedStrings.xml').decode('utf-8')
        # each <si> may contain <t> or <r><t>
        for si in re.findall(r'<si>(.*?)</si>', xml, re.S):
            texts = re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)
            ss.append(''.join(texts))
    # workbook to get sheet names
    wb = z.read('xl/workbook.xml').decode('utf-8')
    sheets = re.findall(r'<sheet [^>]*name="([^"]*)" [^>]*r:id="([^"]*)"', wb)
    # rels
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid_to_target = {}
    for m in re.finditer(r'<Relationship Id="([^"]*)" Type="[^"]*" Target="([^"]*)"', rels):
        rid_to_target[m.group(1)] = m.group(2)
    # pick sheet
    name, rid = sheets[sheet_index]
    target = rid_to_target[rid]
    if not target.startswith('/'):
        target = 'xl/' + target
    else:
        target = target.lstrip('/')
    sheet_xml = z.read(target).decode('utf-8')
    # parse rows
    rows = []
    for rowm in re.finditer(r'<row [^>]*r="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.S):
        ridx = int(rowm.group(1))
        cells = rowm.group(2)
        # collect cell values by column letter
        row = {}
        for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"(?: [^>]*)?>(.*?)</c>', cells, re.S):
            col = cm.group(1)
            body = cm.group(3)
            # type
            tmatch = re.search(r't="([^"]*)"', cm.group(0))
            typ = tmatch.group(1) if tmatch else ''
            vmatch = re.search(r'<v>(.*?)</v>', body, re.S)
            val = vmatch.group(1) if vmatch else ''
            if typ == 's':
                try:
                    val = ss[int(val)]
                except Exception:
                    pass
            elif typ == 'inlineStr':
                t2 = re.search(r'<t[^>]*>(.*?)</t>', body, re.S)
                val = t2.group(1) if t2 else ''
            # unescape
            val = val.replace('&amp;','&').replace('&lt;','<').replace('&gt;','>').replace('&quot;','"')
            row[col] = val
        rows.append((ridx, row))
    return name, rows, sheets

if __name__ == '__main__':
    path = sys.argv[1]
    idx = int(sys.argv[2]) if len(sys.argv)>2 else 0
    name, rows, sheets = read_xlsx(path, idx)
    print('SHEETS:', [s[0] for s in sheets])
    print('USING:', name, 'rows:', len(rows))
    for ridx, row in rows[:40]:
        # print columns sorted
        cols = sorted(row.keys(), key=lambda c: (len(c), c))
        print(ridx, ' | '.join(f"{c}={row[c]}" for c in cols))
