# -*- coding: utf-8 -*-
"""Parser sementara untuk mengekstrak tabel progres pengolahan Susenas (.xls = HTML)."""
import re
import json

FILES = {
    'K': r'C:/laragon/www/majapahit/susenas/data-sm2-2026/25-09-2026_154754-Progress_Pengolahan_DokM_3509.xls',
    'KP': r'C:/laragon/www/majapahit/susenas/data-sm2-2026/25-09-2026_154833-Progress_Pengolahan_DokKP_3509.xls',
}


def strip_tags(html):
    return re.sub(r'<[^>]+>', ' ', html)


def parse(path):
    with open(path, encoding='utf-8') as fh:
        html = fh.read()
    # Parser mengabaikan <tr> yang tidak seimbang: ekstrak seluruh <td>/<th> berurutan
    cells = re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', html, re.S)
    vals = [re.sub(r'\s+', ' ', strip_tags(c)).strip() for c in cells]
    data = []
    total = None
    i = 0
    while i < len(vals):
        if vals[i] == 'TOTAL' or (i + 6 < len(vals) and vals[i].startswith('TOTAL')):
            total = vals[i:i + 7]
            break
        if vals[i].isdigit() and i + 6 < len(vals):
            data.append({
                'no': int(vals[i]),
                'kode': vals[i + 1],
                'target': vals[i + 2],
                'clean_jumlah': vals[i + 3],
                'clean_pct': vals[i + 4],
                'error_jumlah': vals[i + 5],
                'error_pct': vals[i + 6],
            })
            i += 7
        else:
            i += 1
    return data, total


for key, path in FILES.items():
    data, total = parse(path)
    print('=' * 70)
    print('FILE', key, '->', path)
    print('rows parsed:', len(data))
    for d in data:
        print('  {no:>2} | {kode:<28} | tgt={target:>3} | clean={clean_jumlah:>3} {clean_pct:>9} | err={error_jumlah:>3} {error_pct:>9}'.format(**d))
    print('  TOTAL ROW:', total)
    sum_t = sum(int(d['target']) for d in data)
    sum_c = sum(int(d['clean_jumlah']) for d in data)
    sum_e = sum(int(d['error_jumlah']) for d in data)
    print('  SUMS target=%d clean=%d error=%d' % (sum_t, sum_c, sum_e))
