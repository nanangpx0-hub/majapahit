import os
import pymupdf
from PIL import Image, ImageDraw

PDF = r'C:/laragon/www/majapahit/susenas/3509_Laporan_Pengolahan_Susenas_September_2026_50Persen.pdf'
OUT = r'C:/laragon/www/majapahit/susenas/_sheet'
os.makedirs(OUT, exist_ok=True)

doc = pymupdf.open(PDF)
DPI = 100
imgs = {}
for i in range(doc.page_count):
    pix = doc[i].get_pixmap(dpi=DPI)
    imgs[i + 1] = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)

SHEETS = [
    ('sheetA_babVII', [6, 7]),
    ('sheetB_pengesahan_lampiran', [9, 10]),
    ('sheetC_lampiran_tabel_gambar', [11, 12]),
    ('sheetD_babVI_babV', [4, 5]),
]

for name, pages in SHEETS:
    sel = [imgs[p] for p in pages]
    w = max(im.width for im in sel)
    h = max(im.height for im in sel)
    label_h = 34
    sheet = Image.new('RGB', (w, (h + label_h) * len(sel)), 'white')
    d = ImageDraw.Draw(sheet)
    for k, (p, im) in enumerate(zip(pages, sel)):
        y = k * (h + label_h)
        d.rectangle([0, y, w, y + label_h], fill=(30, 60, 120))
        d.text((10, y + 10), f'PAGE {p}', fill='white')
        sheet.paste(im, (0, y + label_h))
    path = f'{OUT}/{name}.png'
    sheet.save(path)
    print('saved', path, sheet.size)

print('sheets ready')
