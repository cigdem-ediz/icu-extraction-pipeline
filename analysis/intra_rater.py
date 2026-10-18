"""Intra-rater agreement: first manual abstraction vs. second blind re-abstraction.

Kullanim:  python intra_rater.py ICU_VeriTabani-manuel.xlsx ikinci_cikarim_SABLON.xlsx
Satirlar sekme icinde sira ile eslestirilir (ikinci sablon ilk cikarimin satir yapisini kullanir).
Mikrobiyoloji ikinci kez yalniz 1 hastada cikarildigi icin disarida.
Normalizasyon karsilastir.py ile aynidir (NFC, Turkce I katlama, ondalik, ISO tarih).
"""
import sys, re, math, random, datetime, unicodedata, collections
import openpyxl

SHEETS = ['DEMOGRAFI', 'KLINIK', 'LABORATUVAR', 'KAN_GAZI', 'ANTIMIKROBIAL', 'VAZOPRESSOR']
# Turetilmis alanlar, metadata ve klinik yargi alanlari karsilastirmaya girmez
EXCL = {'Hasta_ID', 'Yas', 'Yatis_Suresi_Gun', 'Gun_No', 'Bas_Gun_No', 'Baslangic_Gun_No', 'Bitis_Gun_No',
        'MIC', 'Kaynak', 'Zaman_Noktasi', 'Tarih', 'TDM', 'Tedavi_Turu', 'Uygunluk',
        'Uygunsuzluk_Nedeni', 'Deskalasyon', 'Eskalasyon_Nedeni'}


def load(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb}


def norm(v):
    if v is None:
        return ''
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = unicodedata.normalize('NFC', str(v)).strip()
    for ch in 'İIı':
        s = s.replace(ch, 'i')
    s = re.sub(r'\s+', ' ', s.lower())
    if re.fullmatch(r'-?\d+,\d+', s):
        s = s.replace(',', '.')
    try:
        f = float(s)
        s = str(int(f)) if f.is_integer() else repr(round(f, 6))
    except ValueError:
        pass
    if re.fullmatch(r'\d{4}-\d\d-\d\d 00:00:00', s):
        s = s[:10]
    return '' if s in ('none', '—', '-', 'nan', '?') else s


def wilson(x, n, z=1.96):
    p = x / n
    d = 1 + z * z / n
    c = p + z * z / (2 * n)
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return 100 * (c - h) / d, 100 * (c + h) / d


def main(first, second):
    A, B = load(first), load(second)
    prots = sorted({str(r[0]) for r in A['DEMOGRAFI'][1:]})
    pmap = {p: f'ICU-{i + 1:03d}' for i, p in enumerate(prots)}
    per_patient = collections.defaultdict(lambda: [0, 0])
    total = [0, 0]
    print(f"{'Domain':16s}{'Fields':>7s}{'Pairs':>7s}{'Agree':>7s}{'%':>7s}  95% CI")
    for sh in SHEETS:
        ha, hb = A[sh][0], B[sh][0]
        cols = [c for c in hb if c in ha and c not in EXCL]
        ra = [r for r in A[sh][1:] if r[0] is not None and str(r[0]).strip() and not str(r[0]).startswith('★')]
        rb = [r for r in B[sh][1:] if r[0] is not None and str(r[0]).strip()]
        if sh != 'KAN_GAZI':
            rb = [r for r in rb if any(v not in (None, '') for v in r[1:])]
        n = x = 0
        for a, b in zip(ra, rb):
            pid = pmap.get(str(a[0]).strip(), str(a[0]).strip())
            for c in cols:
                va, vb = norm(a[ha.index(c)]), norm(b[hb.index(c)])
                if va == '' and vb == '':
                    continue
                n += 1
                per_patient[pid][0] += 1
                if va == vb:
                    x += 1
                    per_patient[pid][1] += 1
                else:
                    print(f'  fark: {sh} {pid} {c}')
        total[0] += n
        total[1] += x
        lo, hi = wilson(x, n)
        print(f'{sh:16s}{len(cols):7d}{n:7d}{x:7d}{100 * x / n:7.1f}  {lo:.1f}-{hi:.1f}')
    lo, hi = wilson(total[1], total[0])
    print(f"{'TOTAL':16s}{'':7s}{total[0]:7d}{total[1]:7d}{100 * total[1] / total[0]:7.1f}  {lo:.1f}-{hi:.1f}")
    random.seed(2025)
    ids = list(per_patient)
    boots = []
    for _ in range(2000):
        s = [random.choice(ids) for _ in ids]
        boots.append(100 * sum(per_patient[p][1] for p in s) / sum(per_patient[p][0] for p in s))
    boots.sort()
    print(f'Cluster bootstrap 95% CI: {boots[49]:.1f}-{boots[1949]:.1f}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
