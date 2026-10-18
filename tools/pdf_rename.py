"""
Kalan Epikriz -1, Epikriz -2 dosyalarini duzenle
"""
import os
import re

KLASOR = "./pdfs"

dosyalar = [f for f in os.listdir(KLASOR) if f.lower().endswith('.pdf')]
hatalar = []
yapildi = 0

for fname in dosyalar:
    stem = os.path.splitext(fname)[0]
    
    # Protokol no
    m = re.match(r'^(\d{5,8})', stem)
    if not m:
        continue
    prot = m.group(1)
    rest = stem[len(prot):].strip()
    rest_norm = rest.lower()
    for ch in 'iışşççööüüğğ':
        rest_norm = rest_norm.replace(ch, {
            'i':'i','ı':'i','ş':'s','ç':'c','ö':'o','ü':'u','ğ':'g'
        }.get(ch, ch))

    tip = None

    # "Epikriz -1", "epikriz -1", "ePİKRİZ -1" -> E1
    if re.search(r'epikri[zs]\s+-\s*1|epikiriz\s+-\s*1', rest_norm):
        tip = 'E1'
    elif re.search(r'epikri[zs]\s+-\s*2|epikiriz\s+-\s*2', rest_norm):
        tip = 'E2'
    # "Mikrobiyoloji -2", "mikrobiyoloji 2" gibi kalan
    elif re.search(r'mikro\w*\s+-?\s*(\d+)', rest_norm):
        mm = re.search(r'mikro\w*\s+-?\s*(\d+)', rest_norm)
        tip = f'M{mm.group(1)}'
    # "mikrobyoloji" (yazim hatali) -> M1
    elif 'mikro' in rest_norm and not re.search(r'\d', rest_norm):
        tip = 'M1'
    # "DIŞ LAB"
    elif 'lab' in rest_norm:
        tip = 'LAB'

    if not tip:
        continue

    yeni_ad = f"{prot}_{tip}.pdf"
    eski_yol = os.path.join(KLASOR, fname)
    yeni_yol = os.path.join(KLASOR, yeni_ad)

    if fname == yeni_ad:
        continue

    # Cakisma
    if os.path.exists(yeni_yol):
        if tip.startswith('M'):
            for i in range(1, 20):
                alt = f"{prot}_M{i}.pdf"
                alt_yol = os.path.join(KLASOR, alt)
                if not os.path.exists(alt_yol):
                    yeni_ad = alt
                    yeni_yol = alt_yol
                    break
            else:
                hatalar.append(f"{fname} -> cakisma")
                continue
        else:
            hatalar.append(f"{fname} -> {yeni_ad} cakisma")
            continue

    try:
        os.rename(eski_yol, yeni_yol)
        yapildi += 1
        print(f"  {fname} -> {yeni_ad}")
    except Exception as e:
        hatalar.append(f"{fname}: {e}")

print(f"\n{yapildi} dosya yeniden adlandirildi.")
if hatalar:
    print(f"{len(hatalar)} hata/atlanan:")
    for h in hatalar:
        print(f"  {h}")
