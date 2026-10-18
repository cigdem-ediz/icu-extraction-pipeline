"""
amr_mapper.py — Kohort antimikrobiyallerini TİTCK + WHO AWaRe ile eşleştirir.
================================================================================
Zincir:  PDF ham ilaç adı  →  TİTCK (ürün adı fuzzy)  →  etken madde + ATC
         ATC  →  WHO AWaRe 2023  →  Access/Watch/Reserve
         ATC  →  DDD tablosu (elle, WHO ATC/DDD Index'ten)  →  DDD

Kullanım:
    python amr_mapper.py --titck TITCK.xlsx --aware WHO_AWARE.xlsx \
                         --ilaclar ilac_listesi.txt --ddd ddd_tablosu.csv \
                         --cikti eslesme_sonuc.xlsx
"""
import re, sys, argparse
from difflib import SequenceMatcher
import openpyxl


# ── 1) Ham ilaç adından marka adını ayıkla ────────────────────────────────────
def marka_ayikla(ham):
    """
    'COLİMYCİN İM/İV Enj. ve İnh. İçin Liy. Toz...' -> 'COLİMYCİN'
    İlk kelime(ler)i alır; doz/form/birim bilgisini atar.
    """
    if not ham:
        return ""
    s = str(ham).strip().upper()
    # Türkçe büyük harf normalizasyonu
    s = s.replace("İ", "I").replace("Ş", "S").replace("Ğ", "G") \
         .replace("Ü", "U").replace("Ö", "O").replace("Ç", "C")
    # Dozaj/form anahtar kelimeleri: bunlardan öncesini marka kabul et
    stop = re.compile(r'\b(IV|IM|İV|İM|ENJ|FLAKON|FLK|AMPUL|AMP|TABLET|TAB|'
                      r'INF|İNF|INFUZYON|SOLUSYON|COZELTI|LIY|TOZ|SASE|TORBA|'
                      r'MG|ML|GR|G|IU|LIYOFILIZE|ENJEKTABL|KAPSUL|%)\b', re.I)
    m = stop.search(s)
    if m:
        s = s[:m.start()]
    # Sayıları at, fazla boşlukları temizle
    s = re.sub(r'\d+', ' ', s)
    s = re.sub(r'[^\w\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# ── 2) TİTCK Excel'ini sözlüğe çevir ──────────────────────────────────────────
def titck_yukle(path):
    """
    TİTCK ürün Excel'ini oku. Beklenen sütunlar:
    ÜRÜN ADI | ETKİN MADDE | ATC KODU  (başlık satırı otomatik bulunur)
    Döner: list[ (urun_adi_normalize, urun_adi_ham, etken, atc) ]
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Başlık satırını bul
    hdr_idx = None
    for i, r in enumerate(rows[:40]):
        cells = [str(c).upper().strip() if c else "" for c in r]
        if any("ÜRÜN AD" in c or "URUN AD" in c for c in cells):
            hdr_idx = i
            header = cells
            break
    if hdr_idx is None:
        raise ValueError("TİTCK başlık satırı bulunamadı (ÜRÜN ADI sütunu yok)")

    def col(anahtarlar):
        for j, h in enumerate(header):
            if any(a in h for a in anahtarlar):
                return j
        return None

    c_urun = col(["ÜRÜN AD", "URUN AD"])
    c_etk  = col(["ETKİN MADDE", "ETKIN MADDE"])
    c_atc  = col(["ATC KOD"])

    if None in (c_urun, c_etk, c_atc):
        raise ValueError(f"Sütun bulunamadı: urun={c_urun}, etken={c_etk}, atc={c_atc}")

    kayitlar = []
    for r in rows[hdr_idx+1:]:
        if not r or c_urun >= len(r):
            continue
        urun = r[c_urun]
        if not urun:
            continue
        etken = r[c_etk] if c_etk < len(r) else ""
        atc   = r[c_atc] if c_atc < len(r) else ""
        kayitlar.append((marka_ayikla(urun), str(urun).strip(),
                         str(etken or "").strip(), str(atc or "").strip()))
    return kayitlar


# ── 3) WHO AWaRe Excel'ini ATC->sınıf sözlüğüne çevir ─────────────────────────
def aware_yukle(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["AWaRe classification 2023"]
    rows = list(ws.iter_rows(values_only=True))
    # Başlık: 'Antibiotic | Class | ATC code | Category'
    atc2cat = {}
    basladi = False
    for r in rows:
        if not r:
            continue
        if r[0] and str(r[0]).strip() == "Antibiotic":
            basladi = True
            continue
        if basladi and r[2]:  # ATC code sütunu
            atc = str(r[2]).strip().upper()
            cat = str(r[3]).strip() if len(r) > 3 and r[3] else ""
            if atc:
                atc2cat[atc] = cat
    return atc2cat


# ── 4) Fuzzy eşleştirme ───────────────────────────────────────────────────────
def en_iyi_eslesme(marka, titck_kayitlar):
    """marka (normalize) -> en yüksek benzerlikteki TİTCK kaydı + skor"""
    en_iyi = None
    en_skor = 0.0
    for norm, ham, etken, atc in titck_kayitlar:
        if not norm:
            continue
        # Tam kelime içeriyorsa bonus
        if marka in norm.split() or norm in marka.split():
            skor = 0.95
        else:
            skor = SequenceMatcher(None, marka, norm).ratio()
        if skor > en_skor:
            en_skor = skor
            en_iyi = (ham, etken, atc)
    return en_iyi, en_skor


# ── 5) Ana akış ───────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--titck", required=True)
    ap.add_argument("--aware", required=True)
    ap.add_argument("--ilaclar", required=True, help="Ham ilaç adları, her satır bir isim")
    ap.add_argument("--ddd", default=None, help="ATC,DDD,birim,yol içeren CSV (opsiyonel)")
    ap.add_argument("--cikti", default="eslesme_sonuc.xlsx")
    ap.add_argument("--esik", type=float, default=0.6, help="Düşük skor uyarı eşiği")
    args = ap.parse_args()

    print("TİTCK yükleniyor..."); titck = titck_yukle(args.titck)
    print(f"  {len(titck)} ürün")
    print("WHO AWaRe yükleniyor..."); atc2cat = aware_yukle(args.aware)
    print(f"  {len(atc2cat)} ATC-sınıf")

    ddd_map = {}
    if args.ddd:
        with open(args.ddd, encoding="utf-8") as f:
            for line in f:
                p = [x.strip() for x in line.split(",")]
                if len(p) >= 2 and p[0].upper().startswith("J"):
                    ddd_map[p[0].upper()] = tuple(p[1:])

    with open(args.ilaclar, encoding="utf-8") as f:
        ham_liste = [l.strip() for l in f if l.strip()]
    # Benzersizleştir
    ham_benzersiz = sorted(set(ham_liste))

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Eslesme"
    ws.append(["Ham İlaç Adı", "Ayıklanan Marka", "TİTCK Ürün",
               "Etken Madde", "ATC", "AWaRe", "DDD", "Skor", "KONTROL"])
    for ham in ham_benzersiz:
        marka = marka_ayikla(ham)
        eslesme, skor = en_iyi_eslesme(marka, titck)
        if eslesme:
            titck_urun, etken, atc = eslesme
            aware = atc2cat.get(atc.upper(), "—")
            ddd = ddd_map.get(atc.upper(), ("",))[0] if ddd_map else ""
        else:
            titck_urun = etken = atc = aware = ddd = ""
        kontrol = "GÖZDEN GEÇİR" if skor < args.esik else ""
        ws.append([ham, marka, titck_urun, etken, atc, aware, ddd,
                   round(skor, 2), kontrol])

    wb.save(args.cikti)
    print(f"\n✓ Sonuç: {args.cikti} ({len(ham_benzersiz)} benzersiz ilaç)")
    print("  Skoru düşük (KONTROL='GÖZDEN GEÇİR') satırları elle doğrulayın.")


if __name__ == "__main__":
    main()
