# -*- coding: utf-8 -*-
"""
Generate fully SYNTHETIC example patients (E1 / E2 / LAB / M1) for demonstrating
the pipeline. Everything here is invented — names, protocol numbers, national IDs,
dates, laboratory values, and microbiology results — and no real institution is
named. Every page carries a visible synthetic banner.

Daily notes are built from a parameterized template so each patient has a multi-day
course with realistic, real-length clinical narrative, while remaining entirely
fictional. The structure mirrors real ICU discharge summaries so `icu_pipeline.py`
can parse the output. Edit the PATIENTS list and regenerate at any time.

Run:  python generate_synthetic.py
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import textwrap, datetime

pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

BANNER = "SYNTHETIC EXAMPLE - NOT REAL PATIENT DATA / SENTETİK ÖRNEK - GERÇEK HASTA VERİSİ DEĞİLDİR"

def make_pdf(path, lines):
    c = canvas.Canvas(path, pagesize=A4); w, h = A4; y = h - 2*cm
    def header():
        c.setFont("DejaVu-Bold", 7); c.drawString(2*cm, h-1.1*cm, BANNER); c.setFont("DejaVu", 8.5)
    header()
    for line in lines:
        for seg in (textwrap.wrap(line, 108) or [""]):
            if y < 2*cm:
                c.showPage(); header(); y = h-2*cm
            c.drawString(2*cm, y, seg); y -= 0.4*cm
    c.showPage(); c.save(); print("wrote", path)

def daily_note(d):
    """Build one dense, realistic-length ICU daily note from invented parameters."""
    return (
        f"{d['tani']} tanılarıyla izlenen hastanın genel durumu {d['gd']}, GKS "
        f"{sum(d['gks'])}, pupiller izokorik, IR +/+. {d['vent']} mekanik ventilatör "
        f"desteğinde; dinlemekle her iki akciğer solunum sesleri kaba ve bilateral azalmış. "
        f"TA {d['ta']} mmHg, nabız {d['nabiz']}/dk, ateş {d['ates']} C. Laboratuvar: "
        f"lökosit {d['wbc']}, CRP {d['crp']}, kreatinin {d['kre']}, üre {d['ure']}, "
        f"laktat {d['laktat']}, trombosit {d['plt']}. {d['hemodinamik']} "
        f"Antibiyoterapi enfeksiyon hastalıkları onayıyla {d['abx']}. "
        f"Kan gazı: ph {d['ph']} pco2 {d['pco2']} po2 {d['po2']} hco3 {d['hco3']} be {d['be']} "
        f"spo2 {d['spo2']} sodyum {d['na']} potasyum {d['k']}. Kan şekerleri {d['seker']} mg/dl "
        f"seyretti, kristalize insülin ile regüle edildi. İdrar çıkışı diüretik infüzyonu ile "
        f"yeterli, AÇT {d['act']}, CVP {d['cvp']}. {d['beslenme']} Solunum fizyoterapisi, vücut "
        f"ve kateter bakımları yapıldı. Mevcut medikal, hemodinamik ve solunum destek "
        f"tedavisine devam edildi."
    )

def e1(p):
    L = ["SENTETİK ÖRNEK HASTANESİ", "YOĞUN BAKIM EPİKRİZ RAPORU", "",
         f"Adı - Soyadı : {p['ad']}    Prot / Dosya No : {p['proto']} / {p['dosya']}",
         "TC Kimlik No : 11111111111",
         f"Cinsiyeti - Yaş : {p['cinsiyet']} - {p['yas']}",
         f"D.Tarihi : {p['dtarih']}",
         "Kurum : ÖZEL",
         f"Bölüm Adı : Genel Yoğun Bakım    Yatış Tarihi : {p['gunler'][0]['tarih']}",
         f"Çıkış Tarihi : {p['gunler'][-1]['tarih']}",
         f"Taburcu Şekli : {p['taburcu']}",
         f"Kesin Tanı : {p['tani_kesin']}",
         f"Apache II skoru: {p['apache']}",
         f"tanısal ölüm oranı %{p['mortalite']}",
         p['giris'], ""]
    for i, d in enumerate(p['gunler'], 1):
        L.append(f"{d['tarih']} {i}.GÜN")
        L.append(daily_note(d))
        L.append(f"GLASGOW Skalası  Gözler : {d['gks'][0]}  Sözel : {d['gks'][1]}  Motor : {d['gks'][2]}")
        L.append("")
    return L

def e2(p):
    L = ["SENTETİK ÖRNEK HASTANESİ", "Epikriz Yoğun Bakım Formu", "",
         f"Adı - Soyadı : {p['ad']}    Prot / Dosya No : {p['proto']} / {p['dosya']}",
         f"Cinsiyeti - Yaş : {p['cinsiyet']} - {p['yas']}",
         f"Yatış Tarihi : {p['gunler'][0]['tarih']}", ""]
    for d in p['gunler']:
        L += [f"İzlem Tarihi : {d['tarih']}",
              "Yoğun Bakım Seviye Bilgisi : 3. BASAMAK",
              f"Septik Şok : {d['septik']}   Sepsis Durumu : {d['sepsis']}",
              "GLASGOW Skalası",
              f"Gözler : {d['gks'][0]}   Sözel : {d['gks'][1]}   Motor : {d['gks'][2]}", ""]
    return L

def lab(p):
    tarihler = "   ".join(d['tarih'] for d in p['gunler'])
    L = ["SENTETİK ÖRNEK HASTANESİ", "TIBBİ LABORATUVAR TETKİK SONUÇ RAPORU", "",
         f"Adı - Soyadı : {p['ad']}    Prot / Dosya No : {p['proto']} / {p['dosya']}", "",
         "        " + tarihler]
    # lab satırları: her parametre gün sayısı kadar değer
    for etiket, vals in p['lab'].items():
        L.append(f"{etiket:22s} " + "   ".join(str(v) for v in vals))
    return L

def m1(p):
    L = ["SENTETİK ÖRNEK HASTANESİ",
         "TIBBİ LABORATUVAR TETKİK SONUÇ RAPORU (TIBBİ MİKROBİYOLOJİ)", "",
         f"Adı - Soyadı : {p['ad']}    Prot / Dosya No : {p['proto']} / {p['dosya']}",
         "TC Kimlik No : 11111111111",
         f"{p['ad']} {p['mikro']['numune']} {p['gunler'][1]['tarih']} 09:20:00 {p['gunler'][1]['tarih']} 11:20:00",
         "Genel Yoğun Bakım", ""]
    if p['mikro']['organizma']:
        L.append(f"MİKROORGANİZMA {p['mikro']['organizma']} Koloni Sayısı 100000")
        L.append("ANTİBİYOGRAM")
        for ab, (durum, mic) in p['mikro']['antibiyogram']:
            L.append(f"{ab} Mic: {mic} {durum}")
    else:
        L.append(f"AÇIKLAMA {p['mikro']['aciklama']}")
    return L

# ---- yardımcı: gün üretici (uydurma değerlerle dolu bir gün) ----
def gun(tarih, tani, gd, gks, vent, ta, nabiz, ates, wbc, crp, kre, ure, laktat, plt,
        hemodinamik, abx, ph, pco2, po2, hco3, be, spo2, na, k, seker, act, cvp, beslenme,
        septik, sepsis):
    return dict(tarih=tarih, tani=tani, gd=gd, gks=gks, vent=vent, ta=ta, nabiz=nabiz,
                ates=ates, wbc=wbc, crp=crp, kre=kre, ure=ure, laktat=laktat, plt=plt,
                hemodinamik=hemodinamik, abx=abx, ph=ph, pco2=pco2, po2=po2, hco3=hco3, be=be,
                spo2=spo2, na=na, k=k, seker=seker, act=act, cvp=cvp, beslenme=beslenme,
                septik=septik, sepsis=sepsis)

PATIENTS = [
 {"ad":"ÖRNEK HASTA BİR","proto":"99001","dosya":"700001","cinsiyet":"ERKEK","yas":"64",
  "dtarih":"12.03.1961","taburcu":"Şifa ile taburcu",
  "tani_kesin":"Pnömosepsis, akut böbrek yetmezliği, KOAH",
  "apache":"22","mortalite":"42",
  "giris":"Bilinen KOAH ve hipertansiyon tanılı hasta, nefes darlığı ve ateş ile acile "
          "başvurdu; solunum yetmezliği ve pnömosepsis nedeniyle yoğun bakıma alındı, entübe "
          "edildi. Sağ subklaviyen santral venöz kateter yerleştirildi, invaziv ve noninvaziv "
          "monitörizasyon sağlandı.",
  "lab":{"WBC (Lökosit)":[18.2,16.9,15.6,13.8,12.4,10.6,9.8,8.1],
         "HGB (Hemoglobin)":[9.1,9.2,9.4,9.8,10.2,10.5,10.8,11.0],
         "CRP (TÜRBIDIMETRIK)":[210,205,180,150,120,95,70,35],
         "KREATİNİN (SERUM)":[2.4,2.3,2.1,1.9,1.8,1.6,1.4,1.2],
         "BUN (KAN ÜRE AZOTU)":[48,45,42,37,32,28,26,22],
         "PLT (Trombosit)":[95,100,110,140,180,220,260,320],
         "CLAC":[2.4,2.3,2.1,2.0,1.9,1.7,1.5,1.2]},
  "gunler":[
    gun("05.04.2023","Pnömosepsis, akut böbrek yetmezliği ve solunum yetmezliği","kötü",(2,1,4),
        "Entübe, basınç kontrol modunda (FİO2 %90, RR 16, PEEP 6)","110/60","112","38.4",
        "18200","210","2.4","138","2.4","95000",
        "Belirgin hipotansiyon nedeniyle sıvı resüsitasyonu sonrası noradrenalin infüzyonu başlandı.",
        "Meronem 3x1gr (1.gün)","7.31","46","78","21","-3","92","141","4.2","180-220",
        "+120/24saat","+2","Parenteral beslenme (+), defekasyon (-).","EVET","EVET"),
    gun("06.04.2023","Pnömosepsis, akut böbrek yetmezliği","kötü",(2,1,4),
        "Entübe, basınç kontrol modunda (FİO2 %80, RR 15, PEEP 6)","112/62","108","38.2",
        "16900","205","2.3","132","2.3","100000",
        "Noradrenalin infüzyonu sürüyor, doz artırıldı.",
        "Meronem 3x1gr (2.gün)","7.32","45","80","21","-3","92","140","4.2","175-215",
        "+180/24saat","+2","Parenteral beslenme (+), defekasyon (-).","EVET","EVET"),
    gun("07.04.2023","Pnömosepsis ve akut böbrek yetmezliği","orta-kötü",(2,1,4),
        "Entübe, basınç destek modunda (FİO2 %60, PEEP 6, PEEP üstü 14)","115/64","104","38.0",
        "16400","195","2.2","128","2.2","110000",
        "Noradrenalin infüzyonu mevcut, doz sabit.",
        "Meronem 3x1gr (3.gün)","7.33","44","82","22","-2","93","140","4.1","170-210",
        "+240/24saat","+3","Enteral beslenmeye geçildi, tolere etti.","EVET","EVET"),
    gun("08.04.2023","Pnömosepsis ve akut böbrek yetmezliği","orta-kötü",(3,1,4),
        "Entübe, basınç destek modunda (FİO2 %55, PEEP 6, PEEP üstü 13)","116/66","100","37.9",
        "13800","150","1.9","104","2.0","140000",
        "Noradrenalin kademeli azaltıldı, idrar çıkışı arttı.",
        "Meronem 3x1gr (4.gün), Tazocin 4x4.5gr (1.gün)","7.35","43","86","22","-1","94","139",
        "4.1","160-200","+300/24saat","+3","Enteral beslenme (+).","EVET","EVET"),
    gun("09.04.2023","Pnömosepsis ve akut böbrek yetmezliği","orta",(3,1,5),
        "Entübe, basınç destek modunda (FİO2 %50, PEEP 6, PEEP üstü 12)","118/68","98","37.8",
        "12400","120","1.8","96","1.9","180000",
        "Antibiyograma göre tedaviye Tazocin eklendi, noradrenalin kademeli azaltıldı.",
        "Meronem 3x1gr (5.gün), Tazocin 4x4.5gr (1.gün)","7.36","42","88","23","-1","95","139",
        "4.0","150-190","+350/24saat","+4","Enteral beslenme (+), defekasyon (+).","EVET","EVET"),
    gun("10.04.2023","Pnömosepsis, klinik düzelme","orta",(3,3,5),
        "Basınç destek modunda (FİO2 %45, PEEP 5, PEEP üstü 11)","120/70","94","37.4",
        "10600","95","1.6","78","1.7","220000",
        "Vazopressör ihtiyacı azaldı, weaning değerlendirildi.",
        "Meronem 3x1gr (6.gün), Tazocin 4x4.5gr (3.gün)","7.37","42","89","23","0","95","138",
        "4.0","145-185","+320/24saat","+3","Enteral beslenme tam doz.","HAYIR","EVET"),
    gun("11.04.2023","Pnömosepsis, klinik düzelme","orta-iyi",(4,3,5),
        "Basınç destek modunda (FİO2 %40, PEEP 5, PEEP üstü 10), weaning başlandı","122/70","90",
        "37.2","10100","70","1.4","64","1.4","260000",
        "Vazopressör ihtiyacı azaldı, noradrenalin çok düşük dozda.",
        "Meronem 3x1gr (7.gün), Tazocin 4x4.5gr (3.gün)","7.38","41","90","24","0","96","138",
        "4.1","140-180","+300/24saat","+3","Enteral beslenme tam doz, tolere etti.","HAYIR","EVET"),
    gun("13.04.2023","Pnömosepsis, iyileşme","iyi",(4,4,6),
        "Ekstübe, nazal oksijen 2 lt/dk ile spontan solunumda","124/72","84","36.8",
        "8100","35","1.2","44","1.2","320000",
        "Vazopressör ihtiyacı kalmadı, noradrenalin kesildi.",
        "Meronem 3x1gr (9.gün), Tazocin 4x4.5gr (5.gün)","7.40","40","94","24","0","97","138",
        "4.1","120-160","+280/24saat","+2","Oral alım açıldı, mobilizasyon başlandı.","HAYIR","HAYIR")],
  "mikro":{"numune":"KAN KÜLTÜRÜ","organizma":"KLEBSIELLA PNEUMONIAE","aciklama":"",
    "antibiyogram":[("Meropenem",("DUYARLI","<=1")),("Piperasilin Tazobaktam",("DUYARLI","<=8")),
      ("Amikasin",("DUYARLI","<=2")),("Siprofloksasin",("DİRENÇLİ",">=4")),
      ("Kolistin",("DUYARLI","<=0.5"))]}},

 {"ad":"ÖRNEK HASTA İKİ","proto":"99002","dosya":"700002","cinsiyet":"KADIN","yas":"71",
  "dtarih":"08.07.1953","taburcu":"Şifa ile taburcu",
  "tani_kesin":"Ürosepsis, septik şok, kronik böbrek hastalığı",
  "apache":"19","mortalite":"33",
  "giris":"Tip 2 diyabet ve kronik böbrek hastalığı öyküsü olan hasta, ateş, titreme ve bilinç "
          "bulanıklığı ile getirildi; ürosepsis ve septik şok tablosuyla yoğun bakıma kabul "
          "edildi. Mesane sondası takıldı, saatlik idrar takibine alındı.",
  "lab":{"WBC (Lökosit)":[22.5,18.0,14.1,9.9,7.9],"HGB (Hemoglobin)":[8.4,8.9,9.5,10.1,10.6],
         "CRP (TÜRBIDIMETRIK)":[280,220,150,80,40],"KREATİNİN (SERUM)":[3.6,3.0,2.4,1.8,1.5],
         "BUN (KAN ÜRE AZOTU)":[62,54,44,34,28],"PLT (Trombosit)":[110,150,190,240,290],
         "CLAC":[3.1,2.4,1.8,1.3,1.1]},
  "gunler":[
    gun("12.05.2023","Ürosepsis ve septik şok","kötü",(3,2,5),
        "Yüksek akımlı oksijen ve aralıklı noninvaziv ventilasyon","85/50","118","38.9",
        "22500","280","3.6","180","3.1","110000",
        "Sıvı resüsitasyonu sonrası noradrenalin ve düşük doz dopamin infüzyonu başlandı.",
        "Meronem 3x1gr (1.gün)","7.28","34","70","17","-7","90","133","5.1","190-240",
        "-200/24saat","+1","Oligürik seyir, hemodiyaliz planlandı.","EVET","EVET"),
    gun("14.05.2023","Ürosepsis, septik şok","kötü-orta",(3,3,5),
        "Aralıklı noninvaziv ventilasyon, oksijen desteği","98/58","110","38.2",
        "18000","220","3.0","150","2.4","150000",
        "Noradrenalin sürüyor, dopamin azaltıldı; bir seans hemodiyaliz uygulandı.",
        "Meronem 3x1gr (3.gün)","7.31","36","78","19","-5","92","135","4.7","170-220",
        "+100/24saat","+3","Parenteral beslenme (+).","EVET","EVET"),
    gun("16.05.2023","Ürosepsis, yanıt alınıyor","orta",(4,4,6),
        "Nazal oksijen, spontan solunum","118/70","92","37.6",
        "14100","150","2.4","120","1.8","190000",
        "Hemodinamik toparlanma ile dopamin kesildi, noradrenalin azaltıldı; ikinci hemodiyaliz yapıldı.",
        "Meronem 3x1gr (5.gün), Kolistin 3x150mg (1.gün)","7.35","38","85","21","-2","94","137",
        "4.3","150-200","+150/24saat","+3","Enteral beslenme başlandı.","EVET","EVET"),
    gun("18.05.2023","Ürosepsis, düzelme","orta-iyi",(4,5,6),
        "Oda havasında, ek oksijen minimal","122/72","86","37.1",
        "9800","90","1.9","95","1.4","240000",
        "Vazopressör ihtiyacı azaldı; renal fonksiyonlar diyalizsiz stabil.",
        "Meronem 3x1gr (7.gün), Kolistin 3x150mg (3.gün)","7.38","40","90","23","-1","96","138",
        "4.2","140-180","+200/24saat","+2","Enteral beslenme tam doz.","HAYIR","EVET"),
    gun("20.05.2023","Ürosepsis, iyileşme","iyi",(4,5,6),
        "Oda havasında satürasyon yeterli","126/74","80","36.7",
        "7900","40","1.5","70","1.1","290000",
        "Vazopressör ihtiyacı yok, hemodinamik stabil.",
        "Meronem 3x1gr (9.gün), Kolistin 3x150mg (5.gün)","7.41","41","96","25","1","98","139",
        "4.0","120-160","+180/24saat","+1","Oral alımı iyi, mobilize.","HAYIR","HAYIR")],
  "mikro":{"numune":"İDRAR KÜLTÜRÜ","organizma":"ESCHERICHIA COLI","aciklama":"",
    "antibiyogram":[("Meropenem",("DUYARLI","<=1")),("Seftriakson",("DİRENÇLİ",">=4")),
      ("Amikasin",("DUYARLI","<=2")),("Kolistin",("DUYARLI","<=0.5")),
      ("Siprofloksasin",("DİRENÇLİ",">=4"))]}},

 {"ad":"ÖRNEK HASTA ÜÇ","proto":"99003","dosya":"700003","cinsiyet":"ERKEK","yas":"58",
  "dtarih":"22.11.1966","taburcu":"Şifa ile taburcu",
  "tani_kesin":"Aspirasyon pnömonisi, solunum yetmezliği",
  "apache":"16","mortalite":"25",
  "giris":"Serebrovasküler olay sekeli ve disfaji öyküsü olan hasta, oral alım sırasında "
          "aspirasyon sonrası solunum sıkıntısı ile yoğun bakıma alındı; entübasyon ihtiyacı "
          "olmadı, yüksek akımlı oksijen ile takip edildi.",
  "lab":{"WBC (Lökosit)":[14.8,11.5,9.2,6.8],"HGB (Hemoglobin)":[11.2,11.5,11.8,12.4],
         "CRP (TÜRBIDIMETRIK)":[95,60,40,12],"KREATİNİN (SERUM)":[1.1,1.0,1.0,0.9],
         "BUN (KAN ÜRE AZOTU)":[28,24,22,18],"PLT (Trombosit)":[210,230,260,300],
         "CLAC":[1.8,1.4,1.2,1.0]},
  "gunler":[
    gun("03.06.2023","Aspirasyon pnömonisi ve solunum yetmezliği","orta",(3,3,5),
        "Yüksek akımlı nazal oksijen (FİO2 %60, 40 lt/dk), entübasyon gerekmedi","132/78","96",
        "38.1","14800","95","1.1","44","1.8","210000",
        "Hemodinamik stabil, vazopressör ihtiyacı yok.",
        "Tazocin 4x4.5gr (1.gün)","7.33","50","72","22","-2","91","140","4.4","130-170",
        "+150/24saat","+2","Nazogastrik sonda ile enteral beslenme, aspirasyon önlemleri.","HAYIR","EVET"),
    gun("05.06.2023","Aspirasyon pnömonisi","orta-iyi",(4,4,5),
        "Yüksek akımlı nazal oksijen (FİO2 %45)","130/76","90","37.6",
        "11500","60","1.0","36","1.4","230000",
        "Hemodinamik stabil.",
        "Tazocin 4x4.5gr (3.gün)","7.36","46","84","23","-1","94","139","4.3","120-160",
        "+120/24saat","+2","Enteral beslenme tolere edildi.","HAYIR","HAYIR"),
    gun("07.06.2023","Aspirasyon pnömonisi, düzelme","iyi",(4,4,6),
        "Nazal kanül 3 lt/dk","128/76","84","37.0",
        "9200","40","1.0","34","1.2","260000",
        "Trakeal aspirat kültüründe üreme olmadı; de-eskalasyon değerlendirildi.",
        "Tazocin 4x4.5gr (5.gün)","7.38","44","90","24","0","96","139","4.2","110-150",
        "+100/24saat","+1","Enteral beslenme tam doz, aspirasyon gözlenmedi.","HAYIR","HAYIR"),
    gun("10.06.2023","Aspirasyon pnömonisi, iyileşme","iyi",(4,5,6),
        "Oda havasında satürasyon %98","124/72","78","36.6",
        "6800","12","0.9","28","1.0","300000",
        "Ek oksijen ihtiyacı yok, hemodinamik stabil.",
        "Tazocin 4x4.5gr (7.gün, tamamlandı)","7.41","40","95","25","1","98","138","4.1",
        "100-140","+120/24saat","+1","Oral alım güvenli, yutma değerlendirmesi planlandı.","HAYIR","HAYIR")],
  "mikro":{"numune":"TRAKEAL ASPİRAT KÜLTÜRÜ","organizma":"",
    "aciklama":"NADİR EPİTEL HÜCRESİ GÖRÜLDÜ. ÜREME OLMADI.","antibiyogram":[]}},
]

if __name__ == "__main__":
    for p in PATIENTS:
        make_pdf(f"{p['proto']}_E1.pdf", e1(p))
        make_pdf(f"{p['proto']}_E2.pdf", e2(p))
        make_pdf(f"{p['proto']}_LAB.pdf", lab(p))
        make_pdf(f"{p['proto']}_M1.pdf", m1(p))
    print(f"\n{len(PATIENTS)} synthetic patients generated. All fictional, safe to publish.")
