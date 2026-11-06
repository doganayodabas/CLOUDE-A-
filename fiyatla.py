#!/usr/bin/env python3
"""
İNDİRİM VE MAĞAZA FİYAT motoru — Excel formüllerinin yerine geçen Python programı.

Kullanım:
    python fiyatla.py GIRDI.xlsx                      # ayarlar.xlsx'i kullanır, CIKTI_<tarih>.xlsx üretir
    python fiyatla.py GIRDI.xlsx -a ayarlar.xlsx -o sonuc.xlsx
    python fiyatla.py --ayar-olustur                  # varsayılan ayarlar.xlsx dosyasını yazar

GIRDI.xlsx içinde şu sekmeler aranır (adı bu ifadeyi İÇERMESİ yeterli):
    "KAMPANYA LİSTESİ"   Nebim kampanya listesi
    "ÜRÜN 1"             spor reyonları, stoğu olan ürünler
    "ÜRÜN 1B"            stoğu biten ama yakında satılan ürünler
    "ÜRÜN 2"             diğer reyonlardan kampanyalı ürünler
    "EK ÜRÜNLER"         sonradan gelen ürünler
    "Depo"               depo dağılım raporu
    "FİYAT —" / "MANUEL ORAN"   (isteğe bağlı) BENİM ORANIM elle oran sütunu

Eski Excel dosyasının kendisi de girdi olarak verilebilir; program yalnız bu ham
sekmeleri okur, formüllü sekmeleri yok sayar.
"""
from __future__ import annotations

import argparse
import datetime as dt
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("openpyxl gerekli:  pip install openpyxl")


# ═══════════════════════════════════════════════════════════════════════════
#  YARDIMCI FONKSİYONLAR
# ═══════════════════════════════════════════════════════════════════════════
def trim(v) -> str:
    """Excel TRIM: baştaki/sondaki boşlukları siler, aradakileri teke indirir."""
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def up(v) -> str:
    """Büyük/küçük harf duyarsız karşılaştırma için normalize eder (Türkçe i/ı dahil)."""
    return trim(v).replace("i", "İ").replace("ı", "I").upper()


def liste(v) -> set[str]:
    """'A, B,C' → {'A','B','C'} (normalize edilmiş)."""
    return {up(x) for x in str(v or "").split(",") if trim(x)}


def sayi(v, varsayilan=0.0) -> float:
    """Excel N(): sayı değilse 0."""
    if v is None or isinstance(v, bool):
        return varsayilan
    if isinstance(v, (int, float)):
        return float(v)
    s = trim(v).replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return varsayilan


def oran(v) -> float:
    """'%40' → 0.40 · 40 → 0.40 · 0.4 → 0.40 · metin → 0."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v) / 100 if v > 1 else float(v)
    s = trim(v)
    if s.startswith("%"):
        return sayi(s[1:]) / 100
    n = sayi(s)
    return n / 100 if n > 1 else n


def yuvarla(x: float, basamak: int = 0) -> float:
    """Excel ROUND (yarım yukarı)."""
    q = Decimal(1).scaleb(-basamak)
    return float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP))


def taban(x: float) -> int:
    """Excel INT (kayan nokta hatasına toleranslı)."""
    return math.floor(x + 1e-9)


def kod(v) -> str:
    """Ürün kodunu metne çevirir (12081832.0 → '12081832')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return trim(v)


def evet(v) -> bool:
    return up(v).startswith("E")


def tarih(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return dt.datetime.strptime(v.strip()[:10], f).date()
            except ValueError:
                pass
    return None


def gecersiz_satir(k: str) -> bool:
    """Başlık / alt bilgi / boş satır mı?"""
    return k == "" or k == "Ürün Kodu" or "odaba" in k.lower()


# ═══════════════════════════════════════════════════════════════════════════
#  AYARLAR
# ═══════════════════════════════════════════════════════════════════════════
# (anahtar, değer, açıklama) — Koşullar sayfasındaki güncel değerlerle aynı.
GENEL_VARSAYILAN = [
    ("KDV Oranı (varsayılan)", 0.10, "Ürün kartında KDV varsa o kullanılır. Yalnız kâr/marj hesabı içindir."),
    ("İthalat Masrafı", 0.08, "Yalnız 'İthalat Markaları' listesindeki markaların maliyetine eklenir."),
    ("İthalat Markaları", "JACK&JONES,JJXX,ONLY", "Virgülle."),
    ("Mağaza Genel Gider Oranı", 0.20, "Kira + personel + enerji. Net marj hesabında kullanılır."),
    ("Min Kârlılık %", 0.05, "Bunun altındaki net marj 'Düşük Marj' olarak işaretlenir."),
    ("Maksimum Toplam İndirim %", 0.50, "Hiçbir ürün etiketin bu oranından fazla indirilmez."),
    ("Yuvarlama Aktif", "Evet", "Evet = fiyatlar …9,99'a yuvarlanır (en yakın)."),
    ("Yuvarlama Basamağı TL", 10, "10 → 779,99 / 789,99 …"),
    ("Kapsam Reyonları", "ÇOCUK REYONU,ERKEK SPOR GİYİM,KADIN SPOR GİYİM",
     "Kurallar yalnız bu reyonlarda çalışır. Diğerleri sadece kampanya oranını alır. Boş = tüm reyonlar."),
    ("Kart Oranını Kullan", "Evet", "Evet = kampanyada olmayan kapsam içi üründe ürün kartındaki 'Mağaza İndirim Oranı' uygulanır."),
    ("Depo Ayrımı Aktif", "Evet", "Mağaza stoğu 0 ve Merkez+Ev depo stoğu > 0 olan ürün DEPO sayılır."),
    ("Son Satış Sağlaması", "Hayır", "Evet = son X gün içinde satılmış ürün depo sayılmaz."),
    ("Depo İçin Son Satıştan Gün", 10, ""),
    ("Stoksuz Ürüne İndirim Yok", "Evet", "Envanteri 0 olan ürün etiket fiyatında kalır, listelere girmez."),
    ("Stoksuz İstisna Günü", 14, "Son satışı bu kadar gün içindeyse stoksuz sayılmaz (iade ihtimali). 0 = istisna yok."),
    ("M2 Aktarım Kapsamı", "Sadece indirimli", "Sadece indirimli / Kapsam içi tümü / Tüm ürünler"),
    ("Çoklu Kampanyada Tek İndirim Yok", "Evet", "Tipi ÇOKLU olan kampanyadaki ürün başka indirim almaz."),
    ("Çoklu Kuralı Hariç Reyonlar", "ÇOCUK REYONU", "Virgülle."),
    ("Depo Tavan Oranı", 0.25, "Depo ürünlerinde oran bunu aşamaz. Boş/0 = tavan yok."),
    ("Tavan Hariç Markalar", "GAP", "Virgülle."),
    ("Tavan Hariç Reyonlar", "ÇOCUK REYONU", "Virgülle."),
    ("Tavan Hariç Sezonlar", "27K,27Y", "Virgülle."),
    ("Depo Raporu — Mağaza Sütunu", "Magaza Depo", "Depo raporundaki başlık adı."),
    ("Depo Raporu — Merkez Sütunu", "Merkez Depo", ""),
    ("Depo Raporu — Ev Depo Sütunu", "Ev Depo Yalı_Mimoza", ""),
    ("Depo Raporu — Ürün Kodu Sütun No", 9, "Raporda 'Ürün Kodu' başlığı yoksa kaçıncı sütunda olduğu."),
]

KURAL_BASLIK = ["Kural Adı", "Kapsam (Depo/Mağaza/Outlet/Mimoza/Noos/boş)", "Reyon", "Ürün Alt Grubu",
                "Hariç Alt Grup", "Marka", "Sadece Sezonlar", "Hariç Sezonlar", "İndirim %",
                "Uygulama (Listeyi de ez / Yalnız listede yoksa / Daha yüksekse)", "Aktif"]
KURAL_VARSAYILAN = [
    ["Parfüm — her koşulda indirimsiz", "", "", "PARFÜM", "", "", "", "", 0, "Listeyi de ez", "Evet"],
    ["Noos Bestseller — %15 (JACK&JONES, JJXX)", "Noos", "", "", "", "JACK&JONES,JJXX", "", "", 0.15, "Listeyi de ez", "Evet"],
    ["GAP — her koşulda %40 (mağaza + depo)", "", "", "", "", "GAP", "", "", 0.40, "Listeyi de ez", "Evet"],
    ["JACK&JONES çorap + bere — %15", "", "", "ÇORAP,BERE", "", "JACK&JONES,JJXX", "", "", 0.15, "Listeyi de ez", "Evet"],
    ["Ayakkabı — her koşulda %30", "", "", "AYAKKABI,BOT,TERLİK", "", "", "", "", 0.30, "Listeyi de ez", "Evet"],
    ["Yeni sezon MAVİ (indirimsiz)", "", "", "", "", "MAVİ", "27K,27Y", "", 0, "Listeyi de ez", "Evet"],
    ["Yeni sezon JACK&JONES", "", "", "", "", "JACK&JONES", "27K,27Y", "", 0.15, "Listeyi de ez", "Evet"],
    ["Yeni sezon JJXX", "", "", "", "", "JJXX", "27K,27Y", "", 0.15, "Listeyi de ez", "Evet"],
    ["MAVİ 26Y — tanımlıysa uygula yoksa yok", "", "", "", "", "MAVİ", "26Y", "", 0, "Yalnız listede yoksa", "Evet"],
    ["Çocuk jean pantolon", "", "ÇOCUK REYONU", "JEAN PANTOLON", "", "", "", "27K,27Y", 0.40, "Daha yüksekse", "Evet"],
    ["MAVİ jean — sadece tanımlı indirim", "", "", "JEAN PANTOLON", "", "MAVİ", "", "", 0, "Yalnız listede yoksa", "Evet"],
    ["Eski sezon jean — tanımsızsa indirim yok", "", "", "JEAN PANTOLON", "", "", "", "26Y,27K,27Y", 0, "Yalnız listede yoksa", "Evet"],
    ["Depo — varsayılan oran", "Depo", "", "", "", "", "", "27K,27Y,26Y,NOOS", 0.25, "Listeyi de ez", "Evet"],
]

KAMPANYA_BASLIK = ["ItemListDescription", "Olması Gereken Oran", "Kampanya Tipi (TEKLİ/TEKFİYAT/BİTTİ/ÇOKLU/ÇİFTLİ)", "Tek Fiyat ₺"]
KAMPANYA_VARSAYILAN = [
    ["%15 İndirim", "%15", "TEKLİ", None],
    ["%20 İndirim", "%20", "TEKLİ", None],
    ["%25 İndirim", "%25", "TEKLİ", None],
    ["%30 İndirim", "%30", "TEKLİ", None],
    ["%40 İndirim", "%40", "TEKLİ", None],
    ["%50 İndirim", "%50", "TEKLİ", None],
    ["2 Al 1 Öde", "2al 1 öde", "TEKLİ", None],
    ["2.Ürüne %50 İndirim çiftli kampanya", "BİTTİ", "BİTTİ", None],
    ["3 Adet Alana %20 İndirim", "BİTTİ", "BİTTİ", None],
    ["Tek Fiyat 299.99₺", "—", "TEKFİYAT", 299.99],
    ["Tek Fiyat 399.99₺", "—", "TEKFİYAT", 399.99],
    ["Tek Fiyat 499.99₺", "—", "TEKFİYAT", 499.99],
]

NEBIM_ORANLARI = [50, 40, 30, 25, 20, 15]


@dataclass
class Kural:
    no: int
    ad: str
    kapsam: str
    reyon: set
    alt_grup: set
    haric_alt: set
    marka: set
    sezon: set
    haric_sezon: set
    oran: float
    uygulama: str  # "L" listeyi de ez · "Y" yalnız listede yoksa · "D" daha yüksekse


@dataclass
class KampanyaTipi:
    oran_metni: str
    tip: str
    tek_fiyat: float | None


@dataclass
class Ayarlar:
    g: dict
    kurallar: list[Kural]
    kampanyalar: dict[str, KampanyaTipi]

    def __getitem__(self, k):
        return self.g.get(k)


def ayar_dosyasi_yaz(yol: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Genel"
    ws.append(["Ayar", "Değer", "Açıklama"])
    for r in GENEL_VARSAYILAN:
        ws.append(list(r))
    ws2 = wb.create_sheet("Kurallar")
    ws2.append(KURAL_BASLIK)
    for r in KURAL_VARSAYILAN:
        ws2.append(r)
    ws2.append([])
    ws2.append(["ⓘ Ürüne uyan İLK aktif kural geçerlidir (yukarıdan aşağı). Satır ekleyip silebilirsiniz; sınır yok."])
    ws3 = wb.create_sheet("Kampanyalar")
    ws3.append(KAMPANYA_BASLIK)
    for r in KAMPANYA_VARSAYILAN:
        ws3.append(r)
    ws3.append([])
    ws3.append(["ⓘ Nebim'de yeni bir kampanya adı çıkarsa buraya ekleyin. TEKFİYAT ise 'Tek Fiyat ₺' sütununu doldurun."])
    for s in wb.worksheets:
        _baslik_bicimle(s, 1)
        for i, w in enumerate([42, 22, 22, 22, 18, 22, 16, 18, 10, 26, 8], 1):
            s.column_dimensions[get_column_letter(i)].width = w
    wb["Genel"].column_dimensions["C"].width = 90
    for row in wb["Genel"].iter_rows(min_row=2):
        row[1].fill = SARI
    wb.save(yol)


def ayar_oku(yol: Path) -> Ayarlar:
    g = {k: v for k, v, _ in GENEL_VARSAYILAN}
    kurallar_ham = KURAL_VARSAYILAN
    kamp_ham = KAMPANYA_VARSAYILAN
    if yol and yol.exists():
        wb = openpyxl.load_workbook(yol, data_only=True)
        if "Genel" in wb.sheetnames:
            for r in wb["Genel"].iter_rows(min_row=2, values_only=True):
                if r and r[0] and trim(r[0]) in g:
                    g[trim(r[0])] = r[1]
        if "Kurallar" in wb.sheetnames:
            kurallar_ham = [list(r) + [None] * 11 for r in wb["Kurallar"].iter_rows(min_row=2, values_only=True)
                            if r and r[0] and not str(r[0]).startswith("ⓘ")]
        if "Kampanyalar" in wb.sheetnames:
            kamp_ham = [list(r) + [None] * 4 for r in wb["Kampanyalar"].iter_rows(min_row=2, values_only=True)
                        if r and r[0] and not str(r[0]).startswith("ⓘ")]
    kurallar = []
    for i, r in enumerate(kurallar_ham, 1):
        if not evet(r[10]):
            continue
        kurallar.append(Kural(i, trim(r[0]), up(r[1])[:2], liste(r[2]), liste(r[3]), liste(r[4]), liste(r[5]),
                              liste(r[6]), liste(r[7]), oran(r[8]), (up(r[9]) or "L")[0]))
    kamp = {trim(r[0]): KampanyaTipi(trim(r[1]), up(r[2]), sayi(r[3]) or None) for r in kamp_ham}
    return Ayarlar(g, kurallar, kamp)


# ═══════════════════════════════════════════════════════════════════════════
#  GİRDİ OKUMA
# ═══════════════════════════════════════════════════════════════════════════
def sekme_bul(wb, *ifadeler, haric=()):
    for ad in wb.sheetnames:
        u = up(ad)
        if all(up(i) in u for i in ifadeler) and not any(up(h) in u for h in haric):
            return wb[ad]
    return None


def tablo_oku(ws, baslik_kod="Ürün Kodu"):
    """Başlık satırını bulup her satırı {başlık: değer} sözlüğü olarak döndürür."""
    if ws is None:
        return []
    satirlar = list(ws.iter_rows(values_only=True))
    bas_i = next((i for i, r in enumerate(satirlar[:15]) if any(trim(c) == baslik_kod for c in r)), None)
    if bas_i is None:
        return []
    basliklar = [trim(c) for c in satirlar[bas_i]]
    cikti = []
    for r in satirlar[bas_i + 1:]:
        d = {b: r[j] for j, b in enumerate(basliklar) if b and j < len(r)}
        if gecersiz_satir(kod(d.get(baslik_kod))):
            continue
        cikti.append(d)
    return cikti


@dataclass
class Urun:
    kod: str
    ad: str
    marka: str
    reyon: str
    alt_grup: str
    sezon: str
    envanter: float
    alis: float | None
    etiket: float
    son_giris: object
    son_satis: object
    kdv: float
    kart_oran_ham: object
    isaret: str
    yeni_eski: str
    kaynak_sekme: str
    # hesaplananlar
    h: dict = field(default_factory=dict)


def urunleri_oku(wb) -> list[Urun]:
    sekmeler = [
        ("ÜRÜN 1", sekme_bul(wb, "ÜRÜN 1", haric=("1B",))),
        ("ÜRÜN 1B", sekme_bul(wb, "ÜRÜN 1B")),
        ("ÜRÜN 2", sekme_bul(wb, "ÜRÜN 2")),
        ("EK ÜRÜNLER", sekme_bul(wb, "EK ÜRÜN")),
    ]
    goruldu, urunler = set(), []
    for ad, ws in sekmeler:
        for d in tablo_oku(ws):
            k = kod(d.get("Ürün Kodu"))
            if k in goruldu:
                continue
            goruldu.add(k)
            kdv = sayi(d.get("KDV"))
            urunler.append(Urun(
                kod=k, ad=trim(d.get("Ürün Adı")), marka=trim(d.get("Marka")), reyon=trim(d.get("Reyon")),
                alt_grup=trim(d.get("Ürün Alt Grubu")), sezon=trim(d.get("Yıl Sezon")),
                envanter=sayi(d.get("Envanter")),
                alis=None if d.get("Satın Alma Fiyatı") in (None, "") else sayi(d.get("Satın Alma Fiyatı")),
                etiket=sayi(d.get("Perakende Satış Fiyatı")),
                son_giris=tarih(d.get("Son Giriş Tarihi")), son_satis=tarih(d.get("Son Satış Tarihi")),
                kdv=kdv / 100 if kdv > 1 else kdv,
                kart_oran_ham=d.get("Mağaza İndirim Oranı"),
                isaret=trim(d.get("Noos_Outlet Ürünler") or d.get("Devam_Noos Ürünler")),
                yeni_eski=trim(d.get("Yeni/Eski Ürün")), kaynak_sekme=ad))
    return urunler


def kampanya_oku(wb):
    ws = sekme_bul(wb, "KAMPANYA LİSTESİ")
    satirlar = []
    for i, d in enumerate(tablo_oku(ws), 1):
        satirlar.append({
            "satir": i, "kod": kod(d.get("Ürün Kodu")), "ad": trim(d.get("Ürün Adı")),
            "marka": trim(d.get("Marka")), "reyon": trim(d.get("Reyon")),
            "kart_oran": d.get("Mağaza İndirim Oranı"),
            "list_code": trim(d.get("ItemListCode")), "list_desc": trim(d.get("ItemListDescription")),
            "psf": sayi(d.get("Perakende Satış Fiyatı")),
            "m2": sayi(next((v for k, v in d.items() if k.startswith("Fiyat Listesi M2")), 0)),
        })
    return satirlar


def depo_oku(wb, ay: Ayarlar):
    """{kod: [mağaza, merkez, ev]} — aynı kod birden çok satırdaysa toplanır."""
    ws = sekme_bul(wb, "Depo", haric=("ORAN",))
    if ws is None:
        return None
    satirlar = list(ws.iter_rows(values_only=True))
    hedef = [trim(ay["Depo Raporu — Mağaza Sütunu"]), trim(ay["Depo Raporu — Merkez Sütunu"]),
             trim(ay["Depo Raporu — Ev Depo Sütunu"])]
    sut = [None, None, None]
    kod_sut = int(sayi(ay["Depo Raporu — Ürün Kodu Sütun No"], 9)) - 1
    for r in satirlar[:10]:
        for j, c in enumerate(r[:25]):
            t = trim(c)
            if t == "Ürün Kodu":
                kod_sut = j
            for n, h in enumerate(hedef):
                if t == h:
                    sut[n] = j
    stok = defaultdict(lambda: [0.0, 0.0, 0.0])
    for r in satirlar:
        if kod_sut >= len(r):
            continue
        k = kod(r[kod_sut])
        if gecersiz_satir(k):
            continue
        # başlık satırları (depo adları) kod sütununda boş olduğu için zaten elenir
        s = stok[k]
        for n, j in enumerate(sut):
            if j is not None and j < len(r):
                s[n] += sayi(r[j])
    return dict(stok)


def manuel_oranlari_oku(wb) -> dict[str, float]:
    """FİYAT — … sekmelerindeki BENİM ORANIM sütunu (veya MANUEL ORANLAR sekmesi)."""
    m = {}
    for ad in wb.sheetnames:
        if not (up(ad).startswith("FİYAT") or "MANUEL" in up(ad)):
            continue
        ws = wb[ad]
        satirlar = list(ws.iter_rows(values_only=True))
        bas = next((i for i, r in enumerate(satirlar[:10]) if any(trim(c) == "Ürün Kodu" for c in r)), None)
        if bas is None:
            continue
        b = [trim(c) for c in satirlar[bas]]
        ko = b.index("Ürün Kodu")
        oc = next((j for j, x in enumerate(b) if up(x) in ("BENİM ORANIM", "ORAN", "MANUEL ORAN")), None)
        if oc is None:
            continue
        for r in satirlar[bas + 1:]:
            k = kod(r[ko]) if ko < len(r) else ""
            v = r[oc] if oc < len(r) else None
            if k and not k.startswith("(") and v not in (None, "") and not str(v).startswith("="):
                m[k] = oran(v)
    return m


# ═══════════════════════════════════════════════════════════════════════════
#  KAMPANYA SATIRLARI
# ═══════════════════════════════════════════════════════════════════════════
def kampanya_isle(satirlar, urun_by_kod, ay: Ayarlar):
    """Her satıra oran/tip/durum ekler, her ürün için 'birincil' kampanya satırını seçer."""
    kod_satirlari = defaultdict(list)
    sayac = defaultdict(int)
    for s in satirlar:
        kod_satirlari[s["kod"]].append(s)
        sayac[(s["kod"], s["list_code"])] += 1
    for s in satirlar:
        kt = ay.kampanyalar.get(s["list_desc"])
        s["olmasi_gereken"] = kt.oran_metni if kt else ""
        s["tip"] = kt.tip if kt else ""
        s["tek_fiyat"] = kt.tek_fiyat if kt and kt.tip == "TEKFİYAT" else None
        s["oran_sayi"] = sayi(s["olmasi_gereken"].replace("%", "")) if kt else 0
        u = urun_by_kod.get(s["kod"])
        kart = s["kart_oran"] if trim(s["kart_oran"]) else (u.kart_oran_ham if u else None)
        s["kart_metin"] = trim(kart)
        s["kart_sayi"] = sayi(trim(kart).replace("%", ""))
        s["listede"] = u is not None
        s["mukerrer"] = sayac[(s["kod"], s["list_code"])]
        s["fiyattan_ind"] = yuvarla((1 - s["m2"] / s["psf"]) * 100) if s["psf"] > 0 and s["m2"] > 0 else None
        # DURUM
        if not kt:
            durum = "TANIMSIZ KAMPANYA"
        elif kt.tip == "BİTTİ":
            durum = "KAMPANYA BİTTİ"
        elif kt.tip == "TEKFİYAT":
            durum = "UYGUN"
        elif s["kart_sayi"] == 0 and s["kart_metin"] and kt.tip in ("ÇOKLU", "ÇİFTLİ"):
            durum = "UYGUN"
        elif not s["kart_metin"]:
            durum = "HATA - ORAN BOŞ"
        elif s["kart_metin"].lower() == kt.oran_metni.lower():
            durum = "UYGUN"
        else:
            durum = "HATA - ORAN UYUŞMUYOR"
        s["durum"] = durum
        s["fiyat_uyum"] = ("FİYAT İLE ORAN UYUŞMUYOR" if s["fiyattan_ind"] and s["oran_sayi"]
                           and abs(s["fiyattan_ind"] - s["oran_sayi"]) > 2 else "")

    birincil = {}
    for k, ss in kod_satirlari.items():
        sec = (next((s for s in ss if s["tip"] == "ÇOKLU"), None)
               or next((s for s in ss if s["list_code"] == "%+2.%50"), None)
               or next((s for s in ss if s["oran_sayi"] > 0 and s["oran_sayi"] == s["kart_sayi"]), None)
               or next((s for s in ss if s["oran_sayi"] > 0), None)
               or ss[0])
        birincil[k] = sec
        diger = [s for s in ss if s["list_code"] not in ("%30_", "BİTTİ")]
        for s in ss:
            s["birincil"] = s is sec
            s["cakisma"] = ""
            if len({x["list_code"] for x in diger}) > 1 and s in diger:
                s["cakisma"] = "BAŞKA KAMPANYADA DA VAR"
            s["bulundugu"] = "   +   ".join(f'{x["list_desc"]} ({x["olmasi_gereken"] or "?"})' for x in ss)
            s["kac_liste"] = len(ss)
            s["sira_metni"] = f"{ss.index(s) + 1} / {len(ss)}"
    return kod_satirlari, birincil


# ═══════════════════════════════════════════════════════════════════════════
#  FİYAT MOTORU
# ═══════════════════════════════════════════════════════════════════════════
def kural_uyar(k: Kural, u: Urun, h: dict) -> bool:
    if k.kapsam:
        bayrak = {"DE": h["depo"], "MA": not h["depo"], "OU": h["outlet"], "MI": h["mimoza"], "NO": h["noos"]}
        if not bayrak.get(k.kapsam, False):
            return False
    if k.reyon and up(u.reyon) not in k.reyon:
        return False
    if k.alt_grup and up(u.alt_grup) not in k.alt_grup:
        return False
    if k.haric_alt and up(u.alt_grup) in k.haric_alt:
        return False
    if k.marka and up(u.marka) not in k.marka:
        return False
    if k.sezon and up(h["etkin_sezon"]) not in k.sezon:
        return False
    if k.haric_sezon and up(h["etkin_sezon"]) in k.haric_sezon:
        return False
    return True


def fiyat_99(ham: float, alt_sinir: float, adim: float) -> float:
    """En yakın …9,99'a yuvarlar; alt sınırın altına düşürmez."""
    c = yuvarla((ham + 0.01) / adim) * adim - 0.01
    if c >= alt_sinir - 1e-6:
        return round(c, 2)
    f = taban((alt_sinir + 0.01) / adim) * adim - 0.01
    return round(f if f >= alt_sinir - 1e-6 else f + adim, 2)


def hesapla(urunler, kamp_satirlari, birincil, depo, manuel, ay: Ayarlar, bugun: dt.date):
    kapsam = liste(ay["Kapsam Reyonları"])
    ithalat_m = liste(ay["İthalat Markaları"])
    kamp_kodlari = {s["kod"] for s in kamp_satirlari}
    adim = max(1.0, sayi(ay["Yuvarlama Basamağı TL"], 10))
    yuv = evet(ay["Yuvarlama Aktif"])
    maks = sayi(ay["Maksimum Toplam İndirim %"], 0.5)
    tavan = sayi(ay["Depo Tavan Oranı"])
    t_marka, t_reyon, t_sezon = (liste(ay["Tavan Hariç Markalar"]), liste(ay["Tavan Hariç Reyonlar"]),
                                 liste(ay["Tavan Hariç Sezonlar"]))
    coklu_haric = liste(ay["Çoklu Kuralı Hariç Reyonlar"])
    gider = sayi(ay["Mağaza Genel Gider Oranı"])
    min_kar = sayi(ay["Min Kârlılık %"])
    kdv_vars = sayi(ay["KDV Oranı (varsayılan)"], 0.1)
    ithalat = sayi(ay["İthalat Masrafı"])
    stoksuz_gun = sayi(ay["Stoksuz İstisna Günü"])
    depo_gun = sayi(ay["Depo İçin Son Satıştan Gün"])
    m2_kapsam = up(ay["M2 Aktarım Kapsamı"])[:1]  # S / K / T

    kod_kamp = defaultdict(list)
    for s in kamp_satirlari:
        kod_kamp[s["kod"]].append(s)

    for u in urunler:
        h = u.h
        S = u.etiket
        ks = kod_kamp.get(u.kod, [])
        b = birincil.get(u.kod)

        # ── Ürün özellikleri ──────────────────────────────────────────────
        h["kapsam_ici"] = (not kapsam) or up(u.reyon) in kapsam
        h["kapsam_disi"] = not h["kapsam_ici"] and u.kod not in kamp_kodlari
        isr = u.isaret.lower()
        h["outlet"] = "outlet" in isr
        h["noos"] = not h["outlet"] and ("noos" in isr or "devam" in isr)
        h["etkin_sezon"] = "NOOS" if h["noos"] else u.sezon
        h["kart_oran"] = oran(u.kart_oran_ham) if trim(u.kart_oran_ham).startswith("%") or isinstance(
            u.kart_oran_ham, (int, float)) else 0.0
        h["kart_notu"] = trim(u.kart_oran_ham) if trim(u.kart_oran_ham) and h["kart_oran"] == 0 else ""
        mg, mz, ev = depo.get(u.kod, [0, 0, 0]) if depo else (0, 0, 0)
        h["depo_raporunda"] = bool(depo) and u.kod in depo
        h["stok_magaza"], h["stok_merkez"], h["stok_ev"] = mg, mz, ev
        h["depo_stogu"] = mz + ev
        h["mimoza"] = ev > 0
        h["son_satistan_gun"] = (bugun - u.son_satis).days if u.son_satis else None
        yakin = (evet(ay["Son Satış Sağlaması"]) and h["son_satistan_gun"] is not None
                 and h["son_satistan_gun"] <= depo_gun)
        h["depo"] = (h["kapsam_ici"] and evet(ay["Depo Ayrımı Aktif"]) and mg == 0 and h["depo_stogu"] > 0
                     and not yakin)
        h["depo_notu"] = (f"Son {h['son_satistan_gun']} gün içinde satılmış, depo sayılmadı — depoda "
                          f"{h['depo_stogu']:g} adet var") if mg == 0 and h["depo_stogu"] > 0 and yakin else ""
        iade_istisna = stoksuz_gun > 0 and u.son_satis and (bugun - u.son_satis).days <= stoksuz_gun
        h["stoksuz"] = (evet(ay["Stoksuz Ürüne İndirim Yok"])
                        and (u.envanter <= 0 or (h["depo_raporunda"] and mg + mz + ev <= 0))
                        and not iade_istisna)
        h["coklu"] = (evet(ay["Çoklu Kampanyada Tek İndirim Yok"]) and any(s["tip"] == "ÇOKLU" for s in ks)
                      and up(u.reyon) not in coklu_haric)
        h["tek_fiyat"] = next((s["tek_fiyat"] for s in ks if s["tek_fiyat"]), None)
        h["kampanyada"] = b is not None
        h["kampanya_orani"] = b["oran_sayi"] / 100 if b else 0.0
        h["kampanya_adi"] = b["list_desc"] if b else ""
        h["manuel"] = manuel.get(u.kod)

        # ── Temel oran: kampanya > kart oranı > 0 ───────────────────────
        if h["kapsam_disi"]:
            temel, temel_kaynak = 0.0, "Kapsam dışı reyon"
        elif not h["kapsam_ici"]:
            temel, temel_kaynak = h["kampanya_orani"], f"Kampanya [{h['kampanya_adi']}]"
        elif h["kampanyada"]:
            temel, temel_kaynak = h["kampanya_orani"], f"Kampanya [{h['kampanya_adi']}]"
        elif evet(ay["Kart Oranını Kullan"]) and h["kart_oran"] > 0:
            temel, temel_kaynak = h["kart_oran"], "Kart oranı"
        else:
            temel, temel_kaynak = 0.0, "İndirim yok"

        # ── Otomatik kurallar (ilk uyan) ─────────────────────────────────
        kural = next((k for k in ay.kurallar if kural_uyar(k, u, h)), None) if h["kapsam_ici"] else None
        h["kural"] = kural.ad if kural else ""
        if kural is None:
            oran_k, kaynak = temel, temel_kaynak
        elif kural.uygulama == "Y":  # yalnız listede yoksa
            if h["kart_oran"] == 0 and not h["kampanyada"]:
                oran_k, kaynak = kural.oran, f"Kural [{kural.ad}]"
            else:
                oran_k, kaynak = temel, f"{temel_kaynak} · kural: {kural.ad}"
        elif kural.uygulama == "D":  # daha yüksekse
            oran_k = max(temel, kural.oran)
            kaynak = f"Kural [{kural.ad}]" if kural.oran > temel else f"{temel_kaynak} · kural: {kural.ad}"
        else:  # listeyi de ez
            oran_k, kaynak = kural.oran, f"Kural [{kural.ad}]"

        tavan_uygun = (h["depo"] and tavan > 0 and up(u.marka) not in t_marka and up(u.reyon) not in t_reyon
                       and up(u.sezon) not in t_sezon)
        h["sistem_onerisi"] = (0.0 if (h["coklu"] or h["stoksuz"])
                               else tavan if tavan_uygun and oran_k > tavan else oran_k)

        # ── Uygulanan oran (öncelik sırası) ──────────────────────────────
        h["tavan_uygulandi"] = False
        if h["manuel"] is not None:
            U, kaynak = h["manuel"], "BENİM ORANIM (elle)"
        elif h["tek_fiyat"] and S > 0:
            U, kaynak = max(0.0, 1 - h["tek_fiyat"] / S), f"Tek Fiyat Kampanyası [{h['tek_fiyat']:.2f} ₺]"
        elif h["coklu"]:
            U, kaynak = 0.0, "Çoklu kampanya — tek ürün indirimi yok"
        elif h["stoksuz"]:
            U, kaynak = 0.0, "Stoksuz — indirim yok"
        elif tavan_uygun and oran_k > tavan:
            U, kaynak = tavan, f"{kaynak} · depo tavanı %{tavan * 100:.0f}"
            h["tavan_uygulandi"] = True
        else:
            U = oran_k
        if h["depo"] and "Depo" not in kaynak:
            kaynak += " + Depo"
        if h["kart_notu"]:
            kaynak += f" ⚑{h['kart_notu']}"
        h["oran"] = U
        h["kaynak"] = kaynak

        # ── Mağaza fiyatı ────────────────────────────────────────────────
        alt = 0.0
        if S > 0:
            alt = (taban((S * (1 - maks) + 0.01) / adim) * adim - 0.01) if yuv else S * (1 - maks)
        ham = max(S * (1 - U), alt)
        h["maks_indirim_siniri"] = S * (1 - U) < alt - 0.005
        if h["manuel"] is None and h["tek_fiyat"]:
            fiyat = h["tek_fiyat"]
        elif yuvarla(U * 100) == 0 or S <= 0:
            fiyat = S
        else:
            gl = fiyat_99(ham, alt, adim) if yuv else round(ham, 2)
            fiyat = max(0.0, yuvarla(S * (1 - U) / adim) * adim - 0.01) if gl >= S - 0.005 else gl
        h["fiyat"] = round(fiyat, 2)
        h["fiili_indirim"] = max(0.0, 1 - fiyat / S) if S else 0.0

        # ── Maliyet / kârlılık (bilgi amaçlı, fiyatı değiştirmez) ────────
        if u.alis is not None:
            m = u.alis * (1 + (u.kdv if u.kdv > 0 else kdv_vars))
            if up(u.marka) in ithalat_m:
                m *= 1 + ithalat
            h["maliyet"] = m
            h["basabas"] = taban(m / max(0.05, 1 - gider)) + 0.99
            h["brut_kar"] = fiyat - m
            h["net_kar"] = fiyat - m - fiyat * gider
            h["net_marj"] = h["net_kar"] / fiyat if fiyat else None
            h["durum"] = ("⚠ ZARARDA" if h["net_kar"] < 0 else
                          "Düşük Marj" if (h["net_marj"] or 0) < min_kar else "Kârlı")
        else:
            h.update(maliyet=None, basabas=None, brut_kar=None, net_kar=None, net_marj=None, durum="")

        # ── Nebim'deki mevcut M2 ─────────────────────────────────────────
        m2 = max((s["m2"] for s in ks), default=0)
        h["nebim_m2"] = m2 if m2 > 0 else S
        h["degisiyor"] = abs(h["fiyat"] - h["nebim_m2"]) > 0.02

        # ── Hangi listelere girer ────────────────────────────────────────
        yuzde = yuvarla(U * 100)
        h["nebim_listesi"] = int(yuzde) if yuzde in NEBIM_ORANLARI and (
            h["manuel"] is not None or not h["tek_fiyat"]) else None
        h["m2_aktarim"] = yuzde > 0 and (m2_kapsam == "T" or (m2_kapsam == "K" and h["kapsam_ici"])
                                         or (S > 0 and fiyat < S - 0.004))

    # kampanya satırlarına öneri
    by = {u.kod: u for u in urunler}
    hatali_kodlar = {s["kod"] for s in kamp_satirlari if s["durum"].startswith("HATA")}
    for s in kamp_satirlari:
        u = by.get(s["kod"])
        if s["durum"] == "KAMPANYA BİTTİ":
            s["oneri"] = "SİL — kampanya bitti, Nebim listesinden çıkarın"
        elif u is None:
            s["oneri"] = "TEYİT ET — ürün listesinde yok"
        elif not s["oran_sayi"] and s["tip"] != "TEKFİYAT":
            s["oneri"] = "TEYİT ET — kampanya oranı tanımsız"
        elif s["tip"] == "TEKFİYAT":
            s["oneri"] = "KALSIN — tek fiyat" if u.h["tek_fiyat"] == s["tek_fiyat"] else "TEYİT ET"
        else:
            uy = yuvarla(u.h["oran"] * 100)
            s["oneri"] = (f"KALSIN — bu üründe %{uy:.0f} uygulanıyor" if abs(uy - s["oran_sayi"]) <= 1 else
                          f"SİL — bu üründe %{uy:.0f} uygulanıyor, bu satır %{s['oran_sayi']:.0f} listesi")
        s["hata_listesine"] = (s["durum"].startswith("HATA") or s["durum"] in ("TANIMSIZ KAMPANYA", "KAMPANYA BİTTİ")
                               or not s["listede"] or s["mukerrer"] > 1 or s["kod"] in hatali_kodlar)


# ═══════════════════════════════════════════════════════════════════════════
#  ÇIKTI
# ═══════════════════════════════════════════════════════════════════════════
LACIVERT = PatternFill("solid", fgColor="1F3864")
SARI = PatternFill("solid", fgColor="FFF2CC")
TURUNCU = PatternFill("solid", fgColor="FCE4D6")
KIRMIZI = PatternFill("solid", fgColor="F8CBAD")
YESIL = PatternFill("solid", fgColor="E2EFDA")


def _baslik_bicimle(ws, satir):
    for c in ws[satir]:
        if c.value is not None:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = LACIVERT
            c.alignment = Alignment(wrap_text=True, vertical="center")


def sekme(wb, ad, baslik, satirlar, genislik=None, para=(), yuzde=(), not_metni=None):
    ws = wb.create_sheet(ad[:31])
    bas = 1
    if not_metni:
        ws.append([not_metni])
        ws["A1"].font = Font(italic=True, color="595959")
        bas = 2
    ws.append(baslik)
    _baslik_bicimle(ws, bas)
    for r in satirlar:
        ws.append(r)
    ws.freeze_panes = ws.cell(bas + 1, 1)
    if satirlar:
        ws.auto_filter.ref = f"A{bas}:{get_column_letter(len(baslik))}{bas + len(satirlar)}"
    for i, b in enumerate(baslik, 1):
        L = get_column_letter(i)
        ws.column_dimensions[L].width = (genislik or {}).get(b, max(10, min(40, len(str(b)) + 2)))
        fmt = "#,##0.00" if b in para else "0%" if b in yuzde else None
        if fmt:
            for row in ws.iter_rows(min_row=bas + 1, min_col=i, max_col=i):
                row[0].number_format = fmt
    return ws


def cikti_yaz(yol, urunler, kamp_satirlari, ay, bugun, manuel):
    wb = openpyxl.Workbook()
    ozet = wb.active
    ozet.title = "ÖZET"

    # ── Nebim Aktarım (M2 fiyat listesi) ─────────────────────────────────
    nebim_bas = ["Barcode", "ItemTypeCode", "ItemCode", "ColorCode", "ItemDim1Code", "ItemDim2Code",
                 "ItemDim3Code", "Qty1", "Qty2", "LineDescription", "PriceCurrencyCode", "PriceExchangeRate",
                 "Price", "PriceVI", "DocCurrencyCode", "LDisRate1", "LDisRate2", "LDisRate3", "LDisRate4",
                 "LDisRate5", "UnitOfMeasureCode", "PaymentPlanCode", "SalespersonCode", "CostCenterCode",
                 *[f"ITAtt{i:02d}" for i in range(1, 11)], "LineSumID", "LotCode", "ImportFileNumber",
                 "ExportFileNumber", "ProductSerialNumber", "PurchasePlanCode", "BatchCode", "GLTypeCode",
                 "ManufactureDate", "ExpiryDate"]
    m2 = [u for u in urunler if u.h["m2_aktarim"]]
    rows = []
    for u in m2:
        r = [None] * len(nebim_bas)
        r[1], r[2], r[7], r[8], r[12], r[13], r[14] = 1, u.kod, 1, 1, 0, u.h["fiyat"], "TRY"
        r[15], r[20], r[42], r[43] = int(yuvarla(u.h["oran"] * 100)), "AD", 0, 0
        rows.append(r)
    sekme(wb, "Nebim Aktarım", nebim_bas, rows, para=("PriceVI",))

    # ── NEBİM %xx listeleri ──────────────────────────────────────────────
    lb = ["ItemCode", "ItemDescription", "ColorCode", "ColorDescription"]
    adetler = {}
    for y in NEBIM_ORANLARI:
        sec = [u for u in urunler if u.h["nebim_listesi"] == y]
        adetler[f"NEBİM %{y}"] = len(sec)
        sekme(wb, f"NEBİM %{y}", lb, [[u.kod, u.ad, None, None] for u in sec],
              genislik={"ItemCode": 18, "ItemDescription": 45})
    urun_kodlari = {u.kod for u in urunler}
    for desc, kt in ay.kampanyalar.items():
        if kt.tip != "TEKFİYAT":
            continue
        sec = [s for s in kamp_satirlari if s["list_desc"] == desc and s["kod"] in urun_kodlari]
        ad = f"NEBİM TEK FİYAT {kt.tek_fiyat:g}".replace(".99", "")
        adetler[ad] = len(sec)
        sekme(wb, ad, lb, [[s["kod"], s["ad"], None, None] for s in sec],
              genislik={"ItemCode": 18, "ItemDescription": 45})

    # ── NE DEĞİŞECEK ─────────────────────────────────────────────────────
    deg = [u for u in urunler if u.h["degisiyor"]]
    rows = []
    for u in deg:
        fark = u.h["fiyat"] - u.h["nebim_m2"]
        S = u.etiket
        rows.append([u.kod, u.ad, u.marka, u.alt_grup, u.sezon, S, u.h["nebim_m2"], u.h["fiyat"], fark,
                     "▲ YÜKSELİR" if fark > 0 else "▼ düşer",
                     max(0, 1 - u.h["nebim_m2"] / S) if S else None, max(0, 1 - u.h["fiyat"] / S) if S else None,
                     u.h["kaynak"]])
    ws = sekme(wb, "NE DEĞİŞECEK",
               ["Ürün Kodu", "Ürün Adı", "Marka", "Alt Grup", "Sezon", "Etiket", "Nebim'de Şu An", "Yeni M2",
                "Fark", "Yön", "Eski Oran", "Yeni Oran", "Oranın Kaynağı"], rows,
               genislik={"Ürün Adı": 36, "Oranın Kaynağı": 60},
               para=("Etiket", "Nebim'de Şu An", "Yeni M2", "Fark"), yuzde=("Eski Oran", "Yeni Oran"),
               not_metni="Nebim'deki mevcut M2 fiyatı ile yeni M2 fiyatını karşılaştırır. Sadece DEĞİŞECEK ürünler.")
    for row in ws.iter_rows(min_row=3):
        row[9].fill = KIRMIZI if str(row[9].value).startswith("▲") else YESIL

    # ── HATALAR ──────────────────────────────────────────────────────────
    by = {u.kod: u for u in urunler}
    rows = []
    for s in kamp_satirlari:
        if not s["hata_listesine"]:
            continue
        u = by.get(s["kod"])
        rows.append([s["satir"], s["kod"], s["ad"], s["marka"], s["reyon"],
                     ("DEPO" if u.h["depo"] else "MAĞAZA") if u else "",
                     s["list_code"], s["list_desc"], s["kart_metin"], s["olmasi_gereken"],
                     u.h["oran"] if u else None, u.h["kaynak"] if u else "", s["durum"], s["oneri"],
                     s["bulundugu"], s["kac_liste"], s["sira_metni"], s["fiyattan_ind"],
                     "" if s["listede"] else "STOKTA YOK", s["cakisma"],
                     f"MÜKERRER ({s['mukerrer']} satır)" if s["mukerrer"] > 1 else ""])
    ws = sekme(wb, "HATALAR",
               ["Ham Satır", "Ürün Kodu", "Ürün Adı", "Marka", "Reyon", "KONUM", "ItemListCode",
                "ItemListDescription", "Ürün Kartı Oranı", "Kampanya Listesine Göre", "UYGULANAN ORAN",
                "ORAN KAYNAĞI", "DURUM", "ÖNERİ", "Bulunduğu Kampanyalar", "Kaç Listede", "Bu Ürünün Satırı",
                "Fiyattan İndirim %", "Stok Durumu", "Kampanya Çakışması", "Mükerrer Kayıt"], rows,
               genislik={"Ürün Adı": 34, "ORAN KAYNAĞI": 45, "ÖNERİ": 50, "Bulunduğu Kampanyalar": 60,
                         "ItemListDescription": 32},
               yuzde=("UYGULANAN ORAN",),
               not_metni="Kart Oranı ve Kampanya Listesi NEBİM verisidir; UYGULANAN ORAN bu programın kararıdır. "
                         "Uygulanan oranla aynı listedeki satır KALSIN, farklı olan Nebim'den SİLİNMELİ.")
    for row in ws.iter_rows(min_row=3):
        row[10].fill = row[11].fill = TURUNCU
        if str(row[13].value).startswith("SİL"):
            row[13].fill = KIRMIZI

    # ── ÜRÜN DETAY (eski Hesaplama sayfasının okunur hali) ───────────────
    detay_bas = ["Ürün Kodu", "Ürün Adı", "Marka", "Reyon", "Alt Grup", "Sezon", "Etkin Sezon", "İşaret",
                 "Envanter", "Mağaza Stok", "Depo Stok", "Konum", "Kapsam", "Stoksuz?", "Kart Oranı",
                 "Kampanya", "Kampanya Oranı", "Uyan Kural", "Sistem Önerisi", "BENİM ORANIM",
                 "UYGULANAN ORAN", "Etiket", "MAĞAZA FİYATI", "Fiili İndirim", "Nebim'de Şu An", "Değişiyor?",
                 "NEBİM Listesi", "M2'ye Girer?", "KAYNAK", "Maliyet (KDV'li)", "Başabaş Fiyat", "Net Kâr TL",
                 "Net Marj", "Durum", "Depo Notu", "Yeni/Eski", "Girdi Sekmesi"]
    rows = []
    for u in urunler:
        h = u.h
        rows.append([u.kod, u.ad, u.marka, u.reyon, u.alt_grup, u.sezon, h["etkin_sezon"], u.isaret or "—",
                     u.envanter, h["stok_magaza"], h["depo_stogu"], "DEPO" if h["depo"] else "MAĞAZA",
                     "İÇİ" if h["kapsam_ici"] else "DIŞI", "EVET" if h["stoksuz"] else "", h["kart_oran"] or None,
                     h["kampanya_adi"], h["kampanya_orani"] if h["kampanyada"] else None, h["kural"],
                     h["sistem_onerisi"], h["manuel"], h["oran"], u.etiket, h["fiyat"], h["fiili_indirim"],
                     h["nebim_m2"], "EVET" if h["degisiyor"] else "",
                     f"%{h['nebim_listesi']}" if h["nebim_listesi"] else "", "EVET" if h["m2_aktarim"] else "",
                     h["kaynak"], h["maliyet"], h["basabas"], h["net_kar"], h["net_marj"], h["durum"],
                     h["depo_notu"], u.yeni_eski, u.kaynak_sekme])
    ws = sekme(wb, "ÜRÜN DETAY", detay_bas, rows,
               genislik={"Ürün Adı": 34, "KAYNAK": 60, "Uyan Kural": 36, "Kampanya": 24},
               para=("Etiket", "MAĞAZA FİYATI", "Nebim'de Şu An", "Maliyet (KDV'li)", "Başabaş Fiyat",
                     "Net Kâr TL"),
               yuzde=("Kart Oranı", "Kampanya Oranı", "Sistem Önerisi", "BENİM ORANIM", "UYGULANAN ORAN",
                      "Fiili İndirim", "Net Marj"))
    for row in ws.iter_rows(min_row=2):
        row[20].fill = row[22].fill = TURUNCU
        if row[19].value is not None:
            row[19].fill = SARI
        if row[33].value == "⚠ ZARARDA":
            row[33].fill = KIRMIZI

    # ── ÖZET ─────────────────────────────────────────────────────────────
    ozet.append(["İNDİRİM VE MAĞAZA FİYAT — ÖZET"])
    ozet["A1"].font = Font(bold=True, size=14)
    ozet.append([f"Hesap tarihi: {bugun:%d.%m.%Y}   ·   Oluşturma: {dt.datetime.now():%d.%m.%Y %H:%M}"])
    ozet.append([])
    bilgiler = [
        ("Toplam ürün", len(urunler)),
        ("Kampanya listesi satırı", len(kamp_satirlari)),
        ("Elle oran (BENİM ORANIM) girilen ürün", len(manuel)),
        ("Depo sayılan ürün", sum(u.h["depo"] for u in urunler)),
        ("Stoksuz ürün", sum(u.h["stoksuz"] for u in urunler)),
        ("Depo tavanı uygulanan ürün", sum(u.h["tavan_uygulandi"] for u in urunler)),
        ("Nebim Aktarım (M2) ürün sayısı", len(m2)),
        *adetler.items(),
        ("Fiyatı değişecek ürün", len(deg)),
        ("   ↳ yükselecek", sum(u.h["fiyat"] > u.h["nebim_m2"] for u in deg)),
        ("   ↳ düşecek", sum(u.h["fiyat"] < u.h["nebim_m2"] for u in deg)),
        ("HATALAR listesindeki satır", sum(s["hata_listesine"] for s in kamp_satirlari)),
        ("Zararda satılan ürün (gider sonrası)", sum(u.h["durum"] == "⚠ ZARARDA" for u in urunler)),
    ]
    ozet.append(["Kalem", "Adet"])
    _baslik_bicimle(ozet, 4)
    for k, v in bilgiler:
        ozet.append([k, v])
    tanimsiz = sorted({s["list_desc"] for s in kamp_satirlari if s["durum"] == "TANIMSIZ KAMPANYA"})
    if tanimsiz:
        ozet.append([])
        ozet.append(["⚠ ayarlar.xlsx › Kampanyalar sekmesinde tanımlı olmayan kampanya adları:"])
        ozet.cell(ozet.max_row, 1).font = Font(bold=True, color="C00000")
        for t in tanimsiz:
            ozet.append([t])
    ozet.column_dimensions["A"].width = 48
    ozet.column_dimensions["B"].width = 12
    wb.save(yol)


# ═══════════════════════════════════════════════════════════════════════════
#  ANA PROGRAM
# ═══════════════════════════════════════════════════════════════════════════
def calistir(girdi: Path, ayar_yolu: Path, cikti: Path | None, bugun: dt.date):
    ay = ayar_oku(ayar_yolu)
    wb = openpyxl.load_workbook(girdi, data_only=True, read_only=False)
    urunler = urunleri_oku(wb)
    if not urunler:
        sys.exit("⚠ Girdi dosyasında ürün bulunamadı (ÜRÜN 1 / 1B / 2 / EK ÜRÜNLER sekmeleri).")
    kamp = kampanya_oku(wb)
    depo = depo_oku(wb, ay)
    manuel = manuel_oranlari_oku(wb)
    by = {u.kod: u for u in urunler}
    _, birincil = kampanya_isle(kamp, by, ay)
    hesapla(urunler, kamp, birincil, depo, manuel, ay, bugun)
    cikti = cikti or girdi.with_name(f"CIKTI_{bugun:%Y%m%d}.xlsx")
    cikti_yaz(cikti, urunler, kamp, ay, bugun, manuel)
    print(f"✔ {len(urunler)} ürün, {len(kamp)} kampanya satırı işlendi"
          f"{'' if depo is not None else ' (depo raporu bulunamadı)'}.")
    print(f"✔ Çıktı: {cikti}")
    return urunler, kamp


def main():
    p = argparse.ArgumentParser(description="İndirim ve mağaza fiyat hesaplayıcı")
    p.add_argument("girdi", nargs="?", type=Path, help="Ham verilerin bulunduğu Excel dosyası")
    p.add_argument("-a", "--ayarlar", type=Path, default=Path(__file__).with_name("ayarlar.xlsx"))
    p.add_argument("-o", "--cikti", type=Path)
    p.add_argument("-t", "--tarih", help="Hesap tarihi (GG.AA.YYYY). Boş = bugün.")
    p.add_argument("--ayar-olustur", action="store_true", help="Varsayılan ayarlar.xlsx dosyasını yazar")
    a = p.parse_args()
    if a.ayar_olustur:
        ayar_dosyasi_yaz(a.ayarlar)
        print(f"✔ Ayar dosyası yazıldı: {a.ayarlar}")
        return
    if not a.girdi:
        p.error("girdi dosyası gerekli")
    bugun = tarih(a.tarih) if a.tarih else dt.date.today()
    if not a.ayarlar.exists():
        print(f"ⓘ {a.ayarlar} bulunamadı, varsayılan ayarlar kullanılıyor ve dosya oluşturuluyor.")
        ayar_dosyasi_yaz(a.ayarlar)
    calistir(a.girdi, a.ayarlar, a.cikti, bugun)


if __name__ == "__main__":
    main()
