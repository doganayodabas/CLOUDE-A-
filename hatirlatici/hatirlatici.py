#!/usr/bin/env python3
"""
HATIRLATICI — ödeme, çek, banka ve personel takip programı.

Kullanım:
    python hatirlatici.py                       # HATIRLATICI.xlsx okunur, rapor üretilir
    python hatirlatici.py veri.xlsx -t 08.11.2026 -o rapor.xlsx
    python hatirlatici.py --sablon              # boş HATIRLATICI.xlsx yazar
    python hatirlatici.py --ornek               # örnek verili ORNEK_HATIRLATICI.xlsx yazar

Girdi dosyasındaki sekmeler (veriyi siz girersiniz):
    AYARLAR, BANKALAR, BANKA İŞLEMLERİ, VERİLEN ÇEKLER, KREDİ KARTLARI,
    KART EKSTRELERİ, DBS, FATURALAR, DÜZENLİ ÖDEMELER, DİĞER ÖDEMELER,
    PERSONEL, İZİNLER, TATİLLER

Çıktı:
    HATIRLATICI_RAPOR_YYYYAAGG.xlsx  (ÖZET, YAKLAŞAN ÖDEMELER, ÇEK TAKVİMİ,
                                      BANKA DURUMU, RAPORLAR, PERSONEL İZİN,
                                      İZİN TAKVİMİ, TÜM KAYITLAR)
    HATIRLATICI.ics                  (telefon / Outlook / Google takvimine alınır)
"""
from __future__ import annotations

import argparse
import calendar
import datetime as dt
import hashlib
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover
    sys.exit("openpyxl gerekli:  pip install openpyxl")


# ═══════════════════════════════════════════════════════════════════════════
#  YARDIMCI FONKSİYONLAR
# ═══════════════════════════════════════════════════════════════════════════
def trim(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def up(v) -> str:
    """Türkçe i/ı duyarlı büyük harf."""
    return trim(v).replace("i", "İ").replace("ı", "I").upper()


def tutar(v):
    """Sayı → float. '1.234,56' / '1234.56' / '1.234,56 ₺' okunur. Boş → None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def tarih(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:  # Excel seri numarası
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    if isinstance(v, str):
        s = v.strip()[:10]
        for f in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
            try:
                return dt.datetime.strptime(s, f).date()
            except ValueError:
                pass
    return None


def tamsayi(v):
    t = tutar(v)
    return int(t) if t is not None else None


def evet(v) -> bool:
    return up(v).startswith("E")


def tr_tarih(d) -> str:
    return d.strftime("%d.%m.%Y") if d else ""


def tl(x) -> str:
    if x is None:
        return "?"
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " ₺"


def ay_ekle(d: dt.date, n: int) -> dt.date:
    a = d.month - 1 + n
    y, m = d.year + a // 12, a % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def yil_ekle(d: dt.date, n: int) -> dt.date:
    try:
        return d.replace(year=d.year + n)
    except ValueError:  # 29 Şubat
        return d.replace(year=d.year + n, day=28)


def gun_ay(y: int, m: int, gun: int) -> dt.date:
    return dt.date(y, m, max(1, min(gun, calendar.monthrange(y, m)[1])))


AYLAR = ["", "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz",
         "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]


def ay_adi(y: int, m: int) -> str:
    return f"{AYLAR[m]} {y}"


# ═══════════════════════════════════════════════════════════════════════════
#  ŞABLON TANIMI  (sekme → sütunlar)
#  tip: t=tarih  n=tutar  g=tam sayı  s=metin
# ═══════════════════════════════════════════════════════════════════════════
DURUM = ["BEKLİYOR", "ÖDENDİ", "İPTAL"]
EH = ["EVET", "HAYIR"]
FATURA_TURLERI = ["ELEKTRİK", "SU", "DOĞALGAZ", "İNTERNET", "TELEFON", "AİDAT", "DİĞER"]
DUZENLI_TURLERI = ["KİRA", "SGK", "KDV", "MUHTASAR", "GEÇİCİ VERGİ", "KREDİ TAKSİTİ", "MAAŞ",
                   "SİGORTA", "AİDAT", "ELEKTRİK", "SU", "DOĞALGAZ", "İNTERNET", "TELEFON",
                   "MUHASEBE", "DİĞER"]
PERIYOT = ["AYLIK", "2 AYLIK", "3 AYLIK", "6 AYLIK", "YILLIK"]
IZIN_TURLERI = ["YILLIK", "MAZERET", "RAPOR", "ÜCRETSİZ", "DOĞUM", "EVLİLİK", "ÖLÜM", "BABALIK", "DİĞER"]
ONAY = ["ONAYLANDI", "BEKLİYOR", "RED"]

# ODABAŞLAR bilançosundaki gider yapısı (grup → kalemler). GİDER KALEMLERİ sekmesinden değiştirilebilir.
GIDER_VARSAYILAN = [
    ("SATIŞLARIN MALİYETİ", ["Brüt mal alımı (ürün ödemeleri)"]),
    ("PERSONEL GİDERLERİ", ["Personel Maaşları", "Personel İcra Kesintisi Ödemeleri", "Personel Yemek (Setcart)",
                            "Personel Tazminat / İzin Ödemeleri", "SGK Primi", "Bağkur – Gökyıl Odabaş"]),
    ("VERGİ VE HARÇLAR", ["Vergi Ödemeleri"]),
    ("KİRA VE İŞLETME GİDERLERİ", ["Depo Kirası / Giderleri", "Otopark", "Mağaza Elektrik", "Depo Elektrik",
                                   "Muhasebe Telefon Hattı", "Online Telefon Hattı", "İnternet – Mağaza ve Depo",
                                   "Sigorta", "Bakım / Onarım", "Temizlik", "Market / Mutfak"]),
    ("SATIŞ VE PAZARLAMA GİDERLERİ", ["Elektrik Tamir / Bakım", "Kargo", "Terzi", "Kırtasiye Masrafları",
                                      "Online Faturalar"]),
    ("GENEL YÖNETİM GİDERLERİ", ["Muhasebe Gideri", "Hizmet / Danışmanlık", "Program Servis Destek Hizmeti",
                                 "Reklam / Görsel", "Demirbaş / Mobilya", "Yazılım / Abonelik",
                                 "Online Alışveriş (Trendyol / Hepsiburada)", "Diğer"]),
    ("FİNANSMAN GİDERLERİ", ["Banka ÖKC / POS Hizmet Bedeli",
                             "Banka İşlem / Hesap Masrafı (Havale / EFT Ücretleri)", "POS komisyonu (mağaza)",
                             "Kredi Taksiti / Faiz"]),
    ("SINIFLANDIRILMAMIŞ ÖDEMELER", ["Ev İnterneti (Yakanet)", "İstanbul Ev Giderleri", "Ev Elektrik Gideri",
                                     "Nermin Odabaş Ödemesi"]),
    ("ORTAKLARA YAPILAN ÖDEMELER", ["Huzur hakkı", "Emanet verilen"]),
    ("KREDİ KARTI", ["Kredi Kartı Ödemesi (dağıtılmamış)"]),
]
FAALIYET_GRUPLARI = ["PERSONEL GİDERLERİ", "VERGİ VE HARÇLAR", "KİRA VE İŞLETME GİDERLERİ",
                     "SATIŞ VE PAZARLAMA GİDERLERİ", "GENEL YÖNETİM GİDERLERİ", "FİNANSMAN GİDERLERİ",
                     "SINIFLANDIRILMAMIŞ ÖDEMELER"]
MASRAF_YERI = ["MAĞAZA", "ONLİNE", "ORTAK"]
GK = ("Gider Kalemi", 30, "s", "GIDER")
MY = ("Masraf Yeri", 10, "s", MASRAF_YERI)

SABLON: dict[str, list[tuple]] = {
    "BANKALAR": [
        ("Banka", 16, "s", None), ("Hesap Adı", 20, "s", None), ("IBAN", 32, "s", None),
        ("Bakiye", 15, "n", None), ("Bakiye Tarihi", 13, "t", None),
        ("KMH / Ek Hesap Limiti", 15, "n", None), ("Not", 30, "s", None),
    ],
    "BANKA İŞLEMLERİ": [
        ("Tarih", 12, "t", None), ("Banka", 16, "s", None), ("Hesap Adı", 20, "s", None),
        ("Tür", 10, "s", ["GİRİŞ", "ÇIKIŞ"]), ("Açıklama", 36, "s", None), ("Tutar", 15, "n", None), GK, MY,
    ],
    "VERİLEN ÇEKLER": [
        ("Çek No", 12, "s", None), ("Banka", 16, "s", None), ("Hesap Adı", 18, "s", None),
        ("Lehdar (Kime)", 26, "s", None), ("Keşide Tarihi", 13, "t", None), ("Vade Tarihi", 13, "t", None),
        ("Tutar", 15, "n", None), ("Durum", 12, "s", DURUM), ("Ödeme Tarihi", 13, "t", None),
        ("Açıklama", 30, "s", None), GK, MY,
    ],
    "KREDİ KARTLARI": [
        ("Kart Adı", 20, "s", None), ("Banka", 16, "s", None), ("Son 4 Hane", 10, "s", None),
        ("Limit", 14, "n", None), ("Hesap Kesim Günü", 10, "g", None), ("Son Ödeme Günü", 10, "g", None),
        ("Otomatik Ödeme", 10, "s", EH), ("Ödeme Bankası", 16, "s", None), ("Ödeme Hesabı", 18, "s", None),
        ("Aktif", 8, "s", EH),
    ],
    "KART EKSTRELERİ": [
        ("Kart Adı", 20, "s", None), ("Son Ödeme Tarihi", 13, "t", None), ("Ekstre Tutarı", 15, "n", None),
        ("Asgari Tutar", 14, "n", None), ("Ödenen Tutar", 14, "n", None), ("Ödeme Tarihi", 13, "t", None),
        ("Not", 30, "s", None),
    ],
    "DBS": [
        ("Firma", 24, "s", None), ("Banka", 16, "s", None), ("Hesap Adı", 18, "s", None),
        ("DBS Limiti", 15, "n", None), ("Fatura No", 14, "s", None), ("Fatura Tarihi", 13, "t", None),
        ("Vade Tarihi", 13, "t", None), ("Tutar", 15, "n", None), ("Durum", 12, "s", DURUM),
        ("Ödeme Tarihi", 13, "t", None), ("Not", 30, "s", None), GK, MY,
    ],
    "FATURALAR": [
        ("Tür", 12, "s", FATURA_TURLERI), ("Abone / Tesisat No", 18, "s", None), ("Yer (Mağaza/Depo)", 18, "s", None),
        ("Kurum", 16, "s", None), ("Son Ödeme Tarihi", 13, "t", None), ("Tutar", 14, "n", None),
        ("Durum", 12, "s", DURUM), ("Ödeme Tarihi", 13, "t", None), ("Otomatik Ödeme", 10, "s", EH),
        ("Ödeme Bankası", 16, "s", None), ("Not", 28, "s", None), GK, MY,
    ],
    "DÜZENLİ ÖDEMELER": [
        ("Tür", 14, "s", DUZENLI_TURLERI), ("Açıklama", 28, "s", None), ("Abone / Ref No", 16, "s", None),
        ("Ödeme Günü", 10, "g", None), ("Periyot", 10, "s", PERIYOT), ("Başlangıç", 12, "t", None),
        ("Bitiş", 12, "t", None), ("Tahmini Tutar", 14, "n", None), ("Ödeme Bankası", 16, "s", None),
        ("Ödeme Hesabı", 18, "s", None), ("Aktif", 8, "s", EH), GK, MY,
    ],
    "DİĞER ÖDEMELER": [
        ("Tür", 14, "s", None), ("Açıklama", 30, "s", None), ("Kime", 22, "s", None),
        ("Son Ödeme Tarihi", 13, "t", None), ("Tutar", 14, "n", None), ("Durum", 12, "s", DURUM),
        ("Ödeme Tarihi", 13, "t", None), ("Ödeme Bankası", 16, "s", None), ("Ödeme Hesabı", 18, "s", None),
        ("Not", 28, "s", None), GK, MY,
    ],
    "PERSONEL": [
        ("Sicil No", 9, "s", None), ("Ad Soyad", 22, "s", None), ("Görev", 16, "s", None),
        ("Mağaza / Bölüm", 18, "s", None), ("Telefon", 14, "s", None), ("Doğum Tarihi", 12, "t", None),
        ("İşe Giriş", 12, "t", None), ("Çıkış Tarihi", 12, "t", None),
        ("Devreden İzin (gün)", 11, "n", None), ("Devir Tarihi", 12, "t", None), ("Not", 26, "s", None),
    ],
    "İZİNLER": [
        ("Sicil No", 9, "s", None), ("Ad Soyad", 22, "s", None), ("İzin Türü", 12, "s", IZIN_TURLERI),
        ("Başlangıç", 12, "t", None), ("Bitiş", 12, "t", None), ("Gün", 7, "n", None),
        ("Onay", 12, "s", ONAY), ("Açıklama", 30, "s", None),
    ],
    "TATİLLER": [
        ("Tarih", 12, "t", None), ("Açıklama", 30, "s", None), ("Yarım Gün", 10, "s", EH),
    ],
    "GİDER KALEMLERİ": [
        ("Grup", 32, "s", None), ("Gider Kalemi", 52, "s", None),
    ],
}

AYAR_VARSAYILAN = [
    ("Firma Adı", "", "Rapor başlığında görünür."),
    ("Yaklaşan Ödeme Ufku (gün)", 30, "Bugünden itibaren kaç günlük ödemeler listelensin."),
    ("Acil Uyarı (gün)", 3, "Bu kadar gün kalan ödemeler turuncu işaretlenir."),
    ("Takvim Hatırlatma (gün önce)", "3,1", "Takvim dosyasında kaç gün önce alarm çalsın (virgülle)."),
    ("Çek Takvimi (ay)", 6, "Çek takviminde kaç ay ileri gösterilsin."),
    ("Rapor Geçmişi (ay)", 6, "RAPORLAR sekmesinde kaç ay geriye bakılsın."),
    ("Hafta Tatili", "PAZAR", "İzin günü hesabında sayılmayan gün(ler). Örn: PAZAR veya CUMARTESİ,PAZAR"),
    ("Doğum Günü Hatırlat (gün)", 7, "0 = hatırlatma yok."),
]

# Türkiye resmî tatilleri. Dini bayram tarihlerini Diyanet takviminden kontrol edin.
TATIL_VARSAYILAN = [
    ("2026-01-01", "Yılbaşı", False), ("2026-03-19", "Ramazan Bayramı Arifesi", True),
    ("2026-03-20", "Ramazan Bayramı 1. gün", False), ("2026-03-21", "Ramazan Bayramı 2. gün", False),
    ("2026-03-22", "Ramazan Bayramı 3. gün", False), ("2026-04-23", "Ulusal Egemenlik ve Çocuk Bayramı", False),
    ("2026-05-01", "Emek ve Dayanışma Günü", False), ("2026-05-19", "Atatürk'ü Anma, Gençlik ve Spor Bayramı", False),
    ("2026-05-26", "Kurban Bayramı Arifesi", True), ("2026-05-27", "Kurban Bayramı 1. gün", False),
    ("2026-05-28", "Kurban Bayramı 2. gün", False), ("2026-05-29", "Kurban Bayramı 3. gün", False),
    ("2026-05-30", "Kurban Bayramı 4. gün", False), ("2026-07-15", "Demokrasi ve Millî Birlik Günü", False),
    ("2026-08-30", "Zafer Bayramı", False), ("2026-10-28", "Cumhuriyet Bayramı Arifesi", True),
    ("2026-10-29", "Cumhuriyet Bayramı", False),
    ("2027-01-01", "Yılbaşı", False), ("2027-03-08", "Ramazan Bayramı Arifesi", True),
    ("2027-03-09", "Ramazan Bayramı 1. gün", False), ("2027-03-10", "Ramazan Bayramı 2. gün", False),
    ("2027-03-11", "Ramazan Bayramı 3. gün", False), ("2027-04-23", "Ulusal Egemenlik ve Çocuk Bayramı", False),
    ("2027-05-01", "Emek ve Dayanışma Günü", False), ("2027-05-15", "Kurban Bayramı Arifesi", True),
    ("2027-05-16", "Kurban Bayramı 1. gün", False), ("2027-05-17", "Kurban Bayramı 2. gün", False),
    ("2027-05-18", "Kurban Bayramı 3. gün", False), ("2027-05-19", "Kurban Bayramı 4. gün / 19 Mayıs", False),
    ("2027-07-15", "Demokrasi ve Millî Birlik Günü", False), ("2027-08-30", "Zafer Bayramı", False),
    ("2027-10-28", "Cumhuriyet Bayramı Arifesi", True), ("2027-10-29", "Cumhuriyet Bayramı", False),
]

ACIKLAMALAR = {
    "BANKALAR": "Her hesap bir satır. Bakiye'yi internet bankacılığında gördüğünüz tutarla, Bakiye Tarihi ile girin. "
                "Bu tarihten sonraki BANKA İŞLEMLERİ bakiyeye eklenir/çıkarılır.",
    "BANKA İŞLEMLERİ": "İsteğe bağlı hareket defteri. İleri tarihli GİRİŞ (POS blokesi, tahsilat) yazarsanız "
                       "BANKA DURUMU'ndaki projeksiyona dahil edilir. Gider Kalemi yazılan ÇIKIŞ'lar (POS komisyonu, "
                       "banka masrafı…) GİDER RAPORU'na girer — başka sekmede kayıtlı ödemeye gider kalemi yazmayın.",
    "VERİLEN ÇEKLER": "Ödenince Durum = ÖDENDİ yapın (veya Ödeme Tarihi girin). Banka + Hesap Adı, BANKALAR ile aynı yazılmalı.",
    "KREDİ KARTLARI": "Kart tanımı. Program her ay için Son Ödeme Günü'nden tahmini hatırlatma üretir. "
                      "Ekstre gelince KART EKSTRELERİ'ne yazın; tahmini satırın yerine o geçer.",
    "KART EKSTRELERİ": "Kart Adı, KREDİ KARTLARI'ndaki ile aynı olmalı. Ödenen Tutar ≥ Ekstre Tutarı ise ödendi sayılır.",
    "DBS": "Doğrudan Borçlandırma Sistemi faturaları. DBS Limiti firma+banka için bir kez yazılması yeterli.",
    "FATURALAR": "Elektrik, su, doğalgaz… Fatura gelince satır ekleyin. Aynı Tür + Abone No DÜZENLİ ÖDEMELER'de varsa "
                 "o ayın tahmini hatırlatması yerine bu satır geçer.",
    "DÜZENLİ ÖDEMELER": "Kira, SGK, vergi, kredi taksiti… Program her dönem için hatırlatma üretir (bugünden ileri).",
    "DİĞER ÖDEMELER": "Tek seferlik / tarihi yaklaşan diğer ödemeler.",
    "PERSONEL": "Çıkış Tarihi dolu olan personel pasif sayılır. Devreden İzin: Devir Tarihi itibarıyla kalan yıllık izin günü.",
    "İZİNLER": "Gün boşsa hesaplanır: YILLIK/MAZERET için hafta tatili ve resmî tatiller düşülür, diğerlerinde takvim günü.",
    "GİDER KALEMLERİ": "Bilançodaki gider yapısı. Satır ekleyebilir/değiştirebilirsiniz; diğer sekmelerdeki "
                       "'Gider Kalemi' açılır listesi buradan gelir.",
    "TATİLLER": "Resmî tatiller. Dini bayram tarihlerini her yıl Diyanet takviminden kontrol edin.",
}

# ═══════════════════════════════════════════════════════════════════════════
#  STİLLER
# ═══════════════════════════════════════════════════════════════════════════
BASLIK_FILL = PatternFill("solid", fgColor="1F3864")
BASLIK_FONT = Font(bold=True, color="FFFFFF")
ALT_BASLIK_FILL = PatternFill("solid", fgColor="D9E1F2")
KIRMIZI = PatternFill("solid", fgColor="F8CBAD")
TURUNCU = PatternFill("solid", fgColor="FCE4D6")
SARI = PatternFill("solid", fgColor="FFF2CC")
YESIL = PatternFill("solid", fgColor="E2EFDA")
GRI = PatternFill("solid", fgColor="EDEDED")
INCE = Side(style="thin", color="BFBFBF")
CERCEVE = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)
PARA = '#,##0.00 "₺"'
TARIH = "DD.MM.YYYY"


# ═══════════════════════════════════════════════════════════════════════════
#  ŞABLON YAZMA
# ═══════════════════════════════════════════════════════════════════════════
def sablon_yaz(yol: Path, ornek: bool = False, bugun: dt.date | None = None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AYARLAR"
    ws.append(["Ayar", "Değer", "Açıklama"])
    for a in AYAR_VARSAYILAN:
        ws.append(list(a))
    _baslik(ws, 1, 3)
    for c, w in zip("ABC", (30, 14, 80)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A2"

    veriler = ornek_veri(bugun or dt.date.today()) if ornek else {}
    for ad, sutunlar in SABLON.items():
        ws = wb.create_sheet(ad)
        ws.append([s[0] for s in sutunlar])
        _baslik(ws, 1, len(sutunlar))
        ws.freeze_panes = "A2"
        for i, (bas, gen, tip, secenek) in enumerate(sutunlar, 1):
            harf = get_column_letter(i)
            ws.column_dimensions[harf].width = gen
            if secenek:
                f1 = "='GİDER KALEMLERİ'!$B$2:$B$300" if secenek == "GIDER" else '"' + ",".join(secenek) + '"'
                dv = DataValidation(type="list", formula1=f1, allow_blank=True)
                dv.showErrorMessage = False  # listede olmayan değer de yazılabilsin
                ws.add_data_validation(dv)
                dv.add(f"{harf}2:{harf}2000")
            for r in range(2, 501):
                hucre = ws.cell(r, i)
                if tip == "t":
                    hucre.number_format = TARIH
                elif tip == "n":
                    hucre.number_format = PARA
        satirlar = veriler.get(ad, [])
        if ad == "TATİLLER":
            satirlar = [[dt.date.fromisoformat(t), a, "EVET" if y else ""] for t, a, y in TATIL_VARSAYILAN]
        elif ad == "GİDER KALEMLERİ":
            satirlar = [[g, k] for g, ks in GIDER_VARSAYILAN for k in ks]
        for r, satir in enumerate(satirlar, 2):
            for c, v in enumerate(satir, 1):
                ws.cell(r, c, v)
        if ad in ACIKLAMALAR:
            ws.cell(1, 1).comment = openpyxl.comments.Comment(ACIKLAMALAR[ad], "Hatırlatıcı", width=320, height=110)
    wb.save(yol)


def _baslik(ws, satir, n, baslangic=1):
    for c in range(baslangic, baslangic + n):
        h = ws.cell(satir, c)
        h.fill, h.font = BASLIK_FILL, BASLIK_FONT
        h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        h.border = CERCEVE
    ws.row_dimensions[satir].height = 30


def ornek_veri(b: dt.date) -> dict[str, list[list]]:
    """Programı denemek için örnek kayıtlar (bugüne göre)."""
    g = lambda n: b + dt.timedelta(days=n)  # noqa: E731
    return {
        "BANKALAR": [
            ["GARANTİ", "TİCARİ", "TR00 0006 2000 0000 0000 0000 01", 185000, g(-1), 50000, ""],
            ["İŞ BANKASI", "TİCARİ", "TR00 0006 4000 0000 0000 0000 02", 42000, g(-1), 25000, ""],
            ["ZİRAAT", "MAAŞ", "TR00 0001 0000 0000 0000 0000 03", 340000, g(-1), 0, ""],
        ],
        "BANKA İŞLEMLERİ": [
            [g(0), "GARANTİ", "TİCARİ", "GİRİŞ", "POS tahsilatı", 23500],
            [g(4), "İŞ BANKASI", "TİCARİ", "GİRİŞ", "POS blokesi çözülecek", 60000],
            [g(-10), "GARANTİ", "TİCARİ", "ÇIKIŞ", "Ekim POS komisyonu", 7699.54, "POS komisyonu (mağaza)", "MAĞAZA"],
            [g(-9), "GARANTİ", "TİCARİ", "ÇIKIŞ", "Havale / EFT ücretleri", 777.10,
             "Banka İşlem / Hesap Masrafı (Havale / EFT Ücretleri)", "MAĞAZA"],
        ],
        "VERİLEN ÇEKLER": [
            ["100231", "GARANTİ", "TİCARİ", "ABC TEKSTİL", g(-60), g(-2), 75000, "BEKLİYOR", None, "Gecikmiş örnek"],
            ["100232", "GARANTİ", "TİCARİ", "ABC TEKSTİL", g(-50), g(2), 90000, "BEKLİYOR", None, ""],
            ["100233", "İŞ BANKASI", "TİCARİ", "XYZ GİYİM", g(-40), g(9), 120000, "BEKLİYOR", None, ""],
            ["100234", "GARANTİ", "TİCARİ", "DEF AYAKKABI", g(-30), g(25), 48000, "BEKLİYOR", None, ""],
            ["100235", "İŞ BANKASI", "TİCARİ", "XYZ GİYİM", g(-20), g(55), 135000, "BEKLİYOR", None, ""],
            ["100236", "GARANTİ", "TİCARİ", "ABC TEKSTİL", g(-20), g(95), 80000, "BEKLİYOR", None, ""],
            ["100230", "GARANTİ", "TİCARİ", "ABC TEKSTİL", g(-90), g(-30), 70000, "ÖDENDİ", g(-30), ""],
        ],
        "KREDİ KARTLARI": [
            ["GARANTİ BONUS", "GARANTİ", "1234", 150000, 3, 13, "HAYIR", "GARANTİ", "TİCARİ", "EVET"],
            ["İŞ MAXİMUM", "İŞ BANKASI", "5678", 100000, 20, 30, "EVET", "İŞ BANKASI", "TİCARİ", "EVET"],
        ],
        "KART EKSTRELERİ": [
            ["GARANTİ BONUS", gun_ay(b.year, b.month, 13) if b.day <= 13 else gun_ay(ay_ekle(b, 1).year, ay_ekle(b, 1).month, 13),
             38450.75, 7690.15, 0, None, ""],
        ],
        "DBS": [
            ["MAVİ GİYİM", "GARANTİ", "TİCARİ", 500000, "MG2026-1", g(-30), g(5), 64000, "BEKLİYOR", None, ""],
            ["MAVİ GİYİM", "GARANTİ", "TİCARİ", None, "MG2026-2", g(-15), g(20), 82000, "BEKLİYOR", None, ""],
            ["LCW", "İŞ BANKASI", "TİCARİ", 250000, "LC-881", g(-25), g(1), 31500, "BEKLİYOR", None, ""],
        ],
        "FATURALAR": [
            ["ELEKTRİK", "4001234567", "MAĞAZA", "ENERJİSA", g(6), 18450, "BEKLİYOR", None, "HAYIR", "GARANTİ", ""],
            ["ELEKTRİK", "4007654321", "DEPO", "ENERJİSA", g(-1), 3200, "BEKLİYOR", None, "HAYIR", "", "Gecikmiş örnek"],
            ["SU", "SU-55821", "MAĞAZA", "İSKİ", g(-12), 950, "ÖDENDİ", g(-14), "HAYIR", "", ""],
        ],
        "DÜZENLİ ÖDEMELER": [
            ["KİRA", "Depo kirası", "", 5, "AYLIK", None, None, 11400, "GARANTİ", "TİCARİ", "EVET"],
            ["MAAŞ", "Personel maaşları", "", 1, "AYLIK", None, None, 280000, "ZİRAAT", "MAAŞ", "EVET"],
            ["DİĞER", "Program servis destek", "", 15, "AYLIK", None, None, 54992, "GARANTİ", "", "EVET",
             "Program Servis Destek Hizmeti", "ORTAK"],
            ["DİĞER", "Online faturalar", "", 20, "AYLIK", None, None, None, "İŞ BANKASI", "", "EVET",
             "Online Faturalar", "ONLİNE"],
            ["SGK", "SGK primi", "", 26, "AYLIK", None, None, 38000, "ZİRAAT", "MAAŞ", "EVET"],
            ["KDV", "KDV beyanı ödemesi", "", 28, "AYLIK", None, None, None, "GARANTİ", "TİCARİ", "EVET"],
            ["ELEKTRİK", "Mağaza elektrik", "4001234567", 15, "AYLIK", None, None, 17000, "GARANTİ", "", "EVET"],
            ["GEÇİCİ VERGİ", "Geçici vergi", "", 17, "3 AYLIK", dt.date(2026, 2, 17), None, None, "GARANTİ", "", "EVET"],
            ["KREDİ TAKSİTİ", "İşletme kredisi", "", 10, "AYLIK", None, g(150), 22500, "İŞ BANKASI", "TİCARİ", "EVET"],
        ],
        "DİĞER ÖDEMELER": [
            ["SİGORTA", "İşyeri sigortası yenileme", "ANADOLU SİGORTA", g(12), 14500, "BEKLİYOR", None, "GARANTİ", "", ""],
        ],
        "PERSONEL": [
            ["001", "AYŞE YILMAZ", "Mağaza Müdürü", "MAĞAZA", "0555 000 00 01", dt.date(1985, 3, 12), dt.date(2015, 4, 1), None, 6, dt.date(2025, 1, 1), ""],
            ["002", "MEHMET KAYA", "Satış Danışmanı", "SPOR", "0555 000 00 02", dt.date(1999, b.month, min(b.day + 3, 28)), dt.date(2022, b.month, min(b.day + 10, 28)), None, 0, None, ""],
            ["003", "ELİF DEMİR", "Kasiyer", "MAĞAZA", "0555 000 00 03", dt.date(1974, 7, 2), dt.date(2023, 9, 15), None, 0, None, ""],
            ["004", "CAN ÖZ", "Depo", "DEPO", "", None, dt.date(2021, 1, 4), g(-40), 0, None, "Ayrıldı"],
        ],
        "İZİNLER": [
            ["001", "AYŞE YILMAZ", "YILLIK", dt.date(2025, 7, 7), dt.date(2025, 7, 18), None, "ONAYLANDI", ""],
            ["002", "MEHMET KAYA", "YILLIK", g(-2), g(4), None, "ONAYLANDI", ""],
            ["003", "ELİF DEMİR", "RAPOR", g(10), g(11), None, "ONAYLANDI", ""],
            ["001", "AYŞE YILMAZ", "YILLIK", g(20), g(26), None, "BEKLİYOR", "Talep"],
        ],
    }


# ═══════════════════════════════════════════════════════════════════════════
#  OKUMA
# ═══════════════════════════════════════════════════════════════════════════
def sekme_bul(wb, ad):
    hedef = up(ad)
    for ws in wb.worksheets:
        if up(ws.title) == hedef:
            return ws
    for ws in wb.worksheets:
        if hedef in up(ws.title):
            return ws
    return None


class Satir(dict):
    """Başlık adıyla (büyük/küçük harf ve boşluk duyarsız) erişim. r('Vade') → 'VADE TARİHİ' de bulunur."""

    def __call__(self, *adlar):
        for ad in adlar:
            k = up(ad)
            if k in self and self[k] not in (None, ""):
                return self[k]
        for ad in adlar:  # kısmi eşleşme
            k = up(ad)
            for anahtar, v in self.items():
                if k in anahtar and v not in (None, ""):
                    return v
        return None


def sekme_oku(wb, ad) -> list[Satir]:
    ws = sekme_bul(wb, ad)
    if ws is None:
        return []
    basliklar, sonuc = None, []
    for i, satir in enumerate(ws.iter_rows(values_only=True), 1):
        if basliklar is None:
            if any(trim(v) for v in satir):
                basliklar = [up(v) for v in satir]
            continue
        if not any(trim(v) for v in satir):
            continue
        s = Satir({b: v for b, v in zip(basliklar, satir) if b})
        s["_SATIR"] = i
        sonuc.append(s)
    return sonuc


@dataclass
class Ayarlar:
    firma: str = ""
    ufuk: int = 30
    acil: int = 3
    alarmlar: list[int] = field(default_factory=lambda: [3, 1])
    cek_ay: int = 6
    gecmis_ay: int = 6
    hafta_tatili: set[int] = field(default_factory=lambda: {6})
    dogum_gunu: int = 7


def ayar_oku(wb) -> Ayarlar:
    a = Ayarlar()
    d = {up(r("Ayar")): r("Değer") for r in sekme_oku(wb, "AYARLAR") if r("Ayar")}

    def al(ad, var):
        v = d.get(up(ad))
        return var if v in (None, "") else v

    a.firma = trim(al("Firma Adı", ""))
    a.ufuk = tamsayi(al("Yaklaşan Ödeme Ufku (gün)", 30)) or 30
    a.acil = tamsayi(al("Acil Uyarı (gün)", 3)) or 0
    a.alarmlar = sorted({int(x) for x in re.findall(r"\d+", trim(al("Takvim Hatırlatma (gün önce)", "3,1")))}, reverse=True)
    a.cek_ay = tamsayi(al("Çek Takvimi (ay)", 6)) or 6
    a.gecmis_ay = tamsayi(al("Rapor Geçmişi (ay)", 6)) or 6
    gunler = {up(g): i for i, g in enumerate(GUNLER)}
    ht = {gunler[x] for x in (up(t) for t in trim(al("Hafta Tatili", "PAZAR")).split(",")) if x in gunler}
    a.hafta_tatili = ht or {6}
    a.dogum_gunu = tamsayi(al("Doğum Günü Hatırlat (gün)", 7)) or 0
    return a


@dataclass
class Takvim:
    tatil: dict[dt.date, tuple[str, bool]]  # tarih → (açıklama, yarım gün mü)
    hafta_tatili: set[int]

    def banka_gunu_mu(self, d: dt.date) -> bool:
        t = self.tatil.get(d)
        return d.weekday() < 5 and not (t and not t[1])

    def sonraki_banka_gunu(self, d: dt.date) -> dt.date:
        while not self.banka_gunu_mu(d):
            d += dt.timedelta(days=1)
        return d

    def is_gunu_degeri(self, d: dt.date) -> float:
        if d.weekday() in self.hafta_tatili:
            return 0.0
        t = self.tatil.get(d)
        if t:
            return 0.5 if t[1] else 0.0
        return 1.0

    def is_gunu_say(self, bas: dt.date, bit: dt.date) -> float:
        n, d = 0.0, bas
        while d <= bit:
            n += self.is_gunu_degeri(d)
            d += dt.timedelta(days=1)
        return n

    def sonraki_is_gunu(self, d: dt.date) -> dt.date:
        d += dt.timedelta(days=1)
        while self.is_gunu_degeri(d) == 0:
            d += dt.timedelta(days=1)
        return d


def takvim_oku(wb, ay: Ayarlar) -> Takvim:
    tatil = {}
    for r in sekme_oku(wb, "TATİLLER"):
        d = tarih(r("Tarih"))
        if d:
            tatil[d] = (trim(r("Açıklama")), evet(r("Yarım Gün")))
    if not tatil:
        tatil = {dt.date.fromisoformat(t): (a, y) for t, a, y in TATIL_VARSAYILAN}
    return Takvim(tatil, ay.hafta_tatili)


# ═══════════════════════════════════════════════════════════════════════════
#  ÖDEME KALEMLERİ
# ═══════════════════════════════════════════════════════════════════════════
@dataclass
class Kalem:
    vade: dt.date
    kategori: str        # ÇEK / KREDİ KARTI / DBS / FATURA / DÜZENLİ / DİĞER
    tur: str             # ELEKTRİK, KİRA, … (kategori içinde alt tür)
    aciklama: str
    kime: str
    tutar: float | None
    durum: str           # BEKLİYOR / ÖDENDİ / İPTAL / KISMİ
    banka: str = ""
    hesap: str = ""
    odeme_tarihi: dt.date | None = None
    tahmini: bool = False
    otomatik: bool = False
    not_: str = ""
    kaynak: str = ""
    anahtar: tuple = ()
    gider: str = ""      # bilanço gider kalemi
    yer: str = ""        # MAĞAZA / ONLİNE / ORTAK

    @property
    def bekliyor(self) -> bool:
        return self.durum in ("BEKLİYOR", "KISMİ")

    @property
    def baslik(self) -> str:
        return self.kategori if not self.tur or self.tur == self.kategori else f"{self.kategori} · {self.tur}"


def durum_coz(v, odeme_tarihi) -> str:
    d = up(v)
    if d.startswith("ÖDEN") or d.startswith("ODEN") or d in ("TAMAM", "KAPANDI", "OK"):
        return "ÖDENDİ"
    if d.startswith("İPTAL") or d.startswith("IPTAL"):
        return "İPTAL"
    if not d and odeme_tarihi:
        return "ÖDENDİ"
    return "BEKLİYOR"


ORTAK_KALEMLER = {up(x) for x in ("Depo Kirası / Giderleri", "Depo Elektrik", "İnternet – Mağaza ve Depo", "Sigorta",
                                   "Bakım / Onarım", "Market / Mutfak", "Muhasebe Gideri", "Hizmet / Danışmanlık")}


def varsayilan_gider(kategori: str, tur: str, metin: str) -> str:
    """Gider Kalemi boşsa türüne göre bilanço kalemi tahmin eder."""
    m = up(metin)
    if kategori in ("ÇEK", "DBS"):
        return "Brüt mal alımı (ürün ödemeleri)"
    if kategori == "KREDİ KARTI":
        return "Kredi Kartı Ödemesi (dağıtılmamış)"
    if tur == "ELEKTRİK":
        return "Depo Elektrik" if "DEPO" in m else "Ev Elektrik Gideri" if re.search(r"\bEV\b", m) else "Mağaza Elektrik"
    if tur == "İNTERNET":
        return "Ev İnterneti (Yakanet)" if re.search(r"\bEV\b", m) else "İnternet – Mağaza ve Depo"
    if tur == "TELEFON":
        return "Online Telefon Hattı" if "ONL" in m else "Muhasebe Telefon Hattı"
    harita = {"SGK": "SGK Primi", "MAAŞ": "Personel Maaşları", "KDV": "Vergi Ödemeleri", "MUHTASAR": "Vergi Ödemeleri",
              "GEÇİCİ VERGİ": "Vergi Ödemeleri", "VERGİ": "Vergi Ödemeleri", "SİGORTA": "Sigorta",
              "KİRA": "Depo Kirası / Giderleri", "MUHASEBE": "Muhasebe Gideri", "KREDİ TAKSİTİ": "Kredi Taksiti / Faiz",
              "KARGO": "Kargo", "TERZİ": "Terzi", "OTOPARK": "Otopark", "TEMİZLİK": "Temizlik",
              "BAĞKUR": "Bağkur – Gökyıl Odabaş", "REKLAM": "Reklam / Görsel", "YAZILIM": "Yazılım / Abonelik"}
    return harita.get(tur, "Diğer")


def varsayilan_yer(gider: str, metin: str) -> str:
    m = up(metin)
    if "ONL" in m:
        return "ONLİNE"
    if up(gider) in ORTAK_KALEMLER:
        return "ORTAK"
    return "MAĞAZA"


def gider_ata(k: Kalem, r: Satir | None):
    """Satırdaki Gider Kalemi / Masraf Yeri'ni okur, boşsa tahmin eder."""
    metin = " ".join((k.aciklama, k.kime, k.not_, k.tur))
    k.gider = trim(r("Gider Kalemi")) if r is not None and r("Gider Kalemi") else varsayilan_gider(k.kategori, k.tur, metin)
    yer = up(r("Masraf Yeri")) if r is not None and r("Masraf Yeri") else ""
    k.yer = ("ONLİNE" if yer.startswith("ONL") else yer) if yer else varsayilan_yer(k.gider, metin)
    return k


def kalemleri_topla(wb, bugun: dt.date, ay: Ayarlar) -> tuple[list[Kalem], list[str], dict]:
    kalemler: list[Kalem] = []
    uyarilar: list[str] = []
    dbs_limit: dict[tuple[str, str], float] = {}

    def eksik(sekme, r, ne):
        uyarilar.append(f"{sekme} satır {r['_SATIR']}: {ne} okunamadı, satır atlandı.")

    for r in sekme_oku(wb, "VERİLEN ÇEKLER"):
        v = tarih(r("Vade Tarihi", "Vade"))
        if not v:
            eksik("VERİLEN ÇEKLER", r, "Vade Tarihi")
            continue
        ot = tarih(r("Ödeme Tarihi"))
        no = trim(r("Çek No"))
        kalemler.append(gider_ata(Kalem(
            v, "ÇEK", "ÇEK", f"Çek {no}".strip(), trim(r("Lehdar", "Kime")), tutar(r("Tutar")),
            durum_coz(r("Durum"), ot), up(r("Banka")), up(r("Hesap Adı")), ot,
            not_=trim(r("Açıklama")), kaynak="VERİLEN ÇEKLER", anahtar=("ÇEK", no, r["_SATIR"])), r))

    for r in sekme_oku(wb, "DBS"):
        firma, banka = trim(r("Firma")), up(r("Banka"))
        lim = tutar(r("DBS Limiti"))
        if lim:
            dbs_limit[(up(firma), banka)] = lim
        v = tarih(r("Vade Tarihi", "Vade"))
        if not v:
            if tutar(r("Tutar")) is not None:
                eksik("DBS", r, "Vade Tarihi")
            continue
        ot = tarih(r("Ödeme Tarihi"))
        fno = trim(r("Fatura No"))
        kalemler.append(gider_ata(Kalem(
            v, "DBS", "DBS", f"DBS {fno}".strip(), firma, tutar(r("Tutar")), durum_coz(r("Durum"), ot),
            banka, up(r("Hesap Adı")), ot, not_=trim(r("Not")), kaynak="DBS", anahtar=("DBS", fno, r["_SATIR"])), r))

    for r in sekme_oku(wb, "FATURALAR"):
        v = tarih(r("Son Ödeme Tarihi", "Son Ödeme"))
        if not v:
            eksik("FATURALAR", r, "Son Ödeme Tarihi")
            continue
        tur = up(r("Tür")) or "FATURA"
        abone = trim(r("Abone", "Tesisat"))
        yer = trim(r("Yer"))
        ot = tarih(r("Ödeme Tarihi"))
        kalemler.append(gider_ata(Kalem(
            v, "FATURA", tur, " · ".join(x for x in (tur, yer, abone) if x), trim(r("Kurum")),
            tutar(r("Tutar")), durum_coz(r("Durum"), ot), up(r("Ödeme Bankası")), "", ot,
            otomatik=evet(r("Otomatik")), not_=trim(r("Not")), kaynak="FATURALAR", anahtar=(tur, up(abone))), r))

    for r in sekme_oku(wb, "DİĞER ÖDEMELER"):
        v = tarih(r("Son Ödeme Tarihi", "Tarih"))
        if not v:
            eksik("DİĞER ÖDEMELER", r, "Son Ödeme Tarihi")
            continue
        tur, acik = up(r("Tür")) or "DİĞER", trim(r("Açıklama"))
        ot = tarih(r("Ödeme Tarihi"))
        kalemler.append(gider_ata(Kalem(
            v, "DİĞER", tur, acik or tur, trim(r("Kime")), tutar(r("Tutar")), durum_coz(r("Durum"), ot),
            up(r("Ödeme Bankası")), up(r("Ödeme Hesabı")), ot, not_=trim(r("Not")),
            kaynak="DİĞER ÖDEMELER", anahtar=(tur, up(acik))), r))

    # Kredi kartları: tanım + ekstreler
    kartlar = {}
    for r in sekme_oku(wb, "KREDİ KARTLARI"):
        ad = trim(r("Kart Adı"))
        if ad:
            kartlar[up(ad)] = r
    for r in sekme_oku(wb, "KART EKSTRELERİ"):
        ad = trim(r("Kart Adı"))
        v = tarih(r("Son Ödeme Tarihi", "Son Ödeme"))
        if not ad or not v:
            eksik("KART EKSTRELERİ", r, "Kart Adı / Son Ödeme Tarihi")
            continue
        k = kartlar.get(up(ad), Satir())
        if not k:
            uyarilar.append(f"KART EKSTRELERİ satır {r['_SATIR']}: '{ad}' KREDİ KARTLARI'nda tanımlı değil.")
        ekstre, odenen, asgari = tutar(r("Ekstre Tutarı")), tutar(r("Ödenen Tutar")) or 0, tutar(r("Asgari Tutar"))
        if ekstre is not None and ekstre - odenen <= 0.005:
            durum, kalan = "ÖDENDİ", ekstre
        elif odenen > 0:
            durum, kalan = "KISMİ", (ekstre - odenen) if ekstre is not None else None
        else:
            durum, kalan = "BEKLİYOR", ekstre
        notlar = [trim(r("Not"))]
        if asgari:
            notlar.append(f"Asgari: {tl(asgari)}")
        if odenen and durum == "KISMİ":
            notlar.append(f"Ödenen: {tl(odenen)}")
        oto = evet(k("Otomatik Ödeme"))
        if oto and durum != "ÖDENDİ":
            notlar.append("Otomatik ödeme talimatı — hesapta bakiye olmalı")
        kalemler.append(gider_ata(Kalem(
            v, "KREDİ KARTI", "KREDİ KARTI", f"{ad} ekstresi", up(k("Banka")) or "", kalan, durum,
            up(k("Ödeme Bankası")) or up(k("Banka")), up(k("Ödeme Hesabı")), tarih(r("Ödeme Tarihi")),
            otomatik=oto, not_=" · ".join(x for x in notlar if x), kaynak="KART EKSTRELERİ",
            anahtar=("KREDİ KARTI", up(ad))), None))

    # Tekrarlayan kalemler (kart son ödeme günleri + düzenli ödemeler) — bugünden ufka kadar
    acik_aylar = {(k.anahtar[:2], k.vade.year, k.vade.month) for k in kalemler if len(k.anahtar) >= 2}
    son = bugun + dt.timedelta(days=max(ay.ufuk, 31 * ay.cek_ay))

    def aylar(baslangic_ay: dt.date, periyot: int):
        d = dt.date(bugun.year, bugun.month, 1)
        while d <= son:
            fark = (d.year - baslangic_ay.year) * 12 + d.month - baslangic_ay.month
            if fark >= 0 and fark % periyot == 0:
                yield d.year, d.month
            d = ay_ekle(d, 1)

    for anahtar, k in kartlar.items():
        if k("Aktif") is not None and not evet(k("Aktif")):
            continue
        gun = tamsayi(k("Son Ödeme Günü"))
        if not gun:
            continue
        ad = trim(k("Kart Adı"))
        for y, m in aylar(dt.date(2000, 1, 1), 1):
            v = gun_ay(y, m, gun)
            if v < bugun or (("KREDİ KARTI", anahtar), y, m) in acik_aylar:
                continue
            kalemler.append(gider_ata(Kalem(
                v, "KREDİ KARTI", "KREDİ KARTI", f"{ad} ekstresi", up(k("Banka")), None, "BEKLİYOR",
                up(k("Ödeme Bankası")) or up(k("Banka")), up(k("Ödeme Hesabı")), tahmini=True,
                otomatik=evet(k("Otomatik Ödeme")), not_="Tahmini — ekstre tutarını KART EKSTRELERİ'ne girin",
                kaynak="KREDİ KARTLARI", anahtar=("KREDİ KARTI", anahtar)), None))

    periyotlar = {"AYLIK": 1, "2 AYLIK": 2, "3 AYLIK": 3, "6 AYLIK": 6, "YILLIK": 12}
    for r in sekme_oku(wb, "DÜZENLİ ÖDEMELER"):
        if r("Aktif") is not None and not evet(r("Aktif")):
            continue
        gun = tamsayi(r("Ödeme Günü"))
        if not gun:
            eksik("DÜZENLİ ÖDEMELER", r, "Ödeme Günü")
            continue
        tur, acik, ref = up(r("Tür")) or "DİĞER", trim(r("Açıklama")), trim(r("Abone", "Ref"))
        p = periyotlar.get(re.sub(r"\s+", " ", up(r("Periyot")) or "AYLIK"), 1)
        bas, bit = tarih(r("Başlangıç")), tarih(r("Bitiş"))
        anahtar = (tur, up(ref or acik))
        cat = "FATURA" if tur in FATURA_TURLERI and tur not in ("DİĞER", "AİDAT") else "DÜZENLİ"
        for y, m in aylar(bas or dt.date(bugun.year, 1, 1), p):
            v = gun_ay(y, m, gun)
            if v < bugun or (bas and v < bas) or (bit and v > bit) or (anahtar, y, m) in acik_aylar:
                continue
            kalemler.append(gider_ata(Kalem(
                v, cat, tur, acik or tur, ref, tutar(r("Tahmini Tutar")), "BEKLİYOR",
                up(r("Ödeme Bankası")), up(r("Ödeme Hesabı")), tahmini=True,
                not_="Düzenli ödeme (tahmini)", kaynak="DÜZENLİ ÖDEMELER", anahtar=anahtar), r))

    kalemler.sort(key=lambda k: (k.vade, k.kategori, k.aciklama))
    return kalemler, uyarilar, dbs_limit


# ═══════════════════════════════════════════════════════════════════════════
#  BANKA
# ═══════════════════════════════════════════════════════════════════════════
@dataclass
class Hesap:
    banka: str
    ad: str
    iban: str
    bakiye: float
    bakiye_tarihi: dt.date | None
    kmh: float
    guncel: float = 0.0
    hareketler: list = field(default_factory=list)   # (tarih, açıklama, tutar ±, tür)

    @property
    def etiket(self):
        return f"{self.banka} / {self.ad}" if self.ad else self.banka


def bankalari_oku(wb, bugun: dt.date, uyarilar: list[str]):
    hesaplar: list[Hesap] = []
    for r in sekme_oku(wb, "BANKALAR"):
        b = up(r("Banka"))
        if not b:
            continue
        hesaplar.append(Hesap(b, up(r("Hesap Adı")), trim(r("IBAN")), tutar(r("Bakiye")) or 0.0,
                              tarih(r("Bakiye Tarihi")), tutar(r("KMH", "Limit")) or 0.0))
    islemler, banka_giderleri = [], []
    for r in sekme_oku(wb, "BANKA İŞLEMLERİ"):
        d, t = tarih(r("Tarih")), tutar(r("Tutar"))
        if not d or t is None:
            uyarilar.append(f"BANKA İŞLEMLERİ satır {r['_SATIR']}: Tarih / Tutar okunamadı.")
            continue
        tur = up(r("Tür"))
        isaret = -1 if tur.startswith("Ç") or tur.startswith("C") or t < 0 else 1
        islemler.append((d, up(r("Banka")), up(r("Hesap Adı")), isaret * abs(t), trim(r("Açıklama")),
                         "GİRİŞ" if isaret > 0 else "ÇIKIŞ"))
        if isaret < 0 and r("Gider Kalemi"):
            banka_giderleri.append(gider_ata(Kalem(
                d, "BANKA", "BANKA", trim(r("Açıklama")) or "Banka çıkışı", "", abs(t), "ÖDENDİ", up(r("Banka")),
                up(r("Hesap Adı")), d, kaynak="BANKA İŞLEMLERİ"), r))
    for h in hesaplar:
        h.guncel = h.bakiye
    for isl in islemler:
        h = hesap_bul(hesaplar, isl[1], isl[2])
        if not h:
            continue
        if h.bakiye_tarihi and isl[0] <= h.bakiye_tarihi:
            continue  # bakiyenin içinde
        if isl[0] <= bugun:
            h.guncel += isl[3]
        else:
            h.hareketler.append((isl[0], isl[4] or "Beklenen hareket", isl[3], isl[5]))
    return hesaplar, islemler, banka_giderleri


def hesap_bul(hesaplar: list[Hesap], banka: str, ad: str) -> Hesap | None:
    if not banka:
        return None
    adaylar = [h for h in hesaplar if h.banka == banka]
    if ad:
        for h in adaylar:
            if h.ad == ad:
                return h
    return adaylar[0] if adaylar else None


# ═══════════════════════════════════════════════════════════════════════════
#  PERSONEL / İZİN
# ═══════════════════════════════════════════════════════════════════════════
def yillik_izin_hakki(kidem_yil: int, yas: int | None) -> int:
    """4857 s. İş Kanunu m.53: 1-5 yıl (dahil) 14, 5-15 yıl 20, 15 yıl ve üzeri 26 gün.
    18 yaş ve altı ile 50 yaş ve üstü en az 20 gün."""
    hak = 14 if kidem_yil <= 5 else 20 if kidem_yil < 15 else 26
    if yas is not None and (yas <= 18 or yas >= 50):
        hak = max(hak, 20)
    return hak


def yas_hesapla(dogum: dt.date | None, gun: dt.date) -> int | None:
    if not dogum:
        return None
    return gun.year - dogum.year - ((gun.month, gun.day) < (dogum.month, dogum.day))


@dataclass
class Izin:
    sicil: str
    ad: str
    tur: str
    bas: dt.date
    bit: dt.date
    gun: float
    onay: str
    aciklama: str
    donus: dt.date


@dataclass
class Personel:
    sicil: str
    ad: str
    gorev: str
    bolum: str
    telefon: str
    dogum: dt.date | None
    giris: dt.date | None
    cikis: dt.date | None
    devreden: float
    devir_tarihi: dt.date | None
    izinler: list[Izin] = field(default_factory=list)

    def aktif(self, gun: dt.date) -> bool:
        return not self.cikis or self.cikis >= gun


def personel_oku(wb, tk: Takvim, uyarilar: list[str]):
    personel: list[Personel] = []
    for r in sekme_oku(wb, "PERSONEL"):
        ad = trim(r("Ad Soyad"))
        if not ad:
            continue
        personel.append(Personel(trim(r("Sicil")), ad, trim(r("Görev")), trim(r("Mağaza", "Bölüm")),
                                 trim(r("Telefon")), tarih(r("Doğum")), tarih(r("İşe Giriş", "Giriş")),
                                 tarih(r("Çıkış")), tutar(r("Devreden")) or 0.0, tarih(r("Devir Tarihi"))))
    sicil = {up(p.sicil): p for p in personel if p.sicil}
    isim = {up(p.ad): p for p in personel}
    izinler: list[Izin] = []
    for r in sekme_oku(wb, "İZİNLER"):
        bas, bit = tarih(r("Başlangıç")), tarih(r("Bitiş"))
        if not bas:
            uyarilar.append(f"İZİNLER satır {r['_SATIR']}: Başlangıç okunamadı.")
            continue
        bit = bit or bas
        if bit < bas:
            uyarilar.append(f"İZİNLER satır {r['_SATIR']}: Bitiş, başlangıçtan önce.")
            continue
        tur = up(r("İzin Türü", "Tür")) or "YILLIK"
        gun = tutar(r("Gün"))
        if gun is None:
            gun = tk.is_gunu_say(bas, bit) if tur in ("YILLIK", "MAZERET") else float((bit - bas).days + 1)
        onay = up(r("Onay")) or "ONAYLANDI"
        p = sicil.get(up(r("Sicil"))) or isim.get(up(r("Ad Soyad")))
        iz = Izin(p.sicil if p else trim(r("Sicil")), p.ad if p else trim(r("Ad Soyad")), tur, bas, bit, gun,
                  onay, trim(r("Açıklama")), tk.sonraki_is_gunu(bit))
        izinler.append(iz)
        if p:
            p.izinler.append(iz)
        else:
            uyarilar.append(f"İZİNLER satır {r['_SATIR']}: '{iz.ad or iz.sicil}' PERSONEL listesinde yok.")
    return personel, izinler


def izin_durumu(p: Personel, bugun: dt.date) -> dict:
    sonuc = {"kidem": None, "hak": None, "hak_edilen": 0, "kullanilan": 0.0, "planlanan": 0.0,
             "kalan": None, "sonraki_hak": None, "sonraki_hak_gun": None}
    if not p.giris:
        return sonuc
    ref = min(bugun, p.cikis) if p.cikis else bugun
    kidem = ref.year - p.giris.year - ((ref.month, ref.day) < (p.giris.month, p.giris.day))
    sonuc["kidem"] = max(kidem, 0)
    baslangic = p.devir_tarihi or p.giris
    hak_edilen = 0
    for k in range(1, kidem + 1):
        d = yil_ekle(p.giris, k)
        if d > baslangic or (not p.devir_tarihi and d >= baslangic):
            hak_edilen += yillik_izin_hakki(k, yas_hesapla(p.dogum, d))
    sonuc["hak_edilen"] = hak_edilen
    sonraki = yil_ekle(p.giris, kidem + 1)
    sonuc["sonraki_hak"] = sonraki
    sonuc["sonraki_hak_gun"] = yillik_izin_hakki(kidem + 1, yas_hesapla(p.dogum, sonraki))
    sonuc["hak"] = yillik_izin_hakki(max(kidem, 1), yas_hesapla(p.dogum, ref))
    for iz in p.izinler:
        if iz.tur != "YILLIK" or iz.onay == "RED":
            continue
        if p.devir_tarihi and iz.bas <= p.devir_tarihi:
            continue
        if iz.bas <= bugun:
            sonuc["kullanilan"] += iz.gun
        else:
            sonuc["planlanan"] += iz.gun
    sonuc["kalan"] = p.devreden + hak_edilen - sonuc["kullanilan"] - sonuc["planlanan"]
    return sonuc


# ═══════════════════════════════════════════════════════════════════════════
#  ÇIKTI YARDIMCILARI
# ═══════════════════════════════════════════════════════════════════════════
def blok_baslik(ws, satir, metin, genislik=8):
    h = ws.cell(satir, 1, metin)
    h.font = Font(bold=True, size=12, color="1F3864")
    for c in range(1, genislik + 1):
        ws.cell(satir, c).fill = ALT_BASLIK_FILL
    return satir + 1


def tablo(ws, satir, basliklar, satirlar, para=(), tarihler=(), renkler=None, sutun=1):
    """basliklar: list[str]. satirlar: list[list]. renkler: satır → PatternFill|None. Sonraki boş satırı döndürür."""
    for i, b in enumerate(basliklar):
        ws.cell(satir, sutun + i, b)
    _baslik(ws, satir, len(basliklar), sutun)
    for j, s in enumerate(satirlar, 1):
        dolgu = renkler[j - 1] if renkler else None
        for i, v in enumerate(s):
            h = ws.cell(satir + j, sutun + i, v)
            h.border = CERCEVE
            if i in para:
                h.number_format = PARA
            elif i in tarihler or isinstance(v, dt.date):
                h.number_format = TARIH
            if dolgu:
                h.fill = dolgu
    return satir + len(satirlar) + 2


def genislik(ws, degerler):
    for i, w in enumerate(degerler, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def kalan_metin(kalan: int) -> str:
    if kalan < 0:
        return f"GECİKTİ ({-kalan} gün)"
    if kalan == 0:
        return "BUGÜN"
    if kalan == 1:
        return "YARIN"
    return f"{kalan} gün"


def renk_kalan(kalan: int, acil: int):
    if kalan < 0:
        return KIRMIZI
    if kalan <= acil:
        return TURUNCU
    if kalan <= 7:
        return SARI
    return None


# ═══════════════════════════════════════════════════════════════════════════
#  RAPOR
# ═══════════════════════════════════════════════════════════════════════════
def rapor_yaz(yol: Path, kalemler, hesaplar, islemler, personel, izinler, dbs_limit,
              uyarilar, ay: Ayarlar, tk: Takvim, bugun: dt.date, banka_giderleri=(), gruplar=None):
    wb = openpyxl.Workbook()
    ufuk_son = bugun + dt.timedelta(days=ay.ufuk)
    bekleyen = [k for k in kalemler if k.bekliyor]
    yaklasan = [k for k in bekleyen if k.vade <= ufuk_son]

    # ── banka projeksiyonu (ÖZET'te de kullanılır)
    proj = banka_projeksiyonu(hesaplar, bekleyen, tk, bugun, ufuk_son)

    # ── personel hatırlatmaları
    p_hat = personel_hatirlatmalari(personel, izinler, ay, bugun)

    # ─────────────────────────────── ÖZET
    ws = wb.active
    ws.title = "ÖZET"
    genislik(ws, (34, 16, 16, 16, 16, 16, 16, 16))
    baslik = f"{ay.firma + ' — ' if ay.firma else ''}HATIRLATICI · {tr_tarih(bugun)} {GUNLER[bugun.weekday()]}"
    ws.cell(1, 1, baslik).font = Font(bold=True, size=15, color="1F3864")
    r = 3
    araliklar = [("Gecikmiş", None, bugun - dt.timedelta(days=1), KIRMIZI),
                 ("Bugün", bugun, bugun, TURUNCU),
                 (f"1–{ay.acil} gün", bugun + dt.timedelta(days=1), bugun + dt.timedelta(days=ay.acil), TURUNCU),
                 (f"{ay.acil + 1}–7 gün", bugun + dt.timedelta(days=ay.acil + 1), bugun + dt.timedelta(days=7), SARI),
                 (f"8–{ay.ufuk} gün", bugun + dt.timedelta(days=8), ufuk_son, None)]
    araliklar = [a for a in araliklar if a[1] is None or a[1] <= a[2]]
    kategoriler = ["ÇEK", "KREDİ KARTI", "DBS", "FATURA", "DÜZENLİ", "DİĞER"]
    r = blok_baslik(ws, r, "Ödeme Özeti (bekleyen)", 8)
    satirlar, renkler = [], []
    for ad, b, s, renk in araliklar:
        sec = [k for k in bekleyen if (b is None or k.vade >= b) and k.vade <= s]
        satir = [ad] + [sum(k.tutar or 0 for k in sec if k.kategori == c) for c in kategoriler]
        satir.append(sum(k.tutar or 0 for k in sec))
        satirlar.append(satir)
        renkler.append(renk)
    satirlar.append(["TOPLAM"] + [sum(s[i] for s in satirlar) for i in range(1, len(kategoriler) + 2)])
    renkler.append(GRI)
    r = tablo(ws, r, ["Dönem"] + kategoriler + ["TOPLAM"], satirlar, para=range(1, 9), renkler=renkler)
    ws.column_dimensions["I"].width = 16
    tahmini = [k for k in yaklasan if k.tutar is None]
    if tahmini:
        ws.cell(r - 1, 1, f"* Tutarı bilinmeyen {len(tahmini)} kalem toplamlara dahil değil "
                          "(ekstresi girilmemiş kart, tutarsız düzenli ödeme).").font = Font(italic=True, color="7F7F7F")
        r += 1

    r = blok_baslik(ws, r, "Acil — gecikmiş ve 7 gün içindekiler", 8)
    acil = [k for k in bekleyen if k.vade <= bugun + dt.timedelta(days=7)]
    if acil:
        r = tablo(ws, r, ["Ödeme", "Tür", "Kime", "Vade", "Kalan", "Tutar", "Banka", "Not"],
                  [[k.aciklama, k.baslik, k.kime, k.vade, kalan_metin((k.vade - bugun).days), k.tutar, k.banka,
                    k.not_] for k in acil],
                  para=(5,), renkler=[renk_kalan((k.vade - bugun).days, ay.acil) for k in acil])
    else:
        ws.cell(r, 1, "7 gün içinde bekleyen ödeme yok.")
        r += 2

    r = blok_baslik(ws, r, "Banka Hesapları", 8)
    satirlar, renkler = [], []
    for p in proj:
        satirlar.append([p["etiket"], p["guncel"], p["cikis_7"], p["cikis_ufuk"], p["min_bakiye"], p["min_tarih"],
                         p["durum"]])
        renkler.append(p["renk"])
    if satirlar:
        r = tablo(ws, r, ["Hesap", "Güncel Bakiye", "7 Gün Çıkış", f"{ay.ufuk} Gün Çıkış", "En Düşük Bakiye",
                          "En Düşük Tarih", "Durum"], satirlar, para=(1, 2, 3, 4), renkler=renkler)
    else:
        ws.cell(r, 1, "BANKALAR sekmesi boş.")
        r += 2

    r = blok_baslik(ws, r, "Personel", 8)
    if p_hat:
        r = tablo(ws, r, ["Tarih", "Personel", "Konu", "Kalan", "Detay"],
                  [[h[0], h[1], h[2], kalan_metin((h[0] - bugun).days), h[3]] for h in p_hat[:25]],
                  renkler=[YESIL if h[2] == "İZİNDE" else None for h in p_hat[:25]])
    else:
        ws.cell(r, 1, "Yaklaşan personel hatırlatması yok.")
        r += 2

    if uyarilar:
        r = blok_baslik(ws, r, "Veri Uyarıları", 8)
        for u in uyarilar:
            ws.cell(r, 1, u).font = Font(color="C00000")
            r += 1

    # ─────────────────────────────── YAKLAŞAN ÖDEMELER
    ws = wb.create_sheet("YAKLAŞAN ÖDEMELER")
    basl = ["Vade", "Gün", "Kalan", "İşlem Günü", "Kategori", "Ödeme", "Kime", "Tutar", "Banka", "Hesap",
            "Durum", "Not"]
    satirlar, renkler = [], []
    for k in yaklasan:
        kalan = (k.vade - bugun).days
        islem = tk.sonraki_banka_gunu(k.vade)
        satirlar.append([k.vade, GUNLER[k.vade.weekday()], kalan_metin(kalan), islem if islem != k.vade else None,
                         k.baslik, k.aciklama, k.kime, k.tutar, k.banka, k.hesap,
                         ("TAHMİNİ" if k.tahmini else k.durum) + (" · OTO" if k.otomatik else ""), k.not_])
        renkler.append(renk_kalan(kalan, ay.acil))
    sonraki = tablo(ws, 1, basl, satirlar, para=(7,), renkler=renkler)
    ws.cell(sonraki - 1, 6, "TOPLAM").font = Font(bold=True)
    t = ws.cell(sonraki - 1, 8, sum(k.tutar or 0 for k in yaklasan))
    t.font, t.number_format = Font(bold=True), PARA
    ws.cell(sonraki + 1, 1, "İşlem Günü: vade hafta sonu / resmî tatile denk geliyorsa bankanın işlem yapacağı ilk "
                            "iş günü.").font = Font(italic=True, color="7F7F7F")
    genislik(ws, (12, 11, 15, 12, 20, 32, 24, 15, 14, 14, 14, 44))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{max(len(satirlar) + 1, 2)}"

    # ─────────────────────────────── ÇEK TAKVİMİ
    cek_takvimi_yaz(wb.create_sheet("ÇEK TAKVİMİ"), kalemler, ay, tk, bugun)

    # ─────────────────────────────── BANKA DURUMU
    ws = wb.create_sheet("BANKA DURUMU")
    genislik(ws, (12, 14, 22, 34, 24, 15, 15, 16))
    r = 1
    for p in proj:
        r = blok_baslik(ws, r, f"{p['etiket']}   ·   Güncel bakiye {tl(p['guncel'])}   ·   KMH {tl(p['kmh'])}"
                        f"   ·   {ay.ufuk} günlük projeksiyon", 8)
        renkler = []
        for s in p["satirlar"]:
            renkler.append(KIRMIZI if s[6] < -p["kmh"] else TURUNCU if s[6] < 0 else None)
        r = tablo(ws, r, ["Tarih", "İşlem Günü", "Tür", "Açıklama", "Kime", "Giriş", "Çıkış", "Bakiye"],
                  [s[:5] + [s[5] if s[5] > 0 else None, -s[5] if s[5] < 0 else None, s[6]] for s in p["satirlar"]]
                  or [["", "", "", "Bekleyen hareket yok", "", None, None, p["guncel"]]],
                  para=(5, 6, 7), renkler=renkler or None)
    eslesmeyen = [k for k in bekleyen if not hesap_bul(hesaplar, k.banka, k.hesap)]
    if eslesmeyen:
        r = blok_baslik(ws, r, "Bankası tanımlı olmayan bekleyen ödemeler (bakiye kontrolü yapılamadı)", 8)
        r = tablo(ws, r, ["Vade", "Kategori", "Ödeme", "Kime", "Banka", "Tutar"],
                  [[k.vade, k.baslik, k.aciklama, k.kime, k.banka or "—", k.tutar] for k in eslesmeyen
                   if k.vade <= ufuk_son], para=(5,))

    # ─────────────────────────────── RAPORLAR
    raporlar_yaz(wb.create_sheet("RAPORLAR"), kalemler, hesaplar, islemler, dbs_limit, ay, bugun)

    # ─────────────────────────────── GİDER RAPORU
    gider_raporu_yaz(wb.create_sheet("GİDER RAPORU"), list(kalemler) + list(banka_giderleri),
                     gruplar or gider_gruplari_oku(None), ay, bugun)

    # ─────────────────────────────── PERSONEL İZİN
    ws = wb.create_sheet("PERSONEL İZİN")
    satirlar, renkler = [], []
    for p in personel:
        if not p.aktif(bugun):
            continue
        d = izin_durumu(p, bugun)
        izinde = next((i for i in p.izinler if i.bas <= bugun <= i.bit and i.onay != "RED"), None)
        satirlar.append([p.sicil, p.ad, p.gorev, p.bolum, p.giris, d["kidem"], yas_hesapla(p.dogum, bugun),
                         d["hak"], p.devreden, d["hak_edilen"], d["kullanilan"], d["planlanan"], d["kalan"],
                         d["sonraki_hak"], d["sonraki_hak_gun"],
                         f"İZİNDE ({izinde.tur}, dönüş {tr_tarih(izinde.donus)})" if izinde else ""])
        renkler.append(YESIL if izinde else (KIRMIZI if d["kalan"] is not None and d["kalan"] < 0 else None))
    r = tablo(ws, 1, ["Sicil", "Ad Soyad", "Görev", "Bölüm", "İşe Giriş", "Kıdem (yıl)", "Yaş", "Yıllık Hak (gün)",
                      "Devreden", "Hak Edilen", "Kullanılan", "Planlanan", "KALAN", "Sonraki Hak Ediş",
                      "Sonraki Hak (gün)", "Durum"], satirlar, renkler=renkler)
    ws.cell(r, 1, "Yıllık izin: 1–5 yıl 14 gün · 5–15 yıl 20 gün · 15 yıl+ 26 gün; 18 yaş altı ve 50 yaş üstü en az "
                  "20 gün (İş Kanunu m.53). Hak Edilen = Devir Tarihi'nden (yoksa işe girişten) sonraki hak ediş "
                  "yıldönümleri. Kullanılan/Planlanan yalnız onaylı ya da bekleyen YILLIK izinlerdir.").font = \
        Font(italic=True, color="7F7F7F")
    pasif = [p for p in personel if not p.aktif(bugun)]
    if pasif:
        r = blok_baslik(ws, r + 2, "Ayrılan personel", 8)
        tablo(ws, r, ["Sicil", "Ad Soyad", "Görev", "İşe Giriş", "Çıkış"],
              [[p.sicil, p.ad, p.gorev, p.giris, p.cikis] for p in pasif])
    genislik(ws, (8, 22, 16, 14, 12, 9, 6, 9, 9, 9, 10, 10, 9, 13, 10, 34))
    ws.freeze_panes = "C2"

    # ─────────────────────────────── İZİN TAKVİMİ
    ws = wb.create_sheet("İZİN TAKVİMİ")
    sirali = sorted(izinler, key=lambda i: i.bas)
    simdi = [i for i in sirali if i.bit >= bugun]
    gecmis = [i for i in sirali if i.bit < bugun][::-1]

    def iz_satir(i):
        return [i.ad, i.sicil, i.tur, i.bas, i.bit, i.gun, i.donus, i.onay, i.aciklama]

    def iz_renk(i):
        if i.onay == "RED":
            return GRI
        if i.bas <= bugun <= i.bit:
            return YESIL
        return SARI if i.onay == "BEKLİYOR" else None

    r = blok_baslik(ws, 1, "Devam eden ve gelecek izinler", 9)
    r = tablo(ws, r, ["Personel", "Sicil", "Tür", "Başlangıç", "Bitiş", "Gün", "İşe Dönüş", "Onay", "Açıklama"],
              [iz_satir(i) for i in simdi] or [["Kayıt yok"]], renkler=[iz_renk(i) for i in simdi] or None)
    r = blok_baslik(ws, r, "Geçmiş izinler", 9)
    tablo(ws, r, ["Personel", "Sicil", "Tür", "Başlangıç", "Bitiş", "Gün", "İşe Dönüş", "Onay", "Açıklama"],
          [iz_satir(i) for i in gecmis] or [["Kayıt yok"]])
    genislik(ws, (22, 8, 12, 12, 12, 7, 12, 12, 30))

    # ─────────────────────────────── TÜM KAYITLAR
    ws = wb.create_sheet("TÜM KAYITLAR")
    satirlar = [[k.vade, k.baslik, k.aciklama, k.kime, k.tutar, "TAHMİNİ" if k.tahmini else k.durum,
                 k.odeme_tarihi, k.banka, k.hesap, k.gider, k.yer, k.kaynak, k.not_] for k in kalemler]
    tablo(ws, 1, ["Vade", "Kategori", "Ödeme", "Kime", "Tutar", "Durum", "Ödeme Tarihi", "Banka", "Hesap",
                  "Gider Kalemi", "Masraf Yeri", "Kaynak Sekme", "Not"], satirlar, para=(4,),
          renkler=[GRI if k.durum in ("ÖDENDİ", "İPTAL") else None for k in kalemler])
    genislik(ws, (12, 20, 32, 24, 15, 12, 12, 14, 14, 30, 11, 18, 44))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:M{max(len(satirlar) + 1, 2)}"

    wb.save(yol)
    return {"yaklasan": yaklasan, "proj": proj, "p_hat": p_hat}


def banka_projeksiyonu(hesaplar: list[Hesap], bekleyen: list[Kalem], tk: Takvim, bugun: dt.date, son: dt.date):
    sonuc = []
    for h in hesaplar:
        hareket = [(tk.sonraki_banka_gunu(max(k.vade, bugun)), k.vade, k.baslik, k.aciklama
                    + (" (tahmini)" if k.tahmini else ""), k.kime, -(k.tutar or 0)) for k in bekleyen
                   if k.vade <= son and hesap_bul(hesaplar, k.banka, k.hesap) is h]
        hareket += [(tk.sonraki_banka_gunu(d), d, tur, acik, "", t) for d, acik, t, tur in h.hareketler if d <= son]
        hareket.sort(key=lambda x: (x[0], -x[5]))  # aynı gün önce girişler
        bakiye, min_b, min_t, satirlar = h.guncel, h.guncel, None, []
        c7 = cu = 0.0
        for islem, vade, tur, acik, kime, t in hareket:
            bakiye += t
            satirlar.append([vade, islem if islem != vade else None, tur, acik, kime, t, bakiye])
            if bakiye < min_b:
                min_b, min_t = bakiye, islem
            if t < 0:
                cu -= t
                if islem <= bugun + dt.timedelta(days=7):
                    c7 -= t
        if min_b < -h.kmh:
            durum, renk = f"YETERSİZ — {tl(-h.kmh - min_b)} eksik", KIRMIZI
        elif min_b < 0:
            durum, renk = "KMH kullanılacak", TURUNCU
        else:
            durum, renk = "Yeterli", YESIL
        if h.bakiye_tarihi and (bugun - h.bakiye_tarihi).days > 7:
            durum += f" · bakiye {tr_tarih(h.bakiye_tarihi)} tarihli"
        sonuc.append({"etiket": h.etiket, "guncel": h.guncel, "kmh": h.kmh, "cikis_7": c7, "cikis_ufuk": cu,
                      "min_bakiye": min_b, "min_tarih": min_t, "durum": durum, "renk": renk, "satirlar": satirlar})
    return sonuc


def cek_takvimi_yaz(ws, kalemler, ay: Ayarlar, tk: Takvim, bugun: dt.date):
    cekler = [k for k in kalemler if k.kategori == "ÇEK" and k.bekliyor]
    son = ay_ekle(bugun, ay.cek_ay)
    bankalar = sorted({k.banka or "—" for k in cekler}) or ["—"]
    n = len(bankalar)
    genislik(ws, (24, 14, 22, 26, 13, 14) + (14,) * max(n - 4, 0))
    r = 1
    ws.cell(r, 1, f"VERİLEN ÇEKLER TAKVİMİ · {tr_tarih(bugun)}").font = Font(bold=True, size=14, color="1F3864")
    r = 3
    gecikmis = [k for k in cekler if k.vade < bugun]

    # Aylık özet
    r = blok_baslik(ws, r, "Aylık özet (bekleyen çekler)", n + 3)
    aylar, d = [], dt.date(bugun.year, bugun.month, 1)
    while d <= son:
        aylar.append((d.year, d.month))
        d = ay_ekle(d, 1)
    satirlar, renkler = [], []
    if gecikmis:
        satirlar.append(["GECİKMİŞ"] + [sum(k.tutar or 0 for k in gecikmis if (k.banka or "—") == b) for b in bankalar]
                        + [sum(k.tutar or 0 for k in gecikmis), len(gecikmis)])
        renkler.append(KIRMIZI)
    for y, m in aylar:
        sec = [k for k in cekler if k.vade >= bugun and (k.vade.year, k.vade.month) == (y, m)]
        satirlar.append([ay_adi(y, m)] + [sum(k.tutar or 0 for k in sec if (k.banka or "—") == b) for b in bankalar]
                        + [sum(k.tutar or 0 for k in sec), len(sec)])
        renkler.append(None)
    sonra = [k for k in cekler if k.vade > gun_ay(son.year, son.month, 31)]
    if sonra:
        satirlar.append([f"{ay_adi(son.year, son.month)} sonrası"]
                        + [sum(k.tutar or 0 for k in sonra if (k.banka or "—") == b) for b in bankalar]
                        + [sum(k.tutar or 0 for k in sonra), len(sonra)])
        renkler.append(None)
    satirlar.append(["TOPLAM"] + [sum(s[i] for s in satirlar) for i in range(1, n + 3)])
    renkler.append(GRI)
    r = tablo(ws, r, ["Ay"] + bankalar + ["TOPLAM", "Adet"], satirlar, para=range(1, n + 2), renkler=renkler)

    # Haftalık takvim
    r = blok_baslik(ws, r, "Haftalık takvim (Pazartesi–Pazar)", n + 3)
    hafta = bugun - dt.timedelta(days=bugun.weekday())
    satirlar, renkler = [], []
    while hafta <= son:
        bit = hafta + dt.timedelta(days=6)
        sec = [k for k in cekler if hafta <= k.vade <= bit and k.vade >= bugun]
        if hafta == bugun - dt.timedelta(days=bugun.weekday()):
            sec += gecikmis
        satirlar.append([f"{hafta:%d.%m} – {bit:%d.%m.%Y}"]
                        + [sum(k.tutar or 0 for k in sec if (k.banka or "—") == b) or None for b in bankalar]
                        + [sum(k.tutar or 0 for k in sec) or None, len(sec) or None])
        renkler.append(TURUNCU if sec and hafta <= bugun else SARI if sec and (hafta - bugun).days < 14 else None)
        hafta += dt.timedelta(days=7)
    r = tablo(ws, r, ["Hafta"] + bankalar + ["TOPLAM", "Adet"], satirlar, para=range(1, n + 2), renkler=renkler)
    if gecikmis:
        ws.cell(r - 1, 1, "İlk haftaya gecikmiş çekler de eklendi.").font = Font(italic=True, color="7F7F7F")
        r += 1

    # Firma bazında
    r = blok_baslik(ws, r, "Lehdar (firma) bazında bekleyen çekler", n + 3)
    firma = defaultdict(lambda: [0.0, 0, None])
    for k in cekler:
        f = firma[k.kime or "—"]
        f[0] += k.tutar or 0
        f[1] += 1
        f[2] = min(f[2], k.vade) if f[2] else k.vade
    r = tablo(ws, r, ["Lehdar", "Toplam", "Adet", "İlk Vade"],
              [[a, v[0], v[1], v[2]] for a, v in sorted(firma.items(), key=lambda x: -x[1][0])], para=(1,))

    # Liste
    r = blok_baslik(ws, r, "Bekleyen çek listesi", n + 3)
    tablo(ws, r, ["Vade", "Çek No", "Banka / Hesap", "Lehdar", "Kalan", "Tutar", "İşlem Günü", "Açıklama"],
          [[k.vade, k.aciklama.replace("Çek ", ""), f"{k.banka} / {k.hesap}".strip(" /"), k.kime,
            kalan_metin((k.vade - bugun).days), k.tutar,
            tk.sonraki_banka_gunu(k.vade) if tk.sonraki_banka_gunu(k.vade) != k.vade else None, k.not_]
           for k in cekler],
          para=(5,), renkler=[renk_kalan((k.vade - bugun).days, ay.acil) for k in cekler])
    ws.column_dimensions["G"].width = max(ws.column_dimensions["G"].width or 0, 12)
    ws.column_dimensions["H"].width = max(ws.column_dimensions["H"].width or 0, 30)


def raporlar_yaz(ws, kalemler, hesaplar, islemler, dbs_limit, ay: Ayarlar, bugun: dt.date):
    kategoriler = ["ÇEK", "KREDİ KARTI", "DBS", "FATURA", "DÜZENLİ", "DİĞER"]
    genislik(ws, (22, 15, 15, 15, 15, 15, 15, 16))
    bas = ay_ekle(dt.date(bugun.year, bugun.month, 1), -ay.gecmis_ay)
    aylar, d = [], bas
    while d <= ay_ekle(bugun, ay.cek_ay):
        aylar.append((d.year, d.month))
        d = ay_ekle(d, 1)

    r = blok_baslik(ws, 1, "Aylık ÖDENEN (ödeme tarihine göre)", 8)
    satirlar = []
    for y, m in aylar:
        if (y, m) > (bugun.year, bugun.month):
            break
        sec = [k for k in kalemler if k.durum == "ÖDENDİ" and (k.odeme_tarihi or k.vade).year == y
               and (k.odeme_tarihi or k.vade).month == m]
        satirlar.append([ay_adi(y, m)] + [sum(k.tutar or 0 for k in sec if k.kategori == c) for c in kategoriler]
                        + [sum(k.tutar or 0 for k in sec)])
    r = tablo(ws, r, ["Ay"] + kategoriler + ["TOPLAM"], satirlar, para=range(1, 8))

    r = blok_baslik(ws, r, "Aylık BEKLEYEN (vadeye göre; tutarı bilinen)", 8)
    bekleyen = [k for k in kalemler if k.bekliyor]
    satirlar, renkler = [], []
    gec = [k for k in bekleyen if k.vade < bugun]
    if gec:
        satirlar.append(["GECİKMİŞ"] + [sum(k.tutar or 0 for k in gec if k.kategori == c) for c in kategoriler]
                        + [sum(k.tutar or 0 for k in gec)])
        renkler.append(KIRMIZI)
    for y, m in aylar:
        if (y, m) < (bugun.year, bugun.month):
            continue
        sec = [k for k in bekleyen if k.vade >= bugun and (k.vade.year, k.vade.month) == (y, m)]
        satirlar.append([ay_adi(y, m)] + [sum(k.tutar or 0 for k in sec if k.kategori == c) for c in kategoriler]
                        + [sum(k.tutar or 0 for k in sec)])
        renkler.append(None)
    r = tablo(ws, r, ["Ay"] + kategoriler + ["TOPLAM"], satirlar, para=range(1, 8), renkler=renkler)

    r = blok_baslik(ws, r, "DBS limit kullanımı", 8)
    dbs = defaultdict(float)
    for k in bekleyen:
        if k.kategori == "DBS":
            dbs[(up(k.kime), k.banka)] += k.tutar or 0
    anahtarlar = sorted(set(dbs) | set(dbs_limit))
    satirlar, renkler = [], []
    for a in anahtarlar:
        lim, kul = dbs_limit.get(a), dbs.get(a, 0.0)
        oran = kul / lim if lim else None
        satirlar.append([a[0], a[1], lim, kul, (lim - kul) if lim else None,
                         f"%{oran * 100:.0f}" if oran is not None else "limit yok"])
        renkler.append(KIRMIZI if oran and oran > 1 else TURUNCU if oran and oran > 0.8 else None)
    r = tablo(ws, r, ["Firma", "Banka", "Limit", "Kullanılan", "Kalan Limit", "Doluluk"],
              satirlar or [["Kayıt yok"]], para=(2, 3, 4), renkler=renkler or None)

    r = blok_baslik(ws, r, "Fatura türlerine göre (son 12 ay, ödenen + bekleyen)", 8)
    on_iki = ay_ekle(bugun, -12)
    ft = defaultdict(lambda: [0.0, 0])
    for k in kalemler:
        if k.kategori == "FATURA" and not k.tahmini and k.durum != "İPTAL" and k.vade >= on_iki:
            ft[(k.tur, k.aciklama)][0] += k.tutar or 0
            ft[(k.tur, k.aciklama)][1] += 1
    r = tablo(ws, r, ["Tür", "Abone / Yer", "Toplam", "Adet", "Ortalama"],
              [[a[0], a[1], v[0], v[1], v[0] / v[1] if v[1] else None] for a, v in sorted(ft.items())]
              or [["Kayıt yok"]], para=(2, 4))

    r = blok_baslik(ws, r, "Banka hareketleri — aylık giriş / çıkış", 8)
    hb = defaultdict(lambda: [0.0, 0.0])
    for d, banka, hesap, t, _a, _tur in islemler:
        hb[(d.year, d.month, banka + (" / " + hesap if hesap else ""))][0 if t > 0 else 1] += abs(t)
    tablo(ws, r, ["Ay", "Hesap", "Giriş", "Çıkış", "Net"],
          [[ay_adi(a[0], a[1]), a[2], v[0], v[1], v[0] - v[1]] for a, v in sorted(hb.items())] or [["Kayıt yok"]],
          para=(2, 3, 4))


def gider_gruplari_oku(wb) -> dict[str, tuple[str, str]]:
    """up(kalem) → (grup, kalem). Sekme boşsa varsayılan bilanço yapısı."""
    sonuc: dict[str, tuple[str, str]] = {}
    for r in (sekme_oku(wb, "GİDER KALEMLERİ") if wb is not None else []):
        g, k = up(r("Grup")), trim(r("Gider Kalemi"))
        if g and k:
            sonuc[up(k)] = (g, k)
    if not sonuc:
        sonuc = {up(k): (g, k) for g, ks in GIDER_VARSAYILAN for k in ks}
    return sonuc


def gider_raporu_yaz(ws, kalemler: list[Kalem], gruplar: dict, ay: Ayarlar, bugun: dt.date):
    """Bilanço düzeninde aylık gider tablosu: geçmiş aylar ödenen, bu ay ve sonrası ödenen + bekleyen."""
    sira: list[tuple[str, list[str]]] = []
    for g, k in gruplar.values():
        if not sira or sira[-1][0] != g:
            if g in [x[0] for x in sira]:
                next(x for x in sira if x[0] == g)[1].append(k)
                continue
            sira.append((g, []))
        sira[-1][1].append(k)
    if not any(g == "SINIFLANDIRILMAMIŞ ÖDEMELER" for g, _ in sira):
        sira.append(("SINIFLANDIRILMAMIŞ ÖDEMELER", []))
    ekstra = sorted({k.gider for k in kalemler if up(k.gider) not in gruplar})
    for g, ks in sira:
        if g == "SINIFLANDIRILMAMIŞ ÖDEMELER":
            ks.extend(x for x in ekstra if x not in ks)

    def kalem_adi(k: Kalem) -> str:
        return gruplar.get(up(k.gider), (None, k.gider))[1]

    def ay_anahtar(k: Kalem):
        d = (k.odeme_tarihi or k.vade) if k.durum == "ÖDENDİ" else k.vade
        if k.bekliyor and d < bugun:
            d = bugun  # gecikmiş bekleyen → bu aya
        return d.year, d.month

    gecerli = [k for k in kalemler if k.durum in ("ÖDENDİ", "BEKLİYOR", "KISMİ")]
    bas = ay_ekle(dt.date(bugun.year, bugun.month, 1), -min(ay.gecmis_ay, 6))
    aylar, d = [], bas
    while d <= ay_ekle(dt.date(bugun.year, bugun.month, 1), 2):
        aylar.append((d.year, d.month))
        d = ay_ekle(d, 1)
    tablo_ = defaultdict(float)
    for k in gecerli:
        tablo_[(kalem_adi(k), ay_anahtar(k))] += k.tutar or 0

    nc = len(aylar) + 2
    genislik(ws, (46,) + (14,) * (nc - 1))
    ws.cell(1, 1, f"{ay.firma + ' — ' if ay.firma else ''}GİDER RAPORU (bilanço kalemleri) · "
                  f"{tr_tarih(bugun)}").font = Font(bold=True, size=14, color="1F3864")
    ws.cell(2, 1, f"Geçmiş aylar: ödenen. {ay_adi(bugun.year, bugun.month)} ve sonrası: ödenen + bekleyen "
                  "(gecikmişler bu aya yazıldı). Tutarı bilinmeyen tahmini kalemler dahil değil.").font = \
        Font(italic=True, color="7F7F7F")
    basl = ["KALEM"] + [ay_adi(y, m) + (" *" if (y, m) >= (bugun.year, bugun.month) else "") for y, m in aylar] + \
           ["TOPLAM"]
    r = 4
    for i, b in enumerate(basl, 1):
        ws.cell(r, i, b)
    _baslik(ws, r, len(basl))
    r += 1

    def yaz(satir, etiket, degerler, dolgu=None, kalin=False):
        ws.cell(satir, 1, etiket)
        for i, v in enumerate(degerler + [sum(degerler)], 2):
            h = ws.cell(satir, i, v or None)
            h.number_format = PARA
        for c in range(1, nc + 1):
            h = ws.cell(satir, c)
            h.border = CERCEVE
            if dolgu:
                h.fill = dolgu
            if kalin:
                h.font = Font(bold=True)

    faaliyet = [0.0] * len(aylar)
    for g, ks in sira:
        grup_t = [sum(tablo_[(k, a)] for k in ks) for a in aylar]
        if not any(grup_t) and g not in FAALIYET_GRUPLARI and g != "SATIŞLARIN MALİYETİ":
            continue
        yaz(r, g, grup_t, ALT_BASLIK_FILL, True)
        r += 1
        for k in ks:
            t = [tablo_[(k, a)] for a in aylar]
            yaz(r, "   " + k, t)
            r += 1
        if g in FAALIYET_GRUPLARI:
            faaliyet = [x + y for x, y in zip(faaliyet, grup_t)]
    r += 1
    yaz(r, "TOPLAM FAALİYET GİDERLERİ", faaliyet, GRI, True)
    r += 1
    genel = [0.0] * len(aylar)
    for k in gecerli:
        a = ay_anahtar(k)
        if a in aylar:
            genel[aylar.index(a)] += k.tutar or 0
    yaz(r, "TÜM ÖDEMELER (mal alımı + kart + ortaklar dahil)", genel, GRI, True)
    r += 3

    # Bu ay mağaza / online ayrımı (ORTAK yarı yarıya)
    bu = (bugun.year, bugun.month)
    for baslik_ay in (ay_ekle(dt.date(*bu, 1), -1), dt.date(*bu, 1)):
        a = (baslik_ay.year, baslik_ay.month)
        r = blok_baslik(ws, r, f"MAĞAZA / ONLİNE MASRAF AYRIMI · {ay_adi(*a)}"
                        + (" (ödenen + bekleyen)" if a == bu else " (ödenen)"), 5)
        for i, b in enumerate(["MASRAF KALEMİ", "MAĞAZA", "ONLİNE", "TOPLAM", "% PAY"], 1):
            ws.cell(r, i, b)
        _baslik(ws, r, 5)
        r += 1
        ayrim = defaultdict(lambda: [0.0, 0.0])
        for k in gecerli:
            if ay_anahtar(k) != a or gruplar.get(up(k.gider), ("SINIFLANDIRILMAMIŞ ÖDEMELER",))[0] not in FAALIYET_GRUPLARI:
                continue
            t = k.tutar or 0
            if k.yer == "ONLİNE":
                ayrim[kalem_adi(k)][1] += t
            elif k.yer == "ORTAK":
                ayrim[kalem_adi(k)][0] += t / 2
                ayrim[kalem_adi(k)][1] += t / 2
            else:
                ayrim[kalem_adi(k)][0] += t
        toplam = sum(sum(v) for v in ayrim.values())
        for g, ks in sira:
            if g not in FAALIYET_GRUPLARI:
                continue
            gm = sum(ayrim[k][0] for k in ks if k in ayrim)
            go = sum(ayrim[k][1] for k in ks if k in ayrim)
            satirlar = [(g, gm, go, True)] + [("   " + k, ayrim[k][0], ayrim[k][1], False) for k in ks if k in ayrim]
            if not gm and not go:
                continue
            for et, m_, o_, kalin in satirlar:
                for i, v in enumerate([et, m_ or None, o_ or None, (m_ + o_) or None,
                                       (m_ + o_) / toplam if toplam else None], 1):
                    h = ws.cell(r, i, v)
                    h.border = CERCEVE
                    h.number_format = "0.0%" if i == 5 else PARA
                    if kalin:
                        h.font, h.fill = Font(bold=True), ALT_BASLIK_FILL
                r += 1
        tm = sum(v[0] for v in ayrim.values())
        to = sum(v[1] for v in ayrim.values())
        for i, v in enumerate(["TOPLAM FAALİYET GİDERLERİ", tm, to, tm + to, 1 if toplam else None], 1):
            h = ws.cell(r, i, v)
            h.border, h.font, h.fill = CERCEVE, Font(bold=True), GRI
            h.number_format = "0.0%" if i == 5 else PARA
        r += 3
    ws.cell(r, 1, "ORTAK masraf yeri mağaza ve online'a yarı yarıya bölünür. Kredi kartı ödemeleri "
                  "'dağıtılmamış' görünür; kalem dağıtımı kredi kartı masraf dağıtım tablosunda yapılır.").font = \
        Font(italic=True, color="7F7F7F")
    ws.freeze_panes = "B5"


def personel_hatirlatmalari(personel, izinler, ay: Ayarlar, bugun: dt.date):
    """(tarih, personel, konu, detay) listesi."""
    son = bugun + dt.timedelta(days=ay.ufuk)
    h = []
    for i in izinler:
        if i.onay == "RED":
            continue
        if i.bas <= bugun <= i.bit:
            h.append((bugun, i.ad, "İZİNDE", f"{i.tur} · {tr_tarih(i.bas)}–{tr_tarih(i.bit)} · dönüş {tr_tarih(i.donus)}"))
        elif bugun < i.bas <= son:
            h.append((i.bas, i.ad, "İZİN BAŞLIYOR",
                      f"{i.tur} · {i.gun:g} gün · dönüş {tr_tarih(i.donus)}" + (" · ONAY BEKLİYOR" if i.onay == "BEKLİYOR" else "")))
        if bugun <= i.donus <= son and i.bas <= bugun:
            h.append((i.donus, i.ad, "İŞE DÖNÜŞ", f"{i.tur} izni bitti"))
    for p in personel:
        if not p.aktif(bugun):
            continue
        d = izin_durumu(p, bugun)
        if d["sonraki_hak"] and d["sonraki_hak"] <= son:
            h.append((d["sonraki_hak"], p.ad, "İZİN HAK EDİŞ",
                      f"{d['kidem'] + 1}. yıl · +{d['sonraki_hak_gun']} gün yıllık izin"))
        if ay.dogum_gunu and p.dogum:
            dg = yil_ekle(p.dogum, bugun.year - p.dogum.year)
            if dg < bugun:
                dg = yil_ekle(p.dogum, bugun.year + 1 - p.dogum.year)
            if (dg - bugun).days <= ay.dogum_gunu:
                h.append((dg, p.ad, "DOĞUM GÜNÜ", f"{dg.year - p.dogum.year} yaş"))
        if p.giris and bugun <= ay_ekle(p.giris, 2) <= son:
            h.append((ay_ekle(p.giris, 2), p.ad, "DENEME SÜRESİ BİTİYOR", f"İşe giriş {tr_tarih(p.giris)}"))
        if p.cikis and bugun <= p.cikis <= son:
            h.append((p.cikis, p.ad, "İŞTEN ÇIKIŞ", ""))
    h.sort(key=lambda x: (x[0], x[1]))
    return h


# ═══════════════════════════════════════════════════════════════════════════
#  TAKVİM DOSYASI (.ics)
# ═══════════════════════════════════════════════════════════════════════════
def ics_kacis(s: str) -> str:
    return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def ics_satir(s: str) -> str:
    """RFC 5545: 75 bayttan uzun satırları katla."""
    b = s.encode("utf-8")
    if len(b) <= 75:
        return s
    parcalar, simdiki = [], ""
    for ch in s:
        if len((simdiki + ch).encode("utf-8")) > (75 if not parcalar else 74):
            parcalar.append(simdiki)
            simdiki = ""
        simdiki += ch
    parcalar.append(simdiki)
    return "\r\n ".join(parcalar)


def ics_yaz(yol: Path, kalemler: list[Kalem], p_hat, ay: Ayarlar, bugun: dt.date):
    simdi = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    satirlar = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Odabaslar//Hatirlatici//TR", "CALSCALE:GREGORIAN",
                "METHOD:PUBLISH", "X-WR-CALNAME:Ödeme Hatırlatıcı", "X-WR-TIMEZONE:Europe/Istanbul"]

    def olay(gun: dt.date, ozet: str, aciklama: str, uid_kaynak: str, alarm: bool):
        uid = hashlib.sha1(uid_kaynak.encode("utf-8")).hexdigest()[:20] + "@hatirlatici"
        satirlar.extend(["BEGIN:VEVENT", f"UID:{uid}", f"DTSTAMP:{simdi}",
                         f"DTSTART;VALUE=DATE:{gun:%Y%m%d}",
                         f"DTEND;VALUE=DATE:{gun + dt.timedelta(days=1):%Y%m%d}",
                         f"SUMMARY:{ics_kacis(ozet)}", f"DESCRIPTION:{ics_kacis(aciklama)}", "TRANSP:TRANSPARENT"])
        if alarm:
            for n in ay.alarmlar:
                # tüm gün etkinliği 00:00'da başlar; n gün önce saat 09:00 → -P(n-1)DT15H
                tetik = (f"-P{n - 1}DT15H" if n > 1 else "-PT15H") if n >= 1 else "PT9H"
                satirlar.extend(["BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{ics_kacis(ozet)}",
                                 f"TRIGGER:{tetik}", "END:VALARM"])
            satirlar.extend(["BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{ics_kacis(ozet)}",
                             "TRIGGER:PT9H", "END:VALARM"])
        satirlar.append("END:VEVENT")

    son = bugun + dt.timedelta(days=max(ay.ufuk, 31 * ay.cek_ay))
    for k in kalemler:
        if not k.bekliyor or k.vade > son:
            continue
        ozet = f"💳 {k.aciklama}" if k.kategori == "KREDİ KARTI" else f"{k.baslik}: {k.aciklama}"
        if k.tutar is not None:
            ozet += f" — {tl(k.tutar)}"
        if k.vade < bugun:
            ozet = "GECİKMİŞ! " + ozet
        acik = "\n".join(x for x in (f"Kime: {k.kime}" if k.kime else "", f"Banka: {k.banka} {k.hesap}".strip()
                                     if k.banka else "", f"Vade: {tr_tarih(k.vade)}", k.not_) if x)
        olay(max(k.vade, bugun), ozet, acik, f"{k.kaynak}|{k.anahtar}|{k.vade}|{k.aciklama}", True)
    for gun, ad, konu, detay in p_hat:
        if konu == "İZİNDE":
            continue
        olay(gun, f"👤 {konu}: {ad}", detay, f"P|{ad}|{konu}|{gun}", konu in ("İZİN HAK EDİŞ", "DENEME SÜRESİ BİTİYOR"))
    satirlar.append("END:VCALENDAR")
    yol.write_text("\r\n".join(ics_satir(s) for s in satirlar) + "\r\n", encoding="utf-8")


# ═══════════════════════════════════════════════════════════════════════════
#  ÇALIŞTIR
# ═══════════════════════════════════════════════════════════════════════════
def calistir(girdi: Path, cikti: Path | None, ics: Path | None, bugun: dt.date):
    wb = openpyxl.load_workbook(girdi, data_only=True)
    ay = ayar_oku(wb)
    tk = takvim_oku(wb, ay)
    kalemler, uyarilar, dbs_limit = kalemleri_topla(wb, bugun, ay)
    hesaplar, islemler, banka_giderleri = bankalari_oku(wb, bugun, uyarilar)
    gruplar = gider_gruplari_oku(wb)
    bilinmeyen = sorted({k.gider for k in kalemler + banka_giderleri if up(k.gider) not in gruplar})
    for g in bilinmeyen:
        uyarilar.append(f"Gider kalemi '{g}' GİDER KALEMLERİ listesinde yok — SINIFLANDIRILMAMIŞ sayıldı.")
    personel, izinler = personel_oku(wb, tk, uyarilar)
    if max(tk.tatil, default=bugun) < bugun + dt.timedelta(days=ay.ufuk):
        uyarilar.append("TATİLLER sekmesinde ileri tarihli tatil yok — gelecek yılın tatillerini ekleyin.")

    cikti = cikti or girdi.with_name(f"HATIRLATICI_RAPOR_{bugun:%Y%m%d}.xlsx")
    ozet = rapor_yaz(cikti, kalemler, hesaplar, islemler, personel, izinler, dbs_limit, uyarilar, ay, tk, bugun,
                     banka_giderleri, gruplar)
    ics = ics or girdi.with_name("HATIRLATICI.ics")
    ics_yaz(ics, kalemler, ozet["p_hat"], ay, bugun)

    # Konsol özeti
    print(f"\nHATIRLATICI · {tr_tarih(bugun)} {GUNLER[bugun.weekday()]}")
    print("─" * 60)
    yak = ozet["yaklasan"]
    gec = [k for k in yak if k.vade < bugun]
    hafta = [k for k in yak if bugun <= k.vade <= bugun + dt.timedelta(days=7)]
    if gec:
        print(f"⚠ GECİKMİŞ: {len(gec)} ödeme, {tl(sum(k.tutar or 0 for k in gec))}")
        for k in gec:
            print(f"   {tr_tarih(k.vade)}  {k.baslik[:26]:<26} {k.aciklama[:30]:<30} {tl(k.tutar):>16}")
    print(f"7 gün içinde: {len(hafta)} ödeme, {tl(sum(k.tutar or 0 for k in hafta))}")
    for k in hafta:
        print(f"   {tr_tarih(k.vade)}  {k.baslik[:26]:<26} {k.aciklama[:30]:<30} {tl(k.tutar):>16}")
    print(f"{ay.ufuk} gün içinde toplam: {len(yak)} ödeme, {tl(sum(k.tutar or 0 for k in yak))}")
    for p in ozet["proj"]:
        if p["min_bakiye"] < 0:
            print(f"⚠ {p['etiket']}: {p['durum']} (en düşük {tl(p['min_bakiye'])}, {tr_tarih(p['min_tarih'])})")
    for gun, ad, konu, detay in ozet["p_hat"][:10]:
        print(f"   👤 {tr_tarih(gun)}  {konu}: {ad} {('· ' + detay) if detay else ''}")
    if uyarilar:
        print(f"{len(uyarilar)} veri uyarısı var — ÖZET sekmesine bakın.")
    print("─" * 60)
    print(f"Rapor : {cikti}")
    print(f"Takvim: {ics}  (telefona / Outlook'a / Google Takvim'e içe aktarın)\n")


def main():
    p = argparse.ArgumentParser(description="Ödeme, çek, banka ve personel hatırlatıcısı")
    p.add_argument("girdi", nargs="?", type=Path, default=Path(__file__).with_name("HATIRLATICI.xlsx"))
    p.add_argument("-o", "--cikti", type=Path, help="Rapor dosyası (varsayılan HATIRLATICI_RAPOR_<tarih>.xlsx)")
    p.add_argument("-i", "--ics", type=Path, help="Takvim dosyası (varsayılan HATIRLATICI.ics)")
    p.add_argument("-t", "--tarih", help="Rapor tarihi (GG.AA.YYYY). Boş = bugün.")
    p.add_argument("--sablon", action="store_true", help="Boş HATIRLATICI.xlsx şablonunu yazar")
    p.add_argument("--ornek", action="store_true", help="Örnek verili ORNEK_HATIRLATICI.xlsx yazar")
    a = p.parse_args()
    bugun = tarih(a.tarih) if a.tarih else dt.date.today()
    if a.tarih and not bugun:
        sys.exit(f"Tarih okunamadı: {a.tarih}  (GG.AA.YYYY yazın)")

    if a.sablon or a.ornek:
        yol = a.girdi if a.girdi.name != "HATIRLATICI.xlsx" or a.sablon else a.girdi.with_name("ORNEK_HATIRLATICI.xlsx")
        if a.ornek and a.girdi.name == "HATIRLATICI.xlsx":
            yol = a.girdi.with_name("ORNEK_HATIRLATICI.xlsx")
        if a.sablon and yol.exists():
            sys.exit(f"{yol} zaten var. Üzerine yazmamak için durdum; önce silin ya da başka ad verin.")
        sablon_yaz(yol, ornek=a.ornek, bugun=bugun)
        print(f"Yazıldı: {yol}")
        return
    if not a.girdi.exists():
        sablon_yaz(a.girdi)
        print(f"{a.girdi} bulunamadı; boş şablon oluşturuldu. Verileri girip tekrar çalıştırın.")
        return
    calistir(a.girdi, a.cikti, a.ics, bugun)


if __name__ == "__main__":
    main()
