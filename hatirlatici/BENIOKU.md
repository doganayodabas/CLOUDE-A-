# HATIRLATICI — Ödeme, Çek, Banka ve Personel Takibi

Verilen çekler, kredi kartları, DBS, elektrik ve diğer faturalar, düzenli ödemeler (kira, SGK, vergi,
kredi taksiti, maaş), banka bakiyeleri, personel ve izinler tek bir Excel dosyasından takip edilir.
Program bu dosyayı okur ve şunları üretir:

* **Rapor** `HATIRLATICI_RAPOR_YYYYAAGG.xlsx`: tarihi yaklaşan ve gecikmiş ödemeler, çek takvimi,
  banka bakiyesinin yetip yetmediği, bilanço düzeninde gider raporu, personel izin durumu.
* **Takvim** `HATIRLATICI.ics`: telefona, Outlook'a ya da Google Takvim'e alınır.
  Her ödeme için 3 gün önce, 1 gün önce ve aynı gün sabah 09:00'da alarm çalar.

## Kullanım

```bash
pip install openpyxl
```

```bash
python hatirlatici.py
```

`HATIRLATICI.xlsx` yoksa program boş şablonu oluşturur. Verileri girip tekrar çalıştırın.

Denemek için `ORNEK_HATIRLATICI.xlsx` örnek verilerle hazırdır:

```bash
python hatirlatici.py ORNEK_HATIRLATICI.xlsx
```

Seçenekler: `-t 15.11.2026` (rapor tarihi), `-o rapor.xlsx`, `-i takvim.ics`,
`--sablon` (boş şablon yazar), `--ornek` (bugüne göre örnek dosyayı yeniden yazar).

## Veri sekmeleri (`HATIRLATICI.xlsx`)

| Sekme | Ne girilir |
|---|---|
| AYARLAR | Firma adı, kaç günlük ödemeler listelensin (30), acil uyarı günü (3), takvim alarmları, hafta tatili |
| BANKALAR | Banka, hesap adı, IBAN, **bakiye + bakiye tarihi**, KMH limiti |
| BANKA İŞLEMLERİ | İsteğe bağlı hareket defteri. İleri tarihli GİRİŞ (POS blokesi, tahsilat) projeksiyona eklenir |
| VERİLEN ÇEKLER | Çek no, banka/hesap, lehdar, keşide, vade, tutar, durum |
| KREDİ KARTLARI | Kart tanımı: hesap kesim ve son ödeme günü, otomatik ödeme, ödeme hesabı |
| KART EKSTRELERİ | Gelen ekstre: son ödeme tarihi, ekstre / asgari / ödenen tutar |
| DBS | Firma, banka, DBS limiti, fatura no, vade, tutar |
| FATURALAR | Elektrik, su, doğalgaz, internet, telefon… abone no, yer, son ödeme, tutar |
| DÜZENLİ ÖDEMELER | Kira, SGK, KDV, muhtasar, geçici vergi, kredi taksiti, maaş… her ayın X günü (aylık / 3 aylık / yıllık) |
| DİĞER ÖDEMELER | Tek seferlik, tarihi yaklaşan ödemeler |
| PERSONEL | Sicil, ad, görev, bölüm, doğum tarihi, işe giriş, çıkış, devreden izin |
| İZİNLER | Personel, izin türü, başlangıç, bitiş, gün (boşsa hesaplanır), onay |
| TATİLLER | 2026–2027 resmî tatilleri hazır. Dini bayramları her yıl Diyanet takviminden kontrol edin |
| GİDER KALEMLERİ | Bilanço gider yapısı (aşağıda). Diğer sekmelerdeki *Gider Kalemi* açılır listesi buradan gelir |

Sekmelerin ilk hücresinde açıklama notu vardır. Açılır listeler var ama listede olmayan değer de yazılabilir.

**Ödendi işaretleme:** Durum = `ÖDENDİ` yazın ya da Ödeme Tarihi girin. Kartta *Ödenen Tutar ≥ Ekstre
Tutarı* ise ödendi sayılır. Kısmi ödemede kalan tutar hatırlatılır.

**Tahmini hatırlatmalar:** Kart ve düzenli ödemeler için program, bugünden ileriye her dönem bir
hatırlatma üretir. O ayın gerçek ekstresi (KART EKSTRELERİ) ya da faturası (FATURALAR'da aynı Tür +
Abone No) girilince tahmini satırın yerine o geçer.

## Rapor sekmeleri

| Sekme | İçerik |
|---|---|
| ÖZET | Gecikmiş / bugün / 1–3 gün / 4–7 gün / 8–30 gün toplamları (kategori bazında), acil liste, banka yeterliliği, personel hatırlatmaları, veri uyarıları |
| YAKLAŞAN ÖDEMELER | Tüm bekleyen ödemeler. Renkler: kırmızı = gecikmiş, turuncu = acil, sarı = 7 gün içinde. Vade hafta sonu ya da tatile denk gelirse bankanın işlem günü gösterilir |
| ÇEK TAKVİMİ | Aylık özet ve haftalık takvim (banka bazında), lehdar bazında toplamlar, bekleyen çek listesi |
| BANKA DURUMU | Her hesap için güncel bakiyeden başlayıp 30 gün içindeki çıkışları (çek, DBS, kart, fatura, düzenli) ve beklenen girişleri tarih sırasıyla düşer. Bakiye eksiye düşerse turuncu, KMH'yi aşarsa kırmızı ve **"YETERSİZ — X ₺ eksik"** |
| RAPORLAR | Aylık ödenen ve bekleyen (kategori bazında), DBS limit doluluğu, fatura türü ortalamaları, aylık banka giriş/çıkış |
| GİDER RAPORU | Bilanço düzeninde aylık gider tablosu ve Mağaza / Online masraf ayrımı |
| PERSONEL İZİN | Kıdem, yaş, yıllık izin hakkı, hak edilen / kullanılan / planlanan / **kalan izin**, sonraki hak ediş tarihi, şu an izinde mi |
| İZİN TAKVİMİ | Devam eden, gelecek ve geçmiş izinler ile işe dönüş günü |
| TÜM KAYITLAR | Filtrelenebilir tam liste |

## Gider kalemleri (ODABAŞLAR bilançosu)

Her ödemeye bir **Gider Kalemi** ve **Masraf Yeri** (MAĞAZA / ONLİNE / ORTAK) yazılabilir.
ORTAK, mağaza ile online arasında yarı yarıya bölünür. Boş bırakılırsa program türüne göre tahmin eder:

| Ödeme | Tahmin edilen kalem |
|---|---|
| Çek, DBS | Brüt mal alımı (ürün ödemeleri) |
| Kredi kartı | Kredi Kartı Ödemesi (dağıtılmamış) — reyon dağıtımı kart masraf tablosunda yapılır |
| Elektrik | Mağaza Elektrik. Açıklamada "DEPO" geçiyorsa Depo Elektrik, "EV" geçiyorsa Ev Elektrik Gideri |
| İnternet / Telefon | İnternet – Mağaza ve Depo / Muhasebe Telefon Hattı ("ONLİNE" geçiyorsa Online Telefon Hattı) |
| SGK, Maaş | SGK Primi, Personel Maaşları |
| KDV, Muhtasar, Geçici Vergi | Vergi Ödemeleri |
| Kira, Sigorta, Muhasebe | Depo Kirası / Giderleri, Sigorta, Muhasebe Gideri |
| Kredi taksiti | Kredi Taksiti / Faiz |

Gruplar: Satışların Maliyeti · Personel · Vergi ve Harçlar · Kira ve İşletme · Satış ve Pazarlama ·
Genel Yönetim · Finansman · Sınıflandırılmamış · Ortaklara Yapılan Ödemeler.
Listeyi GİDER KALEMLERİ sekmesinden değiştirebilirsiniz.

POS komisyonu, havale/EFT ücreti gibi yalnız bankadan çıkan giderleri BANKA İŞLEMLERİ'ne
*Gider Kalemi* ile yazın; rapora girer. Başka sekmede kayıtlı bir ödemeyi BANKA İŞLEMLERİ'ne yazarken
Gider Kalemi'ni boş bırakın. Yoksa gider iki kez sayılır.

## Yıllık izin hesabı

İş Kanunu m.53'e göre:

* 1–5 yıl: 14 gün
* 5–15 yıl: 20 gün
* 15 yıl ve üzeri: 26 gün
* 18 yaş ve altı ile 50 yaş ve üstü: en az 20 gün

Hak, her işe giriş yıldönümünde kazanılır.

**Devreden İzin + Devir Tarihi** girerseniz, o tarihe kadarki hesap yerine bu sayı kullanılır.
Sonraki hak edişler ve izinler bunun üzerine eklenir.

İzin günü boşsa şöyle sayılır:

* **Yıllık ve mazeret izni:** hafta tatili ve resmî tatiller düşülür. Arife yarım gün sayılır.
* **Rapor, ücretsiz ve diğerleri:** takvim günü sayılır.

Personel hatırlatmaları:

* izin başlangıcı
* işe dönüş
* izin hak ediş günü
* doğum günü
* deneme süresinin bitişi (2 ay)
* işten çıkış
