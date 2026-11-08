# İndirim ve Mağaza Fiyat — Python sürümü

> Ödeme / çek / banka / personel hatırlatıcısı için: [hatirlatici/BENIOKU.md](hatirlatici/BENIOKU.md)

`INDIRIM_VE_MAGAZA_FIYAT` Excel'indeki ~710.000 formülün yerine geçen tek dosyalık program.
Aynı ham verilerle Excel'in ürettiği sonuçların aynısını üretir (06.11.2026 verisiyle 1569 ürünün
oranı ve fiyatı tek tek karşılaştırıldı). Hesap birkaç saniye sürer.

## Kurulum (bir kez)

```bash
pip install openpyxl
```

## Kullanım

1. Nebim raporlarını `GIRDI_SABLON.xlsx` dosyasındaki sekmelere A1'den yapıştırın.
   Eski Excel dosyasını da doğrudan girdi olarak verebilirsiniz.
2. Programı çalıştırın:

```bash
python fiyatla.py GIRDI_SABLON.xlsx
```

3. Aynı klasörde `CIKTI_YYYYAAGG.xlsx` oluşur.

Seçenekler: `-o sonuc.xlsx` (çıktı adı), `-t 06.11.2026` (hesap tarihi), `-a baska_ayar.xlsx`.

## Girdi sekmeleri

| Sekme | İçerik |
|---|---|
| HAM VERİ KAMPANYA LİSTESİ | Nebim kampanya listesi |
| HAM VERİ ÜRÜN 1 | Spor reyonları, stoğu olan ürünler |
| HAM VERİ ÜRÜN 1B | Stoğu biten ama yakın zamanda satılan ürünler |
| HAM VERİ ÜRÜN 2 | Diğer reyonlardan kampanyalı ürünler |
| EK ÜRÜNLER | Sonradan gelen ürünler |
| HAM VERİ — Depo | Depo dağılım raporu |
| MANUEL ORANLAR | (isteğe bağlı) Ürün Kodu + BENİM ORANIM. Eski dosyadaki FİYAT — … sekmeleri de okunur. |

Sütunlar başlık adına göre bulunur, sıraları değişse de çalışır. Aynı ürün birden fazla sekmede
varsa ilk görüldüğü sekme geçerlidir.

## Ayarlar — `ayarlar.xlsx`

* **Genel**: KDV, gider, maksimum indirim, yuvarlama, kapsam reyonları, depo tavanı, stoksuz kuralı…
* **Kurallar**: eski ③ OTOMATİK KURALLAR tablosu. Ürüne uyan ilk aktif kural geçerlidir. 14 satır sınırı yok.
* **Kampanyalar**: Nebim kampanya adı → oran / tip (TEKLİ, TEKFİYAT, BİTTİ, ÇOKLU, ÇİFTLİ).

Dosya silinirse program ilk çalıştığında güncel varsayılanlarla yeniden oluşturur.

## Çıktı sekmeleri

| Sekme | Ne işe yarar |
|---|---|
| ÖZET | Adetler, uyarılar, tanımsız kampanya adları |
| Nebim Aktarım | M2 fiyat listesi (Nebim içe aktarım formatı) |
| NEBİM %50 … %15 | Oran kampanya listeleri (ItemCode / ItemDescription) |
| NEBİM TEK FİYAT … | Tek fiyat kampanya listeleri |
| NE DEĞİŞECEK | Nebim'deki mevcut fiyattan farklılaşan ürünler |
| HATALAR | Kampanya listesinde silinecek / teyit edilecek satırlar |
| ÜRÜN DETAY | Her ürünün tüm hesabı: kapsam, depo, kural, oran, fiyat, kâr. Eski "STOK KODU KONTROL" yerine filtreleyin. |

## Oran nasıl belirlenir (öncelik sırası)

1. **BENİM ORANIM** (elle oran) varsa o uygulanır.
2. **Tek Fiyat** kampanyasındaysa sabit fiyat uygulanır.
3. **Çoklu kampanyada** ya da **stoksuz** ise indirim uygulanmaz.
4. Temel oran şu sırayla belirlenir: kampanya listesindeki oran → ürün kartındaki oran → 0.
   Kapsam dışı reyonlarda yalnız kampanya oranı geçer.
5. **Kurallar** tablosunda ürüne uyan ilk kural uygulanır:
   * *Listeyi de ez*: kuralın oranı geçer.
   * *Yalnız listede yoksa*: kural yalnız kampanya ve kart oranı yoksa devreye girer.
   * *Daha yüksekse*: kural oranı ile temel orandan büyük olanı geçer.
6. **Depo tavanı**: depodaki üründe oran %25'i aşıyorsa %25'e çekilir. Hariç marka, reyon ve sezonlara uygulanmaz.
7. Fiyat = etiket × (1 − oran), en yakın …9,99'a yuvarlanır. Etiketin %50'sinin altına inmez.

## Excel'e göre sadeleştirilenler

Kapalı olan ya da kullanılmayan özellikler çıkarıldı:

* enflasyon düzeltmesi
* sezon matrisi
* marka maliyet tabanı
* ORAN ATAMA panosu
* "Aşağı" yuvarlama yönü

Kâr/marj hesabı bilgi amaçlı olarak duruyor. Fiyatı değiştirmez.

* **Outlet işareti**: Excel'de formül her ürün için listenin 1. satırına baktığı için çalışmıyordu.
  Burada ürünün kendi "Noos_Outlet" alanında "Outlet" yazıyorsa outlet sayılır.
  Şu anki veride outlet ürün yok, sonuç değişmiyor.
* **Depo sayımı**: Excel'deki SUMIF, "0004" gibi sıfırla başlayan kodları stok adediyle karıştırabiliyordu.
  Burada yalnız ürün kodu sütunu eşleşir. Şu anki veride etkilenen tek ürün 0004.
  Tek fiyatlı olduğu için fiyatı değişmiyor.
