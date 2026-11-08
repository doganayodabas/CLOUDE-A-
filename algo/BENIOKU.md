# Trend + Scalp Hibrit Sistem (5 dk, çift yönlü)

İki parçadan oluşur:

| Dosya | Ne işe yarar |
|---|---|
| `idealdata/TrendScalp_5dk.txt` | IdealData Sistem Editörü'ne yapıştırılacak sistem kodu |
| `backtest/motor.py` | Aynı mantığın Python motoru (veri okuma, indikatörler, işlem simülasyonu, istatistik) |
| `backtest/sistem.py` | Komut satırı: test, walk-forward optimizasyon, HTML rapor, işlem listesi |

İki kod aynı veride **bar bar aynı pozisyonu** üretecek şekilde yazıldı. IdealData kodu bir C#
test düzeneğinde derlenip Python motoruyla (`DOLUM=kapanis`) karşılaştırıldı. 3 mod, gün sonu
kapatma açık ve kapalı, trendli ve rastgele verilerle denendi; fark yalnızca verinin son
barında (Python açık pozisyonu veri sonunda kapatıyor).

## İstenen özellikler nasıl karşılanıyor

| İstek | Uygulama |
|---|---|
| Scalp + trend takip | `MOD=1` scalp (TP1'de tam çıkış), `MOD=2` saf trend (izleyen stop), `MOD=3` hibrit: TP1 görülünce stop kâra çekilir, sonra izleyen stop ile trend sürülür |
| Çift yönlü | Long ve short simetrik |
| Düşük MaxDD | ATR'ye göre ilk stop, kâr kilidi, %3 / ATR izleyen stop (hangisi yakınsa), günlük zarar limiti, zaman stopu |
| Kârlı işlem oranı yüksek | TP1 (scalp hedefi) yakın tutulur; hibritte TP1'den sonra işlem zarara dönemez |
| İğne barlarda kâr alma | Girişte `IGNE_TP_ATR` uzaklığına limit kâr al seviyesi; lehimize uzun fitil bırakan barda çıkış; iğne barda giriş yapılmaz |
| Gaplı açılışta koruma | Gün sonu pozisyon kapatma (varsayılan); taşınıyorsa aleyhe büyük gapte açılışta çık; açılışta ve gapli günde bekleme; stop gap ile atlanırsa dolum açılış fiyatından hesaplanır |
| Yatayda işlemsiz | ADX ≥ `ADX_MIN` **ve** Kaufman Verimlilik Oranı ≥ `ER_MIN` değilse yeni işlem açılmaz |

## IdealData'da kullanım

1. Sistem Editörü'nde yeni sistem açın, `idealdata/TrendScalp_5dk.txt` içeriğini yapıştırın.
2. Üstteki **PARAMETRELER** bölümünü Python optimizasyonunda bulduğunuz değerlerle doldurun.
3. 5 dakikalık grafiğe ekleyin.
4. `SON_GIRIS` / `GUN_SONU` saatlerini enstrümanın seansına göre ayarlayın (SSDD biçiminde, ör. 1755).
   VİOP akşam seansında işlem yapacaksanız bu saatleri ona göre değiştirin.

> Kod, iDeal "Sistem Modülü" belgesindeki yazım biçimini kullanır: `Sistem.GrafikVerileri`,
> `Sistem.GrafikFiyatOku(V, "Kapanis")`, `V.Count`, `Sistem.Liste(0)`, `Sistem.Cizgiler[k].Deger`.
> Bunlara ek olarak sinyal için `Sistem.Yon[i]` ("A" / "S" / "F") ve barın zamanı için `V[i].Date`
> kullanılır. EMA/ATR/ADX/ER hazır fonksiyonlara bağlı kalmamak için kod içinde hesaplanır.
> Grafikte 3 çizgi çizilir: hızlı EMA, yavaş EMA ve pozisyondayken aktif stop seviyesi.
> `V[i].Date` sizin sürümünüzde hata verirse yalnızca döngünün başındaki iki satırı
> (`t` ve `yeniGun`) `Sistem.Tarih[i]` ile değiştirmeniz yeterli.

**Önemli fark:** IdealData backtesti bütün işlemleri bar kapanışından gerçekleştirir. Bu yüzden
stop ve iğne seviyeleri bar içinde görülse de dolum kapanış fiyatından olur. İğne kâr al
özelliğinin canlıda tam çalışması için, girişle birlikte `IGNE_TP_ATR` seviyesine limit emir ve
`STOP_ATR` seviyesine stop emri gönderilmelidir (robot/emir ayarlarınızdan). Python'da iki
durumu karşılaştırabilirsiniz: `-p DOLUM=intrabar` (seviyeden dolum) ve `-p DOLUM=kapanis`
(IdealData backtesti gibi).

## Python ile test ve optimizasyon

Gerekenler: Python 3.10+, `numpy`. (`pip install numpy`)

```bash
cd algo/backtest

# 1) IdealData'dan dışa aktardığınız 5dk veriyle varsayılan test
python sistem.py test XU030_5dk.txt

# 2) Parametre değiştirerek
python sistem.py test XU030_5dk.txt -p MOD=1 TP1_ATR=0.8 STOP_ATR=1.2 KALDIRAC=2

# 3) Walk-forward optimizasyon: verinin ilk %70'inde arar, son %30'unda doğrular
python sistem.py opt XU030_5dk.txt -n 400 --oran 0.7

# 4) Bulunan en iyi parametrelerle tekrar test
python sistem.py test XU030_5dk.txt --json sonuc/en_iyi_parametre.json
```

Çıktılar `sonuc/` klasörüne yazılır:
- `rapor.html`: getiri eğrisi, düşüş eğrisi, istatistikler, çıkış nedenleri. Optimizasyonda
  doğrulama dönemi kesikli çizgiyle ayrılır.
- `islemler.csv`: her işlem (Excel'de açılır, `;` ayraçlı)
- `optimizasyon.csv`: denenen tüm adaylar, İÇ ve DIŞ dönem sonuçları yan yana
- `en_iyi_parametre.json`

Veri dosyası: `Tarih;Saat;Açılış;Yüksek;Düşük;Kapanış;...` veya `TarihSaat;Açılış;...`
düzenleri, `; , tab` ayraçları ve ondalık virgül otomatik tanınır.
Veri yoksa `ORNEK` yazarak sentetik veriyle deneyebilirsiniz (`python sistem.py test ORNEK`).
Sentetik verinin sonuçları gerçek piyasayı **yansıtmaz**.

### Optimizasyon puanı

Adaylar şuna göre sıralanır: `net/MaxDD × R² (eğrinin doğrusallığı) × kazanma oranı × kâr faktörü`.
Yani "en düz damar, en küçük düşüş" arayan bir seçim. Seçimi **DIŞ dönemde de tutarlı kalan**
adaylardan yapın. İç dönemde parlayıp dış dönemde bozulan aday aşırı optimizasyondur.

## Parametreler

| Parametre | Varsayılan | Açıklama |
|---|---|---|
| MOD | 3 | 1 scalp, 2 trend, 3 hibrit |
| EMA_HIZLI / EMA_YAVAS | 9 / 34 | Yön filtresi |
| DONCHIAN | 20 | Kırılım kanalı (bar) |
| ADX_MIN / ER_MIN | 20 / 0.30 | Yatay piyasa filtresi |
| STOP_ATR | 1.5 | İlk stop (ATR katı) |
| TP1_ATR | 1.0 | Scalp hedefi / kâr kilidi tetiği |
| BE_ATR | 0.2 | TP1 sonrası stopun girişin ne kadar ötesine çekileceği |
| TRAIL_ATR / TRAIL_YUZDE | 2.0 / 3.0 | İzleyen stop (hangisi daha yakınsa) |
| IGNE_TP_ATR | 3.0 | İğne kâr al limiti |
| IGNE_FITIL_ATR | 1.5 | Bu boyda lehe fitil bırakan barda çık (99 = kapalı) |
| IGNE_GIRIS_ATR | 2.5 | Bu boydan büyük barda giriş yapma |
| MAX_BAR | 36 | TP1 görülmezse çıkış (bar) |
| GAP_ATR | 1.0 | "Büyük gap" eşiği |
| ACILIS_BEKLE / GAP_BEKLE | 2 / 6 | Açılışta beklenecek bar sayısı |
| SON_GIRIS / GUN_SONU | 1730 / 1755 | Saat (SSDD) |
| GUN_SONU_KAPAT | 1 | Gece pozisyon taşıma |
| GUNLUK_MAX_ZARAR_YUZDE | 2.0 | Gün içi zarar limiti |
| KOMISYON_YUZDE / KAYMA_YUZDE | 0.02 / 0.02 | Tek yön maliyet (yalnız Python) |
| KALDIRAC | 1 | Getiri çarpanı (yalnız Python) |
| DOLUM | intrabar | `intrabar` veya `kapanis` (yalnız Python) |

## Doğrulama notları

- Trendsiz rastgele yürüyüş verisinde (12 farklı veri, maliyetsiz) işlem başı ortalama getiri
  istatistiksel olarak sıfır çıktı. Yani motorda geleceği gören (look-ahead) bir hata yok;
  kâr ancak veride gerçek bir trend varsa oluşur.
- Bar içinde hem stop hem hedef görülürse önce stop varsayılır (kötümser).
- Hiçbir backtest gelecekteki performansı garanti etmez. Önce küçük lotla canlı gözlem yapın.
