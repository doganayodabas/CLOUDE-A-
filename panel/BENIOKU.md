# Odabaşlar Giyim Panel

`odabaslar_panel.html`, Odabaşlar Giyim web panelinin **Hatırlatıcı** modülüdür.
Tek dosyalık bir web sayfasıdır.

Yayındaki adresi: https://claude.ai/artifact/NKKzEcYDhn6wzkRUyQm9VQ

Veriler panelin ortak veritabanında tutulur, telefondan ve bilgisayardan aynı veriler görünür.
Sayfa veritabanına bağlanamazsa veriler yalnız o tarayıcıda saklanır.

## Bölümler

| Bölüm | İçerik |
|---|---|
| Özet | Gecikmiş, bugün–3 gün, 4–7 gün ve 30 güne kadarki ödemeler; acil liste; 8 haftalık çek yükü; banka yeterliliği; personel hatırlatmaları |
| Ödemeler | Çek, kredi kartı, DBS, fatura, düzenli ve diğer ödemeler. Filtre, arama, "Ödendi" ve düzenleme |
| Çek Takvimi | Banka bazında aylık ve haftalık takvim, lehdar bazında toplamlar |
| Bankalar | Her hesap için 30 günlük bakiye projeksiyonu (KMH dahil) ve banka hareketleri |
| Raporlar | Bilanço düzeninde gider raporu, Mağaza / Online ayrımı, aylık bekleyenler, DBS limit doluluğu |
| Personel & İzin | Yıllık izin hakkı (İş Kanunu m.53), kalan izin, izin takvimi, hak ediş ve doğum günü hatırlatmaları |
| Tanımlar | Kredi kartları, düzenli ödemeler, genel ayarlar, Excel içe/dışa aktarma |

## Python hatırlatıcısıyla bağı

* Excel içe/dışa aktarma, `hatirlatici/HATIRLATICI.xlsx` şablonuyla aynı düzeni kullanır.
* Panelden aktarılan dosyayı şu komutla işlerseniz telefona alınabilecek `.ics` takvim dosyası da üretilir:

```bash
python hatirlatici/hatirlatici.py dosya.xlsx
```

## Başka bir panele eklemek

Dosya kendi başına çalışır. Bir HTML sayfasına `<iframe>` olarak ya da doğrudan sayfa olarak eklenebilir.

Ortak veritabanı yalnız claude.ai'de yayınlandığında çalışır. Başka bir sunucuda veriler o tarayıcının
`localStorage` alanında tutulur. Kendi sunucunuzda ortak kayıt için `kaydet`, `sil` ve `baslat`
fonksiyonlarını kendi API'nize bağlayın.
