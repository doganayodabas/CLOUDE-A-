"""
Komut satırı: test ve walk-forward optimizasyon.

Örnekler
--------
  # Varsayılan parametrelerle test + HTML rapor
  python sistem.py test VIOP_XU030_5dk.txt

  # Parametre değiştirerek test
  python sistem.py test VIOP_XU030_5dk.txt -p MOD=1 STOP_ATR=1.2 TP1_ATR=0.8

  # IdealData backtestine benzer dolum (her şey bar kapanışından)
  python sistem.py test VIOP_XU030_5dk.txt -p DOLUM=kapanis

  # Optimizasyon: verinin %70'inde ara, %30'unda (hiç görmediği dönemde) doğrula
  python sistem.py opt VIOP_XU030_5dk.txt -n 400 --oran 0.7

  # Veri yoksa sentetik veriyle deneme
  python sistem.py test ORNEK
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import random
import sys
from concurrent.futures import ProcessPoolExecutor

import numpy as np

import motor

# Optimizasyon arama uzayı
ARALIK = {
    "MOD": [1, 2, 3],
    "EMA_HIZLI": [5, 8, 9, 12, 15],
    "EMA_YAVAS": [21, 26, 34, 50, 72],
    "DONCHIAN": [10, 15, 20, 30],
    "ADX_MIN": [15, 18, 20, 23, 25, 30],
    "ER_MIN": [0.15, 0.2, 0.25, 0.3, 0.35, 0.4],
    "STOP_ATR": [0.8, 1.0, 1.2, 1.5, 2.0, 2.5],
    "TP1_ATR": [0.5, 0.7, 0.8, 1.0, 1.2, 1.5],
    "BE_ATR": [0.0, 0.1, 0.2, 0.3],
    "TRAIL_ATR": [1.0, 1.5, 2.0, 2.5, 3.0],
    "TRAIL_YUZDE": [1.0, 2.0, 3.0],
    "IGNE_TP_ATR": [2.0, 2.5, 3.0, 4.0, 5.0],
    "IGNE_FITIL_ATR": [1.0, 1.5, 2.0, 99.0],
    "MAX_BAR": [12, 24, 36, 48, 72],
}


def veri_al(yol: str) -> motor.Veri:
    if yol.upper() == "ORNEK":
        print("UYARI: sentetik örnek veri kullanılıyor; sonuçlar gerçek piyasayı yansıtmaz.")
        return motor.ornek_veri()
    return motor.veri_oku(yol)


def parametre_coz(liste) -> dict:
    p = {}
    for x in liste or []:
        k, val = x.split("=", 1)
        k = k.strip().upper()
        if k not in motor.VARSAYILAN:
            sys.exit(f"Bilinmeyen parametre: {k}")
        ref = motor.VARSAYILAN[k]
        p[k] = val if isinstance(ref, str) else type(ref)(float(val)) if isinstance(ref, int) else float(val)
    return p


def bol(v: motor.Veri, oran: float):
    k = int(len(v) * oran)
    # gün ortasından bölme
    gun = v.zaman.astype("datetime64[D]")
    while 0 < k < len(v) and gun[k] == gun[k - 1]:
        k += 1
    return v.dilim(0, k), v.dilim(k, len(v))


def puan(s: dict, min_islem: int) -> float:
    """Getiri/maxDD, eğri doğrusallığı ve kazanma oranını birlikte ödüllendirir."""
    if s["islem"] < min_islem or s["net_yuzde"] <= 0:
        return -1e9 + s["net_yuzde"]
    return s["net_maxdd"] * max(s["r2"], 0) * (0.5 + s["kazanan_yuzde"] / 100) * min(s["kar_faktoru"], 5)


# ---------------------------------------------------------------- rapor

def html_rapor(yol: str, v: motor.Veri, sonuc: motor.Sonuc, p: dict, baslik: str, bolme_idx: int | None = None):
    oz = sonuc.ozkaynak
    adim = max(1, len(oz) // 1500)
    y = oz[::adim]
    W, H = 1000, 340
    lo, hi = float(min(y.min(), 0)), float(max(y.max(), 0.01))
    def px(i): return 40 + i * (W - 50) / max(len(y) - 1, 1)
    def py(val): return 10 + (hi - val) * (H - 40) / (hi - lo or 1)
    yol_d = " ".join(f"{'M' if i == 0 else 'L'}{px(i):.1f},{py(val):.1f}" for i, val in enumerate(y))
    tepe = np.maximum.accumulate(y)
    dd_d = " ".join(f"{'M' if i == 0 else 'L'}{px(i):.1f},{py(val):.1f}" for i, val in enumerate(y - tepe))
    ayrac = ""
    if bolme_idx:
        x = px(bolme_idx / adim)
        ayrac = (f'<line x1="{x:.1f}" y1="10" x2="{x:.1f}" y2="{H-30}" class="ayr"/>'
                 f'<text x="{x+4:.1f}" y="24" class="t">optimizasyon dışı dönem →</text>')
    def tarih(i): return str(v.zaman[min(i, len(v) - 1)])[:10]
    eksen = "".join(f'<text x="{px(k):.1f}" y="{H-12}" class="t" text-anchor="middle">{tarih(k*adim)}</text>'
                    for k in np.linspace(0, len(y) - 1, 6).astype(int))
    s = sonuc.istat
    satir = "".join(
        f"<tr><td>{k}</td><td>{(f'{x:.2f}' if isinstance(x, float) else x)}</td></tr>"
        for k, x in s.items() if k != "cikis_nedenleri")
    neden = ", ".join(f"{k}: {x}" for k, x in sorted(s["cikis_nedenleri"].items(), key=lambda t: -t[1]))
    par = "".join(f"<tr><td>{k}</td><td>{x}</td></tr>" for k, x in p.items())
    html = f"""<!doctype html><html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{baslik}</title>
<style>
:root{{--bg:#fff;--fg:#1d2330;--mut:#6b7280;--acc:#2563eb;--neg:#dc2626;--bd:#e5e7eb}}
@media (prefers-color-scheme:dark){{:root{{--bg:#111418;--fg:#e6e8eb;--mut:#9aa3ad;--acc:#60a5fa;--neg:#f87171;--bd:#2a2f36}}}}
body{{background:var(--bg);color:var(--fg);font:14px system-ui,sans-serif;margin:0;padding:16px;max-width:1040px}}
svg{{width:100%;height:auto}} .e{{fill:none;stroke:var(--acc);stroke-width:1.6}} .d{{fill:none;stroke:var(--neg);stroke-width:1}}
.t{{fill:var(--mut);font-size:11px}} .ayr{{stroke:var(--mut);stroke-dasharray:4 4}} .z{{stroke:var(--bd)}}
table{{border-collapse:collapse;margin:8px 16px 16px 0;display:inline-table;vertical-align:top}}
td{{border-bottom:1px solid var(--bd);padding:3px 10px}} td:last-child{{text-align:right;font-variant-numeric:tabular-nums}}
</style></head><body><h2>{baslik}</h2>
<p style="color:var(--mut)">Mavi: kümülatif getiri (% , toplamsal). Kırmızı: tepeden düşüş.</p>
<svg viewBox="0 0 {W} {H}"><line x1="40" x2="{W-10}" y1="{py(0):.1f}" y2="{py(0):.1f}" class="z"/>
<path d="{yol_d}" class="e"/><path d="{dd_d}" class="d"/>{ayrac}{eksen}
<text x="2" y="{py(hi)+4:.1f}" class="t">{hi:.0f}%</text><text x="2" y="{py(lo)+4:.1f}" class="t">{lo:.0f}%</text></svg>
<table>{satir}<tr><td>çıkış nedenleri</td><td>{neden}</td></tr></table><table>{par}</table>
</body></html>"""
    with open(yol, "w", encoding="utf-8") as f:
        f.write(html)


def islem_csv(yol: str, sonuc: motor.Sonuc):
    with open(yol, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["yon", "giris_zaman", "giris", "cikis_zaman", "cikis", "neden", "bar", "getiri_yuzde"])
        for t in sonuc.islemler:
            w.writerow(["LONG" if t.yon == 1 else "SHORT", str(t.giris_zaman), f"{t.giris:.4f}",
                        str(t.cikis_zaman), f"{t.cikis:.4f}", t.neden, t.bar, f"{t.getiri:.4f}".replace(".", ",")])


def ozet(s: dict) -> str:
    return (f"işlem {s['islem']:5d} | kazanan %{s['kazanan_yuzde']:5.1f} | net %{s['net_yuzde']:8.2f} | "
            f"maxDD %{s['maxdd_yuzde']:6.2f} | net/DD {s['net_maxdd']:6.2f} | PF {s['kar_faktoru']:.2f} | R² {s['r2']:.2f}")


# ---------------------------------------------------------------- komutlar

def komut_test(a):
    v = veri_al(a.veri)
    p = dict(motor.VARSAYILAN)
    if a.json:
        p.update(json.load(open(a.json, encoding="utf-8")))
    p.update(parametre_coz(a.p))
    s = motor.calistir(v, p)
    print(f"{len(v)} bar, {str(v.zaman[0])} → {str(v.zaman[-1])}")
    print(ozet(s.istat))
    print("çıkış nedenleri:", s.istat["cikis_nedenleri"])
    os.makedirs(a.cikti, exist_ok=True)
    html_rapor(os.path.join(a.cikti, "rapor.html"), v, s, p, "Trend+Scalp 5dk — test")
    islem_csv(os.path.join(a.cikti, "islemler.csv"), s)
    print(f"Rapor: {os.path.join(a.cikti, 'rapor.html')}  İşlemler: {os.path.join(a.cikti, 'islemler.csv')}")


_GLOBAL = {}


def _isci_hazirla(ic, dis, sabit):
    _GLOBAL["ic"], _GLOBAL["dis"], _GLOBAL["sabit"] = ic, dis, sabit


def _dene(aday):
    p = dict(_GLOBAL["sabit"]); p.update(aday)
    if p["EMA_HIZLI"] >= p["EMA_YAVAS"]:
        return None
    si = motor.calistir(_GLOBAL["ic"], p).istat
    so = motor.calistir(_GLOBAL["dis"], p).istat if len(_GLOBAL["dis"]) > 200 else None
    return aday, si, so


def komut_opt(a):
    v = veri_al(a.veri)
    ic, dis = bol(v, a.oran)
    print(f"Optimizasyon dönemi: {str(ic.zaman[0])[:10]} → {str(ic.zaman[-1])[:10]} ({len(ic)} bar)")
    if len(dis):
        print(f"Doğrulama dönemi  : {str(dis.zaman[0])[:10]} → {str(dis.zaman[-1])[:10]} ({len(dis)} bar)")
    sabit = dict(motor.VARSAYILAN)
    sabit.update(parametre_coz(a.p))
    rnd = random.Random(a.tohum)
    adaylar, gorulen = [], set()
    while len(adaylar) < a.n:
        c = {k: rnd.choice(x) for k, x in ARALIK.items() if k not in parametre_coz(a.p)}
        anahtar = tuple(sorted(c.items()))
        if anahtar not in gorulen:
            gorulen.add(anahtar); adaylar.append(c)
    sonuclar = []
    with ProcessPoolExecutor(initializer=_isci_hazirla, initargs=(ic, dis, sabit)) as ex:
        for k, r in enumerate(ex.map(_dene, adaylar, chunksize=4), 1):
            if r:
                sonuclar.append(r)
            if k % 50 == 0:
                print(f"  {k}/{a.n} denendi")
    min_islem = max(30, int(len(ic) / 2000))
    sonuclar.sort(key=lambda r: puan(r[1], min_islem), reverse=True)

    os.makedirs(a.cikti, exist_ok=True)
    yol = os.path.join(a.cikti, "optimizasyon.csv")
    anahtarlar = list(ARALIK)
    with open(yol, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        alan = ["islem", "kazanan_yuzde", "net_yuzde", "maxdd_yuzde", "net_maxdd", "kar_faktoru", "r2"]
        w.writerow(["sira", "puan"] + [f"IC_{x}" for x in alan] + [f"DIS_{x}" for x in alan] + anahtarlar)
        for s_, (c, si, so) in enumerate(sonuclar, 1):
            w.writerow([s_, f"{puan(si, min_islem):.3f}"] + [f"{si[x]:.3f}" for x in alan]
                       + ([f"{so[x]:.3f}" for x in alan] if so else [""] * len(alan))
                       + [c.get(x, sabit[x]) for x in anahtarlar])

    print(f"\nEn iyi {a.ilk} aday (İÇ = optimizasyon dönemi, DIŞ = görülmemiş dönem):")
    for c, si, so in sonuclar[:a.ilk]:
        print("  İÇ :", ozet(si))
        if so:
            print("  DIŞ:", ozet(so))
        print("     ", json.dumps(c, ensure_ascii=False))
    en_iyi = dict(sabit); en_iyi.update(sonuclar[0][0])
    with open(os.path.join(a.cikti, "en_iyi_parametre.json"), "w", encoding="utf-8") as f:
        json.dump(en_iyi, f, ensure_ascii=False, indent=2)
    tam = motor.calistir(v, en_iyi)
    html_rapor(os.path.join(a.cikti, "rapor.html"), v, tam, en_iyi, "Trend+Scalp 5dk — optimizasyon en iyisi",
               bolme_idx=len(ic) if len(dis) else None)
    islem_csv(os.path.join(a.cikti, "islemler.csv"), tam)
    print(f"\nTüm sonuçlar: {yol}\nEn iyi parametreler: {os.path.join(a.cikti, 'en_iyi_parametre.json')}"
          f"\nRapor: {os.path.join(a.cikti, 'rapor.html')}")
    print("Not: Seçimi İÇ sonuca göre değil, DIŞ dönemde de tutarlı kalan adaylara göre yapın.")


def main():
    ap = argparse.ArgumentParser(description="Trend+Scalp 5dk sistem testi")
    alt = ap.add_subparsers(dest="komut", required=True)
    t = alt.add_parser("test"); t.add_argument("veri"); t.add_argument("-p", nargs="*")
    t.add_argument("--json", help="en_iyi_parametre.json gibi bir dosyadan parametre oku")
    t.add_argument("-o", "--cikti", default="sonuc")
    o = alt.add_parser("opt"); o.add_argument("veri"); o.add_argument("-p", nargs="*", help="sabit tutulacak parametreler")
    o.add_argument("-n", type=int, default=300); o.add_argument("--oran", type=float, default=0.7)
    o.add_argument("--ilk", type=int, default=5); o.add_argument("--tohum", type=int, default=1)
    o.add_argument("-o", "--cikti", default="sonuc")
    a = ap.parse_args()
    (komut_test if a.komut == "test" else komut_opt)(a)


if __name__ == "__main__":
    main()
