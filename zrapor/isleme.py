"""Uçtan uca akış: fotoğraflar -> okuma -> birleştirme -> kontrol -> Excel (+ raporlar.json)."""
from __future__ import annotations

import json
from collections.abc import Callable, Iterable
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from .ayarlar import VARSAYILAN_ARKA_UC, VARSAYILAN_MODEL
from .excel import yaz
from .kontrol import birlestir, kontrol_et, ozet
from .okuma import oku

UZANTI = {".jpg", ".jpeg", ".png", ".webp", ".heic"}
Bildir = Callable[[str], None]


def fotolar(girdiler: Iterable[str | Path]) -> list[Path]:
    """Dosya ve klasör listesinden fotoğraf yollarını toplar (klasör içi sıralı)."""
    out: list[Path] = []
    for g in map(Path, girdiler):
        if g.is_dir():
            out += sorted(p for p in g.iterdir() if p.suffix.lower() in UZANTI)
        elif g.suffix.lower() in UZANTI:
            out.append(g)
    return out


def _z_degil_ayikla(ad: str, oks: dict[int, list], bildir: Bildir) -> tuple[list[list[dict]], list[str]]:
    """Okumalardan Z raporu olmayan belgeleri ayırır; temiz okuma listeleri ve not döndürür."""
    temiz, z_degil = [], []
    for n in sorted(oks):
        # Z no'su da kümülatifi de olmayan belge Z raporu değildir (satış fişi, fatura, slip)
        for r in oks[n]:
            if "_z_degil" not in r and not r.get("z_no") and r.get("kum_toplam") is None:
                r["_z_degil"] = [{"tur": "Z no'su olmayan belge", "satici": r.get("firma_unvan")}]
        temiz.append([r for r in oks[n] if "_z_degil" not in r])
        if n != min(oks):
            continue
        for r in oks[n]:
            if "_z_degil" in r:
                turler = ", ".join(sorted({(b.get("tur") or "?") + (f" ({b['satici']})" if b.get("satici") else "")
                                           for b in r["_z_degil"]}))
                z_degil.append(f"{ad}: {turler}")
                bildir(f"{ad}: Z raporu OLMAYAN {len(r['_z_degil'])} belge atlandi — {turler}")
    return temiz, z_degil


def kontrol_ve_yaz(raporlar: list[dict], cikti: Path, bildir: Bildir = print) -> list[dict]:
    """Okunmuş raporları kontrol eder, Excel ve yanına `raporlar.json` yazar; firma özetlerini döndürür."""
    firmalar = kontrol_et(raporlar)
    cikti.parent.mkdir(parents=True, exist_ok=True)
    yaz(firmalar, cikti)
    (cikti.parent / "raporlar.json").write_text(
        json.dumps(firmalar, ensure_ascii=False, default=str), encoding="utf-8")
    sonuc = ozet(firmalar)
    for o in sonuc:
        bildir(f"{o['firma']} ({o['vkn']}): {o['z']} Z · {o['hata']} hata · "
               f"{o['uyari']} uyarı · {o['duzeltme']} düzeltme")
    return sonuc


def isle(liste: list[Path], cikti: Path, model: str = VARSAYILAN_MODEL, arka_uc: str = VARSAYILAN_ARKA_UC,
         okuma: int = 2, paralel: int = 3, yeniden: bool = False, bildir: Bildir = print,
         calisma: Path | None = None) -> list[dict]:
    """Fotoğrafları oku, birleştir, kontrol et, Excel yaz. Firma başına özet döndürür.

    Her fotoğraf `okuma` kez bağımsız okunur; farklar uyarı olur. Okumalar `calisma`
    (varsayılan: <çıktı klasörü>/_okumalar) altında önbelleklenir.
    """
    calisma = calisma or cikti.parent / "_okumalar"
    okumalar: dict[Path, dict[int, list]] = {f: {} for f in liste}
    toplam, biten = len(liste) * okuma, 0
    with ThreadPoolExecutor(max(1, paralel)) as ex:
        isler = {ex.submit(oku, f, calisma, model, arka_uc, yeniden, n): (f, n)
                 for f in liste for n in range(1, okuma + 1)}
        for i in as_completed(isler):
            f, n = isler[i]
            biten += 1
            try:
                rs = i.result()
                bildir(f"[{biten}/{toplam}] {f.name} (okuma {n}): {sum('_z_degil' not in r for r in rs)} Z raporu")
                okumalar[f][n] = rs
            except Exception as e:  # bir fotoğraf bozuksa diğerleri devam etsin
                bildir(f"[{biten}/{toplam}] {f.name} (okuma {n}): OKUNAMADI — {e}")

    return okumalardan_yaz({f.name: oks for f, oks in okumalar.items()}, cikti, bildir)


def okumalardan_yaz(okumalar: dict[str, dict[int, list]], cikti: Path, bildir: Bildir = print) -> list[dict]:
    """Fotoğraf adı -> {okuma no: rapor listesi}: Z olmayanları ayıklar, birleştirir, kontrol eder, yazar."""
    raporlar, z_degil = [], []
    for ad, oks in okumalar.items():
        if not oks:
            bildir(f"{ad}: hic okunamadi, Excel'de yok!")
            continue
        temiz, zd = _z_degil_ayikla(ad, oks, bildir)
        z_degil += zd
        raporlar += birlestir(temiz)
    if not raporlar:
        if z_degil:
            raise RuntimeError("Bu fotograflarda Z raporu yok. Gorulen belgeler: " + " | ".join(z_degil)
                               + ". Bu arac yalniz yazar kasa Z (gun sonu) raporlarini okur.")
        raise RuntimeError("hicbir fotograftan Z raporu okunamadi")
    return kontrol_ve_yaz(raporlar, cikti, bildir)


def okuma_dosyasi_yukle(yol: Path) -> list[dict]:
    """Hazır okuma JSON'unu yükler (API'siz doğrulama için).

    Kabul edilen biçimler:
      - `[rapor, ...]`                                   tek okuma
      - `{"raporlar": [rapor, ...]}`                     model çıktısı biçimi
      - `{"okumalar": [[rapor, ...], [rapor, ...]]}`     aynı fişlerin bağımsız okumaları
    """
    veri = json.loads(Path(yol).read_text(encoding="utf-8"))
    if isinstance(veri, dict) and "okumalar" in veri:
        okumalar = [[dict(r) for r in o] for o in veri["okumalar"]]
    elif isinstance(veri, dict):
        okumalar = [list(veri.get("raporlar", []))]
    else:
        okumalar = [list(veri)]
    for o in okumalar:
        for r in o:
            r.setdefault("_kaynak", Path(yol).name)
    return birlestir(okumalar)
