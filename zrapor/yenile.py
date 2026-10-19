"""Model çağrısı yapmadan, `_okumalar/` önbelleğinden kontrol + Excel'i yeniden üretir.

Kontrol ya da Excel kodu değiştiğinde, ya da hesap kodu/şablon güncellendiğinde kullanılır:

    z-rapor yenile <cikti.xlsx>                 (önbellek: dosyanın yanındaki _okumalar/)
    z-rapor yenile --veri ./calisma --is <id>   (telefon sayfasının bitmiş bir işi)

Önbellek dosya adı: `[masraf_]<foto>_<özet12>_<model>_<arka uç>[_o<N>].json`. Aynı fotoğrafın
aynı okuma numarası birden çok kez varsa (ör. farklı modelle) en yenisi kullanılır.
"""
from __future__ import annotations

import json
import re
from collections.abc import Callable
from pathlib import Path

DESEN = re.compile(r"^(?P<onek>masraf_)?(?P<stem>.+)_(?P<ozet>[0-9a-f]{12})_(?P<model>[^_]+)"
                   r"_(?P<uc>api|cli|gemini)(?:_o(?P<n>\d+))?\.json$")


def onbellek_oku(klasor: Path, tur: str) -> dict[str, dict[int, dict]]:
    """Önbellekteki ham okumaları fotoğraf adı -> {okuma no: ham çıktı} olarak toplar.

    tur: "z" (Z raporu) ya da "masraf".
    """
    gruplar: dict[tuple[str, str], dict] = {}
    dosyalar = sorted((p for p in klasor.glob("*.json") if p.is_file()), key=lambda p: p.stat().st_mtime)
    for p in dosyalar:
        m = DESEN.match(p.name)
        if not m or bool(m["onek"]) != (tur == "masraf"):
            continue
        try:
            veri = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        g = gruplar.setdefault((m["stem"], m["ozet"]), {"ad": veri.get("_kaynak") or m["stem"], "oks": {}})
        g["oks"][int(m["n"] or 1)] = veri  # sıralı: aynı numarada en yeni dosya kalır
    adlar = [g["ad"] for g in gruplar.values()]
    sonuc: dict[str, dict[int, dict]] = {}
    for (stem, ozet), g in sorted(gruplar.items()):
        ad = g["ad"] if adlar.count(g["ad"]) == 1 else f"{g['ad']} ({ozet[:6]})"
        sonuc[ad] = g["oks"]
    return sonuc


def tur_bul(klasor: Path) -> str:
    """Önbellekte yalnız masraf okuması varsa "masraf", yoksa "z"."""
    adlar = [m for m in (DESEN.match(p.name) for p in klasor.glob("*.json")) if m]
    if adlar and all(m["onek"] for m in adlar):
        return "masraf"
    return "z"


def yenile(cikti: Path, tur: str | None = None, hazirlayan: str = "", calisma: Path | None = None,
           bildir: Callable[[str], None] = print):
    """`cikti`yı önbellekten yeniden yazar. Z: firma özetleri listesi; masraf: özet sözlüğü döner."""
    calisma = calisma or cikti.parent / "_okumalar"
    if not calisma.is_dir():
        raise FileNotFoundError(f"onbellek klasoru yok: {calisma}")
    tur = tur or tur_bul(calisma)
    ham = onbellek_oku(calisma, tur)
    if not ham:
        raise RuntimeError(f"{calisma} icinde {'masraf' if tur == 'masraf' else 'Z raporu'} okumasi yok")
    bildir(f"onbellekten {len(ham)} fotograf, {sum(len(o) for o in ham.values())} okuma (model cagrisi yok)")
    if tur == "masraf":
        from . import masraf
        return masraf.okumalardan_yaz({ad: {n: masraf.belgelere(v, ad) for n, v in oks.items()}
                                       for ad, oks in ham.items()}, cikti, hazirlayan, bildir)
    from .isleme import okumalardan_yaz
    from .okuma import raporlara
    return okumalardan_yaz({ad: {n: raporlara(v, ad) for n, v in oks.items()} for ad, oks in ham.items()},
                           cikti, bildir)


def yenile_is(veri: Path, is_id: str, bildir: Callable[[str], None] = print):
    """Telefon sayfasının bitmiş bir işini önbellekten yeniden üretir; `durum.json` özetini günceller.

    Çalışan sunucu işleri bellekte tutar: değişikliği sayfada görmek için sunucuyu yeniden başlatın.
    """
    if not re.fullmatch(r"[\w\-]+", is_id):
        raise ValueError(f"gecersiz is kimligi: {is_id}")
    klasor = Path(veri) / "isler" / is_id
    durum_yolu = klasor / "durum.json"
    if not durum_yolu.is_file():
        raise FileNotFoundError(f"is bulunamadi: {durum_yolu}")
    durum = json.loads(durum_yolu.read_text(encoding="utf-8"))
    if durum.get("durum") != "bitti" or not durum.get("excel"):
        raise RuntimeError(f"{is_id}: is bitmemis (durum={durum.get('durum')})")
    cikti = klasor / "cikti" / durum["excel"]
    if durum.get("tur") == "masraf":
        durum["masraf_ozet"] = yenile(cikti, "masraf", durum.get("hazirlayan") or "", bildir=bildir)
    else:
        durum["ozet"] = yenile(cikti, "z", bildir=bildir)
    durum_yolu.write_text(json.dumps(durum, ensure_ascii=False, indent=1), encoding="utf-8")
    return cikti
