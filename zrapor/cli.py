"""`z-rapor` komut satırı.

    z-rapor oku <foto|klasör>... -o cikti.xlsx [--arka-uc api|cli|gemini|karma] [--model sonnet] [--okuma 2]
                                              [--paralel 3] [--yeniden]
    z-rapor dogrula <okuma.json> -o cikti.xlsx          (API'siz: hazır okumadan kontrol + Excel)
    z-rapor yenile <cikti.xlsx> [--tur z|masraf]        (model çağrısız: _okumalar/ önbelleğinden yeniden)
    z-rapor yenile --veri ./calisma --is <id>           (telefon sayfasının bir işini yeniden üret)
    z-rapor luca <raporlar.json> --tur bilanco|isletme [--hesaplar h.json] -o dosya
    z-rapor masraf <foto|klasör>... -o liste.xlsx [--hazirlayan AD]
    z-rapor web [--port 8765] [--veri ./calisma]
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

from . import __version__
from .ayarlar import VARSAYILAN_ARKA_UC, VARSAYILAN_CALISMA, VARSAYILAN_MODEL, VARSAYILAN_PORT, env_yukle
from .okuma import ARKA_UCLAR


def _okuma_secenekleri(p: argparse.ArgumentParser) -> None:
    p.add_argument("--arka-uc", default=VARSAYILAN_ARKA_UC, choices=ARKA_UCLAR,
                   help="api: Anthropic API (ANTHROPIC_API_KEY) | cli: claude -p | gemini: Google Gemini API "
                        "(GEMINI_API_KEY) | karma: 1. okuma api, 2. okuma gemini (varsayılan: %(default)s)")
    p.add_argument("--model", default=VARSAYILAN_MODEL, help="kısa ad (sonnet, opus) ya da tam model kimliği")
    p.add_argument("--okuma", type=int, default=2, help="her fotoğraf kaç kez bağımsız okunsun (farklar uyarı olur)")
    p.add_argument("--paralel", type=int, default=3, help="aynı anda kaç model çağrısı")
    p.add_argument("--yeniden", action="store_true", help="önbelleği yok say, yeniden oku")


def _ayristirici() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(prog="z-rapor", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    alt = ap.add_subparsers(dest="komut", required=True)

    p = alt.add_parser("oku", help="Z raporu fotoğraflarını oku, kontrol et, Excel yaz")
    p.add_argument("girdi", nargs="+")
    p.add_argument("-o", "--cikti")
    _okuma_secenekleri(p)

    p = alt.add_parser("dogrula", help="hazır okuma JSON'undan (API'siz) kontrol + Excel")
    p.add_argument("okuma_json")
    p.add_argument("-o", "--cikti", required=True)

    p = alt.add_parser("yenile", help="model çağrısı yapmadan _okumalar/ önbelleğinden kontrol + Excel")
    p.add_argument("cikti", nargs="?", help="yeniden yazılacak Excel (önbellek: yanındaki _okumalar/)")
    p.add_argument("--tur", choices=("z", "masraf"), help="önbellekte ikisi de varsa hangisi (varsayılan: z)")
    p.add_argument("--hazirlayan", default="", help="masraf listesi için")
    p.add_argument("--veri", help="telefon sayfasının veri klasörü (--is ile)")
    p.add_argument("--is", dest="is_id", help="telefon sayfasındaki iş kimliği")

    p = alt.add_parser("luca", help="kontrol edilmiş raporlar.json'dan Luca aktarım dosyası")
    p.add_argument("raporlar_json")
    p.add_argument("--tur", choices=("bilanco", "isletme"), required=True)
    p.add_argument("--hesaplar", help="firma (VKN) başına hesap kodları JSON dosyası")
    p.add_argument("-o", "--cikti", required=True)

    p = alt.add_parser("masraf", help="masraf fişi fotoğraflarından masraf listesi")
    p.add_argument("girdi", nargs="+")
    p.add_argument("-o", "--cikti", required=True)
    p.add_argument("--hazirlayan", default="")
    _okuma_secenekleri(p)

    p = alt.add_parser("web", help="telefon yükleme sayfası")
    p.add_argument("--port", type=int, default=VARSAYILAN_PORT)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--veri", default=str(VARSAYILAN_CALISMA), help="iş kuyruğu ve hesap kodları klasörü")
    p.add_argument("--arka-uc", default=os.environ.get("ZRAPOR_ARKA_UC") or VARSAYILAN_ARKA_UC, choices=ARKA_UCLAR,
                   help="okuma arka ucu; ortamda ZRAPOR_ARKA_UC (varsayılan: %(default)s)")
    p.add_argument("--model", default=VARSAYILAN_MODEL)
    return ap


def _yaz(m: str) -> None:
    print(m, flush=True)


def _oku(a) -> None:
    from .isleme import fotolar, isle

    liste = fotolar(a.girdi)
    if not liste:
        sys.exit("fotoğraf bulunamadı")
    cikti = Path(a.cikti) if a.cikti else liste[0].parent / f"z_raporlari_{datetime.now():%Y%m%d_%H%M}.xlsx"
    isle(liste, cikti, a.model, a.arka_uc, a.okuma, a.paralel, a.yeniden, bildir=_yaz)
    print(f"Excel: {cikti}")


def _dogrula(a) -> None:
    from .isleme import kontrol_ve_yaz, okuma_dosyasi_yukle

    cikti = Path(a.cikti)
    kontrol_ve_yaz(okuma_dosyasi_yukle(Path(a.okuma_json)), cikti, bildir=_yaz)
    print(f"Excel: {cikti}")
    print(f"Kontrol verisi: {cikti.parent / 'raporlar.json'}")


def _yenile(a) -> None:
    from .yenile import yenile, yenile_is

    if a.is_id:
        if not a.veri:
            sys.exit("--is ile --veri de verin")
        print(f"Excel: {yenile_is(Path(a.veri), a.is_id, bildir=_yaz)}")
        print("Çalışan telefon sayfası özeti yeniden başlatınca görür.")
        return
    if not a.cikti:
        sys.exit("yenilenecek Excel dosyasını ya da --veri ve --is verin")
    cikti = Path(a.cikti)
    yenile(cikti, a.tur, a.hazirlayan, bildir=_yaz)
    print(f"Excel: {cikti}")


def _luca(a) -> None:
    from . import luca

    firmalar = json.loads(Path(a.raporlar_json).read_text(encoding="utf-8"))
    hesaplar = luca.hesaplar_oku(a.hesaplar)
    yaz = luca.bilanco_yaz if a.tur == "bilanco" else luca.isletme_yaz
    o = yaz(firmalar, Path(a.cikti), hesaplar)
    print(f"{o['fis']} fiş, {o['satir']} satır -> {a.cikti}")
    for x in o["aktarilmayan"]:
        print(f"  aktarılmadı: {x['tarih']} Z {x['z_no']} — {x['neden']}")
    if o.get("tahmini_kodlar"):
        print("  kontrol edilmemiş hesap kodları (--hesaplar ile verin): " + ", ".join(o["tahmini_kodlar"]))
    if o.get("eksik_faaliyet"):
        print("  faaliyet kodu eksik: " + ", ".join(o["eksik_faaliyet"]))


def _masraf(a) -> None:
    from . import masraf
    from .isleme import fotolar

    liste = fotolar(a.girdi)
    if not liste:
        sys.exit("fotoğraf bulunamadı")
    masraf.isle(liste, Path(a.cikti), a.model, a.arka_uc, a.okuma, a.paralel, a.yeniden,
                hazirlayan=a.hazirlayan, bildir=_yaz)
    print(f"Excel: {a.cikti}")


def _web(a) -> None:
    from .web import calistir

    calistir(port=a.port, host=a.host, veri=Path(a.veri), model=a.model, arka_uc=a.arka_uc)


KOMUTLAR = {"oku": _oku, "dogrula": _dogrula, "yenile": _yenile, "luca": _luca, "masraf": _masraf, "web": _web}


def main(argv: list[str] | None = None) -> None:
    """Komut satırı girişi."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    env_yukle(Path.cwd() / ".env")
    a = _ayristirici().parse_args(argv)
    KOMUTLAR[a.komut](a)


if __name__ == "__main__":
    main()
