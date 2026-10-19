"""Uydurma bir Z raporu fişini PNG olarak çizer (API ile denemek için).

    python ornekler/uret_fis.py            -> ornekler/z_ornek.png

Rakamlar `okuma_ornek.json`daki ilk günle (Z 0101) aynıdır; okunan sonucu oradakiyle
karşılaştırabilirsiniz. İşletme, VKN ve cihaz numarası uydurmadır.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

GENISLIK, KENAR, SATIR = 640, 36, 34
YAZI_TIPLERI = ("DejaVuSansMono.ttf", "consola.ttf", "cour.ttf", "LiberationMono-Regular.ttf",
                "/System/Library/Fonts/Menlo.ttc")


def yazi_tipi(boyut: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    """Sistemde bulunan ilk eş aralıklı yazı tipi; yoksa Pillow'un varsayılanı."""
    for ad in YAZI_TIPLERI:
        try:
            return ImageFont.truetype(ad, boyut)
        except OSError:
            continue
    return ImageFont.load_default(size=boyut)


def tl(x: float) -> str:
    """1234.5 -> '*1.234,50' (yazar kasa biçimi)."""
    return "*" + f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# (sol, sağ) satırlar; sağ boşsa satır ortalanır, "-" çizgi demek
SATIRLAR: list[tuple[str, str]] = [
    ("ÖRNEK LOKANTA", ""), ("ÖRNEK KİŞİ", ""), ("ÖRNEK MAH. DENEME SK. NO:1", ""),
    ("ÖRNEK VD 1234567890", ""), ("-", ""),
    ("TARİH 01.09.2026", "SAAT 23:50"), ("FİŞ NO", "0612"), ("Z RAPORU", ""), ("-", ""),
    ("Yiyecek  %10", "40"), ("", tl(8110.00)), ("İçecek   %20", "35"), ("", tl(1240.00)), ("-", ""),
    ("NET SATIŞ", tl(9350.00)), ("NAKİT", "9"), ("", tl(2450.00)), ("KASA NAKİT", tl(2450.00)),
    ("K.KARTI", "31"), ("", tl(6900.00)), ("ÖRNEKBANK", tl(6900.00)), ("FİŞ İPTAL", "0"), ("-", ""),
    ("MÜŞTERİ FİŞ ADEDİ", "40"), ("MALİ FİŞ ADEDİ", "0"), ("SLIP FİŞ ADEDİ", "31"),
    ("TOPLAM FİŞ ADEDİ", "71"), ("-", ""), ("VERGİ DÖKÜMÜ", ""),
    ("%10 TOPLAM", tl(8110.00)), ("%10 TOPKDV", tl(737.27)),
    ("%20 TOPLAM", tl(1240.00)), ("%20 TOPKDV", tl(206.67)), ("-", ""),
    ("MALİ VERİ TOPLAM", tl(9350.00)), ("MALİ VERİ TOPKDV", tl(943.94)),
    ("KÜM. TOPLAM", tl(1259350.00)), ("KÜM. TOPKDV", tl(118943.94)), ("-", ""),
    ("Z SAYAÇ", "0101"), ("EKÜ NO 0001", ""), ("AF AV0000000001", ""),
]


def ciz(hedef: Path) -> Path:
    """Fişi çizip `hedef`e kaydeder."""
    f = yazi_tipi(24)
    yukseklik = KENAR * 2 + SATIR * len(SATIRLAR)
    fis = Image.new("RGB", (GENISLIK, yukseklik), (252, 251, 247))
    d = ImageDraw.Draw(fis)
    y = KENAR
    for sol, sag in SATIRLAR:
        if sol == "-":
            d.text((KENAR, y), "-" * 40, font=f, fill=(90, 90, 90))
        elif not sag and sol:
            w = d.textlength(sol, font=f)
            d.text(((GENISLIK - w) / 2, y), sol, font=f, fill=(25, 25, 25))
        else:
            d.text((KENAR, y), sol, font=f, fill=(25, 25, 25))
            w = d.textlength(sag, font=f)
            d.text((GENISLIK - KENAR - w, y), sag, font=f, fill=(25, 25, 25))
        y += SATIR
    # fişi gri bir masa zeminine hafif eğik koy (fotoğraf gibi)
    zemin = Image.new("RGB", (GENISLIK + 240, yukseklik + 240), (122, 128, 136))
    zemin.paste(fis.rotate(1.2, expand=True, fillcolor=(122, 128, 136)), (110, 110))
    zemin.save(hedef)
    return hedef


if __name__ == "__main__":
    print(ciz(Path(__file__).with_name("z_ornek.png")))
