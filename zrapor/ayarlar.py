"""Değişebilecek sabitler ve `.env` yükleyici.

Model adları burada tek yerde durur; yeni bir model çıkınca yalnız bu dosya değişir.
"""
from __future__ import annotations

import os
from pathlib import Path

# Kısa ad -> Anthropic API model kimliği. `--model` bir kısa ad ya da tam model kimliği alır.
API_MODELLERI: dict[str, str] = {"sonnet": "claude-sonnet-5-5", "opus": "claude-opus-5-5"}
VARSAYILAN_MODEL = "sonnet"
VARSAYILAN_ARKA_UC = "api"
API_MAKS_TOKEN = 16000

# Kabaca maliyet tahmini (README'deki not): bir okuma = 1 tam görsel + 9 büyütülmüş parça.
# Görsel başına ~1.600 giriş token'ı -> okuma başına ~16-17 bin giriş, ~2-3 bin çıkış token'ı.
TAHMINI_GIRIS_TOKEN_OKUMA = 17_000
TAHMINI_CIKIS_TOKEN_OKUMA = 3_000

VARSAYILAN_PORT = 8765
VARSAYILAN_CALISMA = Path("calisma")


def api_modeli(ad: str) -> str:
    """Kısa adı ('sonnet') tam model kimliğine çevirir; tanınmayan ad olduğu gibi döner."""
    return API_MODELLERI.get(ad, ad)


def env_yukle(yol: Path | str = ".env") -> None:
    """`.env` dosyasındaki ANAHTAR=DEĞER satırlarını ortam değişkenlerine ekler (var olanı ezmez)."""
    p = Path(yol)
    if not p.is_file():
        return
    for satir in p.read_text(encoding="utf-8").splitlines():
        satir = satir.strip()
        if not satir or satir.startswith("#") or "=" not in satir:
            continue
        anahtar, deger = satir.split("=", 1)
        deger = deger.strip().strip('"').strip("'")
        if deger:
            os.environ.setdefault(anahtar.strip(), deger)
