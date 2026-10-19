"""Ortak test verisi: API'siz, sentetik Z raporları."""
from __future__ import annotations

import copy
import json
from pathlib import Path

import pytest

ORNEKLER = Path(__file__).resolve().parent.parent / "ornekler"


@pytest.fixture
def ornek_raporlar() -> list[dict]:
    """ornekler/okuma_ornek.json'daki 5 günlük okuma (her testte taze kopya)."""
    veri = json.loads((ORNEKLER / "okuma_ornek.json").read_text(encoding="utf-8"))
    return copy.deepcopy(veri["raporlar"])


@pytest.fixture
def dengeli(ornek_raporlar) -> dict:
    """Kendi içinde tamamen tutarlı tek bir Z raporu (1. gün)."""
    return ornek_raporlar[0]


def hatalar(r: dict, seviye: str = "hata") -> list[dict]:
    return [s for s in r["_sorunlar"] if s["seviye"] == seviye]
