"""Gemini arka ucu (sahte httpx yanıtı) ve karma mod dağılımı."""
from __future__ import annotations

import base64
import json
from pathlib import Path

import httpx
import pytest

from conftest import ORNEKLER

from zrapor import masraf, okuma
from zrapor.ayarlar import GEMINI_MODEL
from zrapor.isleme import isle
from zrapor.model import SEMA


def _rapor() -> dict:
    r = json.loads((ORNEKLER / "okuma_ornek.json").read_text(encoding="utf-8"))["raporlar"][0]
    return {k: v for k, v in r.items() if not k.startswith("_")}


def test_gemini_semasi_nullable():
    s = okuma.gemini_semasi(SEMA)
    vkn = s["properties"]["raporlar"]["items"]["properties"]["vkn"]
    assert vkn["type"] == "STRING" and vkn["nullable"] is True
    assert s["properties"]["raporlar"]["type"] == "ARRAY"
    assert "kum_toplam" in s["properties"]["raporlar"]["items"]["required"]


def test_gemini_sahte_yanit(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("GEMINI_API_KEY", "test-anahtari")
    giden: list = []

    def sahte_post(url, json=None, headers=None, timeout=None):
        giden.append((url, json, headers))
        cevap = {"candidates": [{"content": {"parts": [{"text": _json_metin()}]}, "finishReason": "STOP"}]}
        return httpx.Response(200, json=cevap, request=httpx.Request("POST", url))

    def _json_metin():
        return json.dumps({"raporlar": [_rapor()], "z_raporu_olmayan": []})

    monkeypatch.setattr(httpx, "post", sahte_post)
    rs = okuma.oku(ORNEKLER / "z_ornek.png", tmp_path, arka_uc="gemini")
    assert len(rs) == 1 and rs[0]["z_no"] == "0101" and rs[0]["_kaynak"] == "z_ornek.png"
    url, govde, basliklar = giden[0]
    assert url.endswith(f"/models/{GEMINI_MODEL}:generateContent")
    assert basliklar["x-goog-api-key"] == "test-anahtari" and "key=" not in url
    parcalar = govde["contents"][0]["parts"]
    assert len(parcalar) == 2  # tam fotoğraf (tek görsel) + talimat
    assert parcalar[0]["inline_data"]["mime_type"] == "image/jpeg"
    assert base64.b64decode(parcalar[0]["inline_data"]["data"])[:2] == b"\xff\xd8"
    ayar = govde["generationConfig"]
    assert ayar["response_mime_type"] == "application/json" and ayar["response_schema"]["type"] == "OBJECT"
    # ikinci çağrı önbellekten
    okuma.oku(ORNEKLER / "z_ornek.png", tmp_path, arka_uc="gemini")
    assert len(giden) == 1


def test_gemini_anahtarsiz_hata(tmp_path: Path, monkeypatch):
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    with pytest.raises(RuntimeError, match="GEMINI_API_KEY"):
        okuma.oku(ORNEKLER / "z_ornek.png", tmp_path, arka_uc="gemini")


def _sahteler(monkeypatch, cevap: dict) -> list:
    cagrilar: list = []

    def api(gorseller, model, sema, talimat):
        cagrilar.append(("api", model, len(gorseller)))
        return cevap

    def gemini(gorseller, model, sema, talimat):
        cagrilar.append(("gemini", model, len(gorseller)))
        return cevap

    monkeypatch.setattr(okuma, "_api", api)
    monkeypatch.setattr(okuma, "_gemini", gemini)
    return cagrilar


def test_karma_iki_okuma_farkli_arka_uc(tmp_path: Path, monkeypatch):
    cagrilar = _sahteler(monkeypatch, {"raporlar": [_rapor()], "z_raporu_olmayan": []})
    ozet = isle([ORNEKLER / "z_ornek.png"], tmp_path / "z.xlsx", arka_uc="karma", okuma=2,
                paralel=2, bildir=lambda m: None)
    assert sorted(cagrilar) == [("api", "claude-sonnet-5-5", 10), ("gemini", GEMINI_MODEL, 1)]
    assert ozet[0]["z"] == 1


def test_karma_tek_okuma_yalniz_anthropic(tmp_path: Path, monkeypatch):
    cagrilar = _sahteler(monkeypatch, {"raporlar": [_rapor()], "z_raporu_olmayan": []})
    isle([ORNEKLER / "z_ornek.png"], tmp_path / "z.xlsx", arka_uc="karma", okuma=1, bildir=lambda m: None)
    assert [c[0] for c in cagrilar] == ["api"]


def test_karma_masraf(tmp_path: Path, monkeypatch):
    belge = {"belge_turu": "perakende satis fisi", "satici_unvan": "ÖRNEK MARKET", "tarih": "05.09.2026",
             "fis_no": "0042", "akaryakit": False,
             "kdv_satirlari": [{"oran": 20, "kdv": 20.0, "toplam": 120.0}], "toplam_kdv": 20.0,
             "genel_toplam": 120.0, "emin_olunmayan_alanlar": [], "tam_gorunuyor": True}
    cagrilar = _sahteler(monkeypatch, {"belgeler": [belge], "masraf_olmayan": []})
    o = masraf.isle([ORNEKLER / "z_ornek.png"], tmp_path / "m.xlsx", arka_uc="karma", okuma=2,
                    bildir=lambda m: None)
    assert sorted(c[0] for c in cagrilar) == ["api", "gemini"]
    assert o["fis"] == 1 and o["hata"] == 0
