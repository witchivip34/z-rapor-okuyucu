"""Web: --host / --guvenilir-vekil parametreleri ve giriş deneme sınırı."""
from __future__ import annotations

from pathlib import Path

import pytest

pytest.importorskip("fastapi")
from fastapi.testclient import TestClient  # noqa: E402

from zrapor import web  # noqa: E402
from zrapor.cli import main  # noqa: E402


def test_cli_web_parametreleri(monkeypatch, tmp_path: Path):
    gelen: dict = {}
    monkeypatch.setattr(web, "calistir", lambda **k: gelen.update(k))
    monkeypatch.delenv("ZRAPOR_ARKA_UC", raising=False)
    main(["web", "--host", "0.0.0.0", "--port", "9000", "--veri", str(tmp_path), "--guvenilir-vekil"])
    assert gelen["host"] == "0.0.0.0" and gelen["port"] == 9000 and gelen["guvenilir_vekil"] is True
    assert gelen["arka_uc"] == "api"


def test_cli_web_varsayilanlar(monkeypatch, tmp_path: Path):
    gelen: dict = {}
    monkeypatch.setattr(web, "calistir", lambda **k: gelen.update(k))
    monkeypatch.setenv("ZRAPOR_ARKA_UC", "karma")
    main(["web", "--veri", str(tmp_path)])
    assert gelen["host"] == "127.0.0.1" and gelen["guvenilir_vekil"] is False
    assert gelen["arka_uc"] == "karma"  # ZRAPOR_ARKA_UC ortamdan


def test_calistir_parolasiz_baslamaz(monkeypatch, tmp_path: Path):
    monkeypatch.setenv("ZRAPOR_SIFRE", "kisa")
    with pytest.raises(SystemExit):
        web.calistir(veri=tmp_path)


@pytest.fixture
def istemci(monkeypatch):
    monkeypatch.setattr(web.AYAR, "sifre", "dogru-parola-123")
    monkeypatch.setattr(web.AYAR, "giris_denemeleri", {})
    return TestClient(web.app)


def _yanlis(istemci, xff: str | None = None) -> int:
    basliklar = {"X-Forwarded-For": xff} if xff else {}
    return istemci.post("/api/giris", data={"sifre": "yanlis-parola"}, headers=basliklar).status_code


def test_guvenilir_vekil_adres_basina_sinir(istemci, monkeypatch):
    monkeypatch.setattr(web.AYAR, "guvenilir_vekil", True)
    assert [_yanlis(istemci, "198.51.100.7") for _ in range(8)] == [403] * 8
    assert _yanlis(istemci, "198.51.100.7") == 429
    # başka istemci etkilenmez; baştaki sahte değer yok sayılır (vekilin eklediği son değer esas)
    assert _yanlis(istemci, "198.51.100.7, 203.0.113.9") == 403
    r = istemci.post("/api/giris", data={"sifre": "dogru-parola-123"}, headers={"X-Forwarded-For": "203.0.113.9"})
    assert r.status_code == 200 and web.CEREZ in r.cookies


def test_vekil_kapaliyken_baslik_yok_sayilir(istemci, monkeypatch):
    monkeypatch.setattr(web.AYAR, "guvenilir_vekil", False)
    for i in range(8):
        assert _yanlis(istemci, f"198.51.100.{i}") == 403
    assert _yanlis(istemci, "203.0.113.50") == 429  # başlık sahtelense de sınır aşılamaz
