"""Kontrol fonksiyonları ve çoklu okuma birleştirme."""
from __future__ import annotations

import copy

from conftest import hatalar

from zrapor.kontrol import birlestir, kontrol_et, tek_rapor


def _tek_firma(firmalar: dict) -> list[dict]:
    assert len(firmalar) == 1
    return next(iter(firmalar.values()))


def test_dengeli_rapor_hatasiz(dengeli):
    assert [s for s in tek_rapor(dengeli) if s["seviye"] == "hata"] == []


def test_departman_toplami_bozuk_tek_hata(dengeli):
    dengeli["departmanlar"][0]["tutar"] += 500
    sorun = [s for s in tek_rapor(dengeli) if s["seviye"] == "hata"]
    assert len(sorun) == 1
    assert sorun[0]["alan"] == "net_satis" and "departman toplami" in sorun[0]["mesaj"]


def test_kdv_formulu_bozuk(dengeli):
    dengeli["kdv_dokumu"][0]["kdv"] = 700.00  # %10 için 737,27 olmalı
    alanlar = {s["alan"] for s in tek_rapor(dengeli) if s["seviye"] == "hata"}
    assert "kdv_10" in alanlar


def test_odemeler_net_satisi_tutmuyor(dengeli):
    dengeli["odemeler"][0]["tutar"] += 100
    alanlar = {s["alan"] for s in tek_rapor(dengeli) if s["seviye"] == "hata"}
    assert "odemeler" in alanlar


def test_ornek_veri_bir_kirmizi_bir_mavi(ornek_raporlar):
    rs = _tek_firma(kontrol_et(ornek_raporlar))
    assert [r["z_no"] for r in rs] == ["0101", "0102", "0103", "0104", "0105"]
    assert sum(len(hatalar(r)) for r in rs) == 1
    assert hatalar(rs[2])[0]["alan"] == "net_satis"


def test_kumulatif_zincir_duzeltmesi_mavi(ornek_raporlar):
    beklenen = round(ornek_raporlar[2]["kum_toplam"] + ornek_raporlar[3]["net_satis"], 2)
    assert ornek_raporlar[3]["kum_toplam"] != beklenen  # örnekte bilerek yanlış okunmuş
    rs = _tek_firma(kontrol_et(ornek_raporlar))
    z4 = rs[3]
    duzeltme = hatalar(z4, "duzeltme")
    assert len(duzeltme) == 1 and duzeltme[0]["alan"] == "kum_toplam"
    assert z4["kum_toplam"] == beklenen
    assert hatalar(z4) == []


def test_z_no_atlamasi(ornek_raporlar):
    del ornek_raporlar[1]  # Z 0102 yok
    rs = _tek_firma(kontrol_et(ornek_raporlar))
    z3 = next(r for r in rs if r["z_no"] == "0103")
    mesajlar = [s["mesaj"] for s in z3["_sorunlar"] if s["alan"] == "z_no"]
    assert any("Z 102 fotografta yok" in m for m in mesajlar)
    # eksik Z'nin cirosu kümülatiften hesaplanıp söylenir
    assert any("cirosu 10.860,50" in m for m in mesajlar)


def test_tarih_boslugu_uyari(ornek_raporlar):
    ornek_raporlar[4]["tarih"] = "07.09.2026"
    rs = _tek_firma(kontrol_et(ornek_raporlar))
    assert any(s["alan"] == "tarih" and "gun Z yok" in s["mesaj"] for s in hatalar(rs[4], "uyari"))


def test_birlestir_iki_okuma_farkli_uyari(dengeli):
    ikinci = copy.deepcopy(dengeli)
    ikinci["toplam_fis_adedi"] = 77
    sonuc = birlestir([[dengeli], [ikinci]])
    assert len(sonuc) == 1
    farklar = sonuc[0]["_okuma_farklari"]
    assert any(f["alan"] == "toplam_fis_adedi" and f["seviye"] == "uyari"
               and f["mesaj"].startswith("iki okuma farkli") for f in farklar)


def test_birlestir_ilk_okumanin_bosunu_ikinciden_alir(dengeli):
    ikinci = copy.deepcopy(dengeli)
    dengeli["kasa_nakit"] = None
    sonuc = birlestir([[dengeli], [ikinci]])
    assert sonuc[0]["kasa_nakit"] == ikinci["kasa_nakit"]


def test_birlestir_yalniz_bir_okumada_bulunan(ornek_raporlar):
    sonuc = birlestir([ornek_raporlar[:2], ornek_raporlar[:1]])
    z2 = next(r for r in sonuc if r["z_no"] == "0102")
    assert any("yalniz bir okumada" in f["mesaj"] for f in z2["_okuma_farklari"])
