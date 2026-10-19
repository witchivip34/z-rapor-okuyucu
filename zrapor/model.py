"""Z raporu okumasının JSON şeması.

Görsel dil modeli her fotoğraf için `SEMA`ya uyan bir nesne döndürür: `raporlar` (fotoğraftaki her
Z raporu bir kez) ve `z_raporu_olmayan` (fotoğrafta görülen ama Z raporu olmayan belgeler).
Okunamayan alan `null` yazılır; türetme ve düzeltme modelde değil, `zrapor.kontrol`de yapılır.
"""
from __future__ import annotations

SAYI = {"type": ["number", "null"]}
TAMSAYI = {"type": ["integer", "null"]}
METIN = {"type": ["string", "null"]}


def _aciklamali(tur: dict, aciklama: str) -> dict:
    """Bir alan tipine Türkçe açıklama ekler."""
    return {**tur, "description": aciklama}


RAPOR: dict = {
    "type": "object",
    "properties": {
        "firma_unvan": _aciklamali(METIN, "Fişin üstündeki işletme unvanı"),
        "firma_sahibi": _aciklamali(METIN, "Şahıs işletmesinde sahibin adı"),
        "vergi_dairesi": _aciklamali(METIN, "Vergi dairesi adı"),
        "vkn": _aciklamali(METIN, "Vergi kimlik no ya da TCKN"),
        "tarih": _aciklamali(METIN, "GG.AA.YYYY"),
        "saat": _aciklamali(METIN, "SS:DD"),
        "fis_no": _aciklamali(METIN, "Üstteki FİŞ NO, fişteki gibi sıfırlı"),
        "z_no": _aciklamali(METIN, "Z SAYAÇ / Z NO, fişteki gibi sıfırlı (\"0130\")"),
        "eku_no": _aciklamali(METIN, "EKÜ numarası"),
        "cihaz_sicil_no": _aciklamali(METIN, "Cihaz sicil no (\"AF AV...\" satırı)"),
        "departmanlar": {
            "type": "array",
            "description": "Departman satırları (Yiyecek, İçecek...)",
            "items": {"type": "object", "properties": {
                "ad": METIN, "kdv_orani": SAYI, "adet": TAMSAYI, "tutar": SAYI},
                "required": ["ad", "kdv_orani", "adet", "tutar"]}},
        "net_satis": _aciklamali(SAYI, "NET SATIŞ (günlük ciro, KDV dahil)"),
        "odemeler": {
            "type": "array", "description": "NAKİT, K.KARTI, YEMEK KARTI vb.",
            "items": {"type": "object", "properties": {
                "tur": METIN, "adet": TAMSAYI, "tutar": SAYI},
                "required": ["tur", "adet", "tutar"]}},
        "kasa_nakit": _aciklamali(SAYI, "KASA NAKİT"),
        "banka_dokumu": {
            "type": "array", "description": "Kart tahsilatının banka kırılımı",
            "items": {"type": "object", "properties": {
                "banka": METIN, "adet": TAMSAYI, "tutar": SAYI},
                "required": ["banka", "adet", "tutar"]}},
        "fis_iptal_adet": _aciklamali(TAMSAYI, "FİŞ İPTAL adedi"),
        "fis_iptal_tutar": _aciklamali(SAYI, "FİŞ İPTAL tutarı"),
        "musteri_fis_adedi": TAMSAYI,
        "mali_fis_adedi": TAMSAYI,
        "slip_fis_adedi": TAMSAYI,
        "toplam_fis_adedi": TAMSAYI,
        "kdv_dokumu": {
            "type": "array", "description": "VERGİ DÖKÜMÜ: her %X için TOPLAM ve TOPKDV",
            "items": {"type": "object", "properties": {
                "oran": SAYI, "toplam": SAYI, "kdv": SAYI},
                "required": ["oran", "toplam", "kdv"]}},
        "mali_toplam": _aciklamali(SAYI, "MALİ VERİ TOPLAM"),
        "mali_topkdv": _aciklamali(SAYI, "MALİ VERİ TOPKDV"),
        "kum_toplam": _aciklamali(SAYI, "KÜM. TOPLAM (cihazın ilk günden beri cirosu)"),
        "kum_topkdv": _aciklamali(SAYI, "KÜM. TOPKDV"),
        "diger_satirlar": {
            "type": "array", "description": "Şemada olmayan ama fişte yazan satırlar",
            "items": {"type": "object", "properties": {
                "etiket": {"type": "string"}, "deger": {"type": "string"}},
                "required": ["etiket", "deger"]}},
        "emin_olunmayan_alanlar": {
            "type": "array", "items": {"type": "string"},
            "description": "Okunamayan ya da emin olunmayan alan adları, kısa notla"},
        "tam_gorunuyor": {"type": "boolean", "description": "Fişin tamamı fotoğrafta mı"},
    },
    "required": ["firma_unvan", "vkn", "tarih", "z_no", "departmanlar", "net_satis", "odemeler",
                 "kdv_dokumu", "mali_toplam", "mali_topkdv", "kum_toplam", "kum_topkdv",
                 "emin_olunmayan_alanlar", "tam_gorunuyor"],
}

SEMA: dict = {
    "type": "object",
    "properties": {
        "raporlar": {"type": "array", "items": RAPOR},
        "z_raporu_olmayan": {
            "type": "array", "description": "Fotoğraftaki Z RAPORU OLMAYAN belgeler",
            "items": {"type": "object", "properties": {
                "tur": {"type": "string", "description": "satış fişi, e-arşiv fatura, POS slibi..."},
                "satici": METIN}, "required": ["tur"]}},
    },
    "required": ["raporlar", "z_raporu_olmayan"],
}

RAPOR_ALANLARI: frozenset[str] = frozenset(RAPOR["properties"])
