"""Masraf fişi fotoğrafları -> "MASRAF LİSTESİ" Excel şablonu (`sablonlar/masraf_listesi.xlsx`).

Şablon: ana sayfa "MASRAF LİSTESİ" satır 5-84 (85 = toplam), kolonlar
B Sıra | C Tarih | D Fiş no | E Satıcı | F/I/L = %1/%10/%20 KDV TUTARI; matrah, toplam, MATRAH, GENEL
formülle (matrah = KDV / oran). "AKARYAKIT MASRAF LİSTESİ" satır 7-86, yalnız %20 (F).
"MASRAF LİSTESİ 8-18" eski oranlı (%8/%18) fişler için. Yalnız C/D/E ve KDV hücreleri doldurulur;
formül, biçim ve yazdırma alanı şablondan gelir.
"""
from __future__ import annotations

import json
import re
from copy import copy
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

from .ayarlar import VARSAYILAN_ARKA_UC, VARSAYILAN_MODEL
from .kontrol import _norm, _t, _tarih, tl
from .okuma import oku_ham

SABLON = Path(__file__).parent / "sablonlar" / "masraf_listesi.xlsx"
ANA, AKARYAKIT, ESKI = "MASRAF LİSTESİ", "AKARYAKIT MASRAF LİSTESİ", "MASRAF LİSTESİ 8-18"
ANA_ILK, ANA_SON = 5, 84
AKY_ILK, AKY_SON = 7, 86
ORAN_KOLON = {1: "F", 10: "I", 20: "L"}
ESKI_ORAN_KOLON = {1: "F", 8: "I", 18: "L"}  # sablonun eski oran sayfasi
KIRMIZI = PatternFill("solid", fgColor="FFC7CE")
SARI = PatternFill("solid", fgColor="FFEB9C")

NUM = {"type": ["number", "null"]}
STR = {"type": ["string", "null"]}
BELGE = {
    "type": "object",
    "properties": {
        "belge_turu": {"type": "string", "description": "perakende satis fisi | e-arsiv fatura | fatura | akaryakit fisi"},
        "satici_unvan": STR, "satici_vkn": STR, "tarih": {"type": ["string", "null"], "description": "GG.AA.YYYY"},
        "saat": STR, "fis_no": {"type": ["string", "null"], "description": "FIS NO / fatura no, fisteki gibi"},
        "akaryakit": {"type": "boolean", "description": "benzin/motorin/LPG satisi mi"},
        "kdv_satirlari": {"type": "array", "items": {"type": "object", "properties": {
            "oran": NUM, "kdv": NUM, "toplam": {"type": ["number", "null"], "description": "o oranin KDV dahil tutari"}},
            "required": ["oran", "kdv", "toplam"]}},
        "toplam_kdv": NUM, "genel_toplam": NUM,
        "odeme": {"type": ["string", "null"], "description": "NAKIT | KREDI KARTI | ..."},
        "emin_olunmayan_alanlar": {"type": "array", "items": {"type": "string"}},
        "tam_gorunuyor": {"type": "boolean"},
    },
    "required": ["belge_turu", "satici_unvan", "tarih", "fis_no", "akaryakit", "kdv_satirlari",
                 "toplam_kdv", "genel_toplam", "emin_olunmayan_alanlar", "tam_gorunuyor"],
}
SCHEMA = {"type": "object", "properties": {
    "belgeler": {"type": "array", "items": BELGE},
    "masraf_olmayan": {"type": "array", "description": "Masraf belgesi OLMAYANLAR (POS slibi, Z raporu, bilgi fisi...)",
                       "items": {"type": "object", "properties": {"tur": {"type": "string"}, "satici": STR},
                                 "required": ["tur"]}}},
    "required": ["belgeler", "masraf_olmayan"]}

TALIMAT = """Sen bir mali musavir burosu icin MASRAF FISLERINI okuyorsun (isletmenin alislari: market,
akaryakit, restoran, nalbur... fisleri ve e-arsiv faturalar). Okudugun rakamlar KDV indirimine girecek:
TAHMIN ETME, gormedigini null yaz ve alan adini emin_olunmayan_alanlar'a ekle ("kdv_satirlari: %10 KDV
son hanesi kirpik"). Bir degeri baska fislerden ya da "olmasi gereken"den doldurma; capraz kontrolu kodumuz yapar.
Termal baskida 0/8, 6/0, 5/9, 1/7 karisir: her haneye dikkat et.

Gorseller: ilki fotografin tamami, digerleri ayni fotografin buyutulmus, ust uste binen parcalari. Fisleri
tam fotograftan tani, rakamlari parcalardan oku. Her belgeyi TAM BIR KEZ dondur. Fotograf yan ya da ters
cekilmis olabilir; once fisin yonunu bul.

Belge basina:
- satici_unvan: fisin en ustundeki isletme adi (kisa ve okunur: "A101", "MİGROS", "SHELL PETROL",
  "ÖRNEK MARKET"); satici_vkn: VKN/TCKN (VD satirindaki numara).
- tarih GG.AA.YYYY, fis_no: "FIS NO" (faturada fatura no). Sifirlari fisteki gibi koru.
- kdv_satirlari: fisin altindaki KDV dokumunden HER oran icin ayri satir: oran (1, 10, 20...), kdv = o
  oranin KDV tutari (TOPKDV / KDV satiri), toplam = o oranin KDV dahil tutari (gorunuyorsa, yoksa null).
  Tek oranli fiste "TOPKDV" o oranin KDV'sidir. KDV'siz (%0) kalem varsa oran 0 ile yaz.
- toplam_kdv = TOPKDV (tum oranlar), genel_toplam = TOPLAM (odenen tutar).
- akaryakit = benzin/motorin/LPG/akaryakit istasyonu satisi ise true.

- TARIH: yili fiste yazdigi gibi oku. Bugunun tarihinden, fatura numarasindaki yildan ya da KDV
  oranlarindan yil TAHMIN ETME; okunmuyorsa null.
- "E-ARSIV FATURA BILGI FISI" (ustunde "TUR: E-ARSIV FATURA" yazan market fisi) GECERLI bir masraf
  belgesidir: belge_turu "e-arsiv fatura" olarak belgeler'e yaz.

MASRAF BELGESI OLMAYANLARI belgeler'e KOYMA, masraf_olmayan'a yaz: kredi karti POS slibi (yalniz kart
bilgisi ve tutar olan banka fisi), Z raporu / X raporu, iade fisi, siparis/adisyon kagidi.
Ayni alisverisin hem fisi hem POS slibi varsa yalniz fisi yaz.

Sayilar: Turk bicimi "*1.326,47" -> 1326.47 (nokta binlik, virgul ondalik, * isaretini at)."""


def oku(foto: Path, calisma: Path, model: str = VARSAYILAN_MODEL, arka_uc: str = VARSAYILAN_ARKA_UC,
        yeniden: bool = False, deneme: int = 1) -> dict:
    """Bir fotoğraftaki masraf belgeleri (önbellekli)."""
    return belgelere(oku_ham(foto, calisma, SCHEMA, TALIMAT, model, arka_uc, yeniden, deneme, onek="masraf_"),
                     foto.name)


def belgelere(sonuc: dict, kaynak: str) -> dict:
    """Ham model çıktısındaki her belgeye kaynak fotoğraf adını ekler."""
    sonuc = dict(sonuc, _kaynak=kaynak)
    sonuc["belgeler"] = [dict(b, _kaynak=kaynak) for b in sonuc.get("belgeler", [])]
    return sonuc


# --- kontrol ----------------------------------------------------------------------------
def _anahtar(b) -> tuple:
    """Tekrar tespiti için (satıcı, tarih, fiş no) anahtarı."""
    return (_norm(b.get("satici_unvan"))[:12], b.get("tarih"), _norm(b.get("fis_no")).lstrip("0"))


def belge_kontrol(b: dict) -> list[dict]:
    """Tek masraf belgesinin kontrolleri: zorunlu alanlar, KDV formülü, KDV ve genel toplam tutarlılığı."""
    s = []

    def ekle(alan, seviye, mesaj):
        s.append({"alan": alan, "seviye": seviye, "mesaj": mesaj})

    for n in b.get("emin_olunmayan_alanlar") or []:
        ekle("not", "uyari", f"okuma notu: {n}")
    if b.get("tam_gorunuyor") is False:
        ekle("not", "uyari", "fisin tamami fotografta degil")
    if not _tarih(b.get("tarih")):
        ekle("tarih", "hata", "tarih okunamadi")
    if not b.get("satici_unvan"):
        ekle("satici", "hata", "satici okunamadi")
    if not b.get("fis_no"):
        ekle("fis_no", "uyari", "fis no okunamadi")
    ks = b.get("kdv_satirlari") or []
    # tek oranli fiste eksik KDV'yi TOPKDV'den, eksik toplami genel toplamdan tamamla (ayni fis, kesin)
    if len(ks) == 1:
        if ks[0].get("kdv") is None and b.get("toplam_kdv") is not None:
            ks[0]["kdv"] = b["toplam_kdv"]
        if ks[0].get("toplam") is None and b.get("genel_toplam") is not None:
            ks[0]["toplam"] = b["genel_toplam"]
    if not ks:
        ekle("kdv", "hata", "KDV dokumu okunamadi")
    for k in ks:
        o = _t(k.get("oran"))
        if k.get("kdv") is None:
            ekle(f"kdv_{o:g}", "hata", f"%{o:g} KDV okunamadi")
            continue
        if o and int(o) not in ORAN_KOLON and int(o) not in ESKI_ORAN_KOLON:
            ekle(f"kdv_{o:g}", "hata", f"%{o:g} oranı şablonda yok (%1/%10/%20, eski %8/%18); elle girin")
        if k.get("toplam") is not None and o:
            beklenen = _t(k["toplam"]) * o / (100 + o)
            if abs(beklenen - _t(k["kdv"])) > max(0.05, 0.003 * _t(k["kdv"])):
                ekle(f"kdv_{o:g}", "hata", f"%{o:g} KDV {tl(k['kdv'])}, toplamdan hesaplanan {tl(beklenen)}")
    if ks and b.get("toplam_kdv") is not None and None not in [k.get("kdv") for k in ks] and \
            abs(sum(_t(k["kdv"]) for k in ks) - _t(b["toplam_kdv"])) > 0.02:
        ekle("kdv", "hata", f"oran KDV'leri toplamı {tl(sum(_t(k['kdv']) for k in ks))} ≠ TOPKDV {tl(b['toplam_kdv'])}")
    if ks and b.get("genel_toplam") is not None and None not in [k.get("toplam") for k in ks] and \
            abs(sum(_t(k["toplam"]) for k in ks) - _t(b["genel_toplam"])) > 0.05:
        ekle("toplam", "uyari", f"oran toplamları {tl(sum(_t(k['toplam']) for k in ks))} ≠ fiş toplamı "
                                f"{tl(b['genel_toplam'])} (indirim/KDV'siz kalem olabilir)")
    # sablon matrahi KDV/oran ile hesaplar; fisin gercek toplamindan sapma kurus seviyesinde kalmali
    if ks and b.get("genel_toplam") is not None and None not in [k.get("kdv") for k in ks]:
        sablon = sum(_t(k["kdv"]) / (_t(k["oran"]) / 100) * (1 + _t(k["oran"]) / 100)
                     for k in ks if _t(k.get("oran")) in (1, 8, 10, 18, 20))
        sifir = sum(_t(k.get("toplam")) for k in ks if not _t(k.get("oran")))
        if abs(sablon + sifir - _t(b["genel_toplam"])) > max(1.0, 0.01 * _t(b["genel_toplam"])):
            ekle("toplam", "uyari", f"şablonun KDV'den hesapladığı toplam {tl(sablon)}, fiş toplamı "
                                    f"{tl(b['genel_toplam'])}")
    return s


def _benzerlik(x: dict, y: dict) -> int:
    """Iki okumadaki belgenin ayni fis olma puani (0 = degil)."""
    sx, sy = _norm(x.get("satici_unvan")), _norm(y.get("satici_unvan"))
    if not sx or not sy:
        return 0
    kisa = min(len(sx), len(sy), 12)  # "ÖRNEK MARKET" / "ÖRNEK MARKET (ÖRNEK GIDA...)" aynı satıcı
    ad = max(SequenceMatcher(None, sx[:kisa], sy[:kisa]).ratio(), SequenceMatcher(None, sx, sy).ratio())
    tx, ty = x.get("genel_toplam"), y.get("genel_toplam")
    ayni_tutar = tx is not None and ty is not None and _t(tx) > 0 and abs(_t(tx) - _t(ty)) < 0.01
    # tek harf okuma farkı (0.8); aynı kuruşuna tutarda daha zayıf ad benzerliği (0.6) de aynı fiş
    if not (sx.startswith(sy) or sy.startswith(sx) or ad >= 0.8 or (ayni_tutar and ad >= 0.6)):
        return 0
    puan = 1
    if tx is not None and ty is not None:
        if abs(_t(tx) - _t(ty)) < 0.01:
            puan += 3
        else:
            return 0  # ayni satici, farkli tutar: ayri fis
    if x.get("tarih") and x.get("tarih") == y.get("tarih"):
        puan += 2
    fx, fy = _norm(x.get("fis_no")).lstrip("0"), _norm(y.get("fis_no")).lstrip("0")
    if fx and fx == fy:
        puan += 2
    elif fx and fy and len(fx) == len(fy) and sum(a != b for a, b in zip(fx, fy)) == 1:
        puan += 1  # tek hane okuma farki (8839 / 0039)
    return puan


def birlestir(okumalar: list[dict]) -> tuple[list[dict], list[str]]:
    """Ayni fotografin bagimsiz okumalari: ilki esas. Eslesen belgede bos alanlar digerinden dolar,
    farkli KDV/toplam/tarih uyari olur; yalniz bir okumada gorulen belge uyari olur."""
    esas = okumalar[0]
    belgeler = [dict(b, _farklar=[]) for b in esas.get("belgeler", [])]
    olmayan = [f"{x.get('tur')}" + (f" ({x['satici']})" if x.get("satici") else "")
               for x in esas.get("masraf_olmayan") or []]
    for diger in okumalar[1:]:
        kalan = list(diger.get("belgeler", []))
        # en yuksek puanli ciftlerden baslayarak eslestir
        ciftler = sorted(((_benzerlik(b, e), i, j) for i, b in enumerate(belgeler) for j, e in enumerate(kalan)),
                         reverse=True)
        eslesen_b, eslesen_e = {}, set()
        for puan, i, j in ciftler:
            if puan >= 2 and i not in eslesen_b and j not in eslesen_e:
                eslesen_b[i] = j
                eslesen_e.add(j)
        # ikinci tur: yalniz satici tutuyor (birinin tutari/tarihi okunamamis) ve iki tarafta da TEK aday
        zayif = [(i, j) for puan, i, j in ciftler if puan == 1 and i not in eslesen_b and j not in eslesen_e]
        for i, j in zayif:
            if sum(1 for a, _ in zayif if a == i) == 1 and sum(1 for _, b in zayif if b == j) == 1:
                eslesen_b[i] = j
                eslesen_e.add(j)
        for i, b in enumerate(belgeler):
            if i not in eslesen_b:
                b["_farklar"].append({"alan": "not", "seviye": "uyari", "mesaj": "bu fiş yalnız bir okumada bulundu"})
                continue
            e = kalan[eslesen_b[i]]
            for alan in ("tarih", "fis_no", "genel_toplam", "toplam_kdv", "satici_vkn"):
                x, y = b.get(alan), e.get(alan)
                if x in (None, "") and y not in (None, ""):
                    b[alan] = y  # ilk okuma goremedi, ikinci gordu
                elif y not in (None, "") and x != y and not (isinstance(x, (int, float)) and
                                                           isinstance(y, (int, float)) and abs(x - y) < 0.01):
                    if alan != "satici_vkn":
                        b["_farklar"].append({"alan": "kdv" if alan == "toplam_kdv" else
                                              ("toplam" if alan == "genel_toplam" else alan), "seviye": "uyari",
                                              "mesaj": f"iki okuma farkli: {alan} = {x} / {y}"})
            if len(_norm(e.get("satici_unvan"))) > len(_norm(b.get("satici_unvan"))):
                b["satici_unvan"] = e["satici_unvan"]  # daha uzun (tam) unvan
            k1 = {_t(k.get("oran")): k for k in b.get("kdv_satirlari") or []}
            k2 = {_t(k.get("oran")): k for k in e.get("kdv_satirlari") or []}
            if not k1 and k2:
                b["kdv_satirlari"] = [dict(k) for k in e["kdv_satirlari"]]
                continue
            for o in set(k1) | set(k2):
                v1 = (k1.get(o) or {}).get("kdv")
                v2 = (k2.get(o) or {}).get("kdv")
                if o not in k1 or o not in k2:  # oran yalniz bir okumada: ekleme (toplam ikilenir), isaretle
                    b["_farklar"].append({"alan": f"kdv_{o:g}", "seviye": "uyari",
                                          "mesaj": f"iki okuma farkli: %{o:g} KDV satırı yalnız bir okumada"
                                                   f" ({'ilk' if o in k1 else 'ikinci'}: {v1 if o in k1 else v2})"})
                elif v1 is None and v2 is not None:
                    k1[o]["kdv"] = v2
                elif v2 is not None and abs(_t(v1) - _t(v2)) >= 0.01:
                    b["_farklar"].append({"alan": f"kdv_{o:g}", "seviye": "uyari",
                                          "mesaj": f"iki okuma farkli: %{o:g} KDV = {v1} / {v2}"})
        for j, x in enumerate(kalan):
            if j not in eslesen_e:
                belgeler.append(dict(x, _farklar=[{"alan": "not", "seviye": "uyari",
                                                   "mesaj": "bu fiş yalnız bir okumada bulundu"}]))
    return belgeler, olmayan


def kontrol_et(belgeler: list[dict]) -> list[dict]:
    """Belge kontrolleri, tekrar ayıklama, ay dışı tarih uyarısı; tarihe göre sıralı liste döndürür."""
    for b in belgeler:
        b["_sorunlar"] = belge_kontrol(b) + b.pop("_farklar", [])
    gorulen = {}
    for b in belgeler:  # ayni fis iki fotografta / iki kez
        k = _anahtar(b)
        if all(k) and k in gorulen:
            b["_sorunlar"].append({"alan": "not", "seviye": "hata",
                                   "mesaj": f"aynı fiş iki kez var ({gorulen[k]['_kaynak']} / {b['_kaynak']}) — biri listeden çıkarıldı"})
            b["_tekrar"] = True
        else:
            gorulen[k] = b
    aylar = [(_tarih(b.get("tarih")).year, _tarih(b.get("tarih")).month) for b in belgeler if _tarih(b.get("tarih"))]
    if aylar:
        cogunluk = max(set(aylar), key=aylar.count)
        for b in belgeler:
            d = _tarih(b.get("tarih"))
            if d and (d.year, d.month) != cogunluk:
                b["_sorunlar"].append({"alan": "tarih", "seviye": "uyari",
                                       "mesaj": f"tarih ({b['tarih']}) listenin ayı dışında"})
    belgeler.sort(key=lambda b: (_tarih(b.get("tarih")) or date.max, _norm(b.get("satici_unvan"))))
    return belgeler


# --- excel ------------------------------------------------------------------------------
def _isaretle(c, sorunlar) -> None:
    """Hücreyi en ağır sorun rengine boyar, mesajları nota yazar."""
    if not sorunlar:
        return
    c.fill = copy(KIRMIZI if any(s["seviye"] == "hata" for s in sorunlar) else SARI)
    c.comment = Comment("\n".join(s["mesaj"] for s in sorunlar), "z-rapor-okuyucu")


def _alan_hucre(alan: str, kolonlar: dict) -> str:
    if alan == "tarih":
        return "C"
    if alan == "fis_no":
        return "D"
    if alan.startswith("kdv_"):
        try:
            return kolonlar.get(int(float(alan[4:])), "E")
        except ValueError:
            return "E"
    return "E"  # satici, toplam, not -> satici hucresi


def yaz(belgeler: list[dict], cikti: Path, hazirlayan: str = "") -> dict:
    """Belgeleri şablona yerleştirir: akaryakıt ayrı sayfaya, %8/%18 eski oran sayfasına."""
    wb = load_workbook(SABLON)
    ana, aky, eski = wb[ANA], wb[AKARYAKIT], wb[ESKI]
    if hazirlayan:
        ana["D2"] = hazirlayan
        aky["D3"] = hazirlayan
        eski["C2"] = f"HAZIRLAYAN: {hazirlayan}"
    satir = {"ana": ANA_ILK, "aky": AKY_ILK, "eski": ANA_ILK}
    tasan, sayac = [], {"ana": 0, "aky": 0, "eski": 0}

    def yerlestir(ws, tur, b, oranlar_kolon):
        if satir[tur] > (AKY_SON if tur == "aky" else ANA_SON):
            tasan.append(b)
            return
        r = satir[tur]
        satir[tur] += 1
        sayac[tur] += 1
        d = _tarih(b.get("tarih"))
        ws[f"C{r}"] = d or b.get("tarih")
        ws[f"D{r}"] = b.get("fis_no")
        ws[f"E{r}"] = (b.get("satici_unvan") or "").upper()
        for o, kol in oranlar_kolon.items():
            if o in b["_kdv"]:
                ws[f"{kol}{r}"] = round(_t(b["_kdv"][o]), 2)
        hucreler: dict[str, list] = {}
        for s in b["_sorunlar"]:
            hucreler.setdefault(_alan_hucre(s["alan"], oranlar_kolon), []).append(s)
        for kol, ss in hucreler.items():
            _isaretle(ws[f"{kol}{r}"], ss)

    for b in belgeler:
        if b.get("_tekrar"):
            continue
        b["_kdv"] = {}
        for k in b.get("kdv_satirlari") or []:
            o = int(round(_t(k.get("oran"))))
            if o and k.get("kdv") is not None:
                b["_kdv"][o] = b["_kdv"].get(o, 0) + _t(k["kdv"])
        if b.get("akaryakit"):
            yerlestir(aky, "aky", b, {20: "F"})
            diger = {o: v for o, v in b["_kdv"].items() if o != 20}
            if diger:  # istasyon fisinde market kalemi: %1/%10 kismi ana listeye
                b2 = dict(b, _kdv=diger, _sorunlar=b["_sorunlar"] + [
                    {"alan": "not", "seviye": "uyari", "mesaj": "akaryakıt fişinin %20 dışı kalemi (akaryakıt listesinde de var)"}])
                yerlestir(ana, "ana", b2, ORAN_KOLON)
        elif any(o in (8, 18) for o in b["_kdv"]):
            yeni_oran = {o: v for o, v in b["_kdv"].items() if o in ORAN_KOLON and o != 1}
            yerlestir(eski, "eski", dict(b, _kdv={o: v for o, v in b["_kdv"].items() if o in ESKI_ORAN_KOLON}),
                      ESKI_ORAN_KOLON)
            if yeni_oran:  # ayni fiste %10/%20 de varsa o kisim ana listeye
                yerlestir(ana, "ana", dict(b, _kdv=yeni_oran), ORAN_KOLON)
        else:
            yerlestir(ana, "ana", b, ORAN_KOLON)
    wb.save(cikti)
    hata = sum(1 for b in belgeler if any(s["seviye"] == "hata" for s in b["_sorunlar"]))
    uyari = sum(1 for b in belgeler if b["_sorunlar"]) - hata
    return {"fis": sayac["ana"], "akaryakit": sayac["aky"], "eski_oran": sayac["eski"], "hata": hata, "uyari": uyari,
            "tekrar": sum(1 for b in belgeler if b.get("_tekrar")),
            "tasan": [f"{b.get('tarih')} {b.get('satici_unvan')}" for b in tasan]}


def isle(liste: list[Path], cikti: Path, model: str = VARSAYILAN_MODEL, arka_uc: str = VARSAYILAN_ARKA_UC,
         okuma: int = 2, paralel: int = 3, yeniden: bool = False, hazirlayan: str = "", bildir=print,
         calisma: Path | None = None) -> dict:
    """Masraf fotoğraflarını oku, birleştir, kontrol et, masraf listesine yaz; özet döndürür."""
    from concurrent.futures import ThreadPoolExecutor, as_completed

    calisma = calisma or cikti.parent / "_okumalar"
    calisma.mkdir(parents=True, exist_ok=True)
    okumalar: dict[Path, dict[int, dict]] = {f: {} for f in liste}
    toplam, biten = len(liste) * okuma, 0
    with ThreadPoolExecutor(max(1, paralel)) as ex:
        isler = {ex.submit(oku, f, calisma, model, arka_uc, yeniden, n): (f, n)
                 for f in liste for n in range(1, okuma + 1)}
        for i in as_completed(isler):
            f, n = isler[i]
            biten += 1
            try:
                s = i.result()
                okumalar[f][n] = s
                bildir(f"[{biten}/{toplam}] {f.name} (okuma {n}): {len(s.get('belgeler', []))} masraf fişi")
            except Exception as e:
                bildir(f"[{biten}/{toplam}] {f.name} (okuma {n}): OKUNAMADI — {e}")
    return okumalardan_yaz({f.name: oks for f, oks in okumalar.items()}, cikti, hazirlayan, bildir)


def okumalardan_yaz(okumalar: dict[str, dict[int, dict]], cikti: Path, hazirlayan: str = "",
                    bildir=print) -> dict:
    """Fotoğraf adı -> {okuma no: okuma}: birleştirir, kontrol eder, masraf listesini yazar; özet döndürür."""
    belgeler, olmayan = [], []
    for ad, oks in okumalar.items():
        if not oks:
            bildir(f"{ad}: hiç okunamadı, listede yok!")
            continue
        b, o = birlestir([oks[n] for n in sorted(oks)])
        belgeler += b
        if o:
            olmayan.append(f"{ad}: {', '.join(o)}")
            bildir(f"{ad}: masraf belgesi OLMAYAN {len(o)} belge atlandı — {', '.join(o)}")
    if not belgeler:
        raise RuntimeError("Bu fotoğraflarda masraf fişi bulunamadı."
                           + (" Görülen belgeler: " + " | ".join(olmayan) if olmayan else ""))
    belgeler = kontrol_et(belgeler)
    ozet = yaz(belgeler, cikti, hazirlayan)
    (cikti.parent / "masraf.json").write_text(json.dumps(belgeler, ensure_ascii=False, default=str), encoding="utf-8")
    tarihler = sorted(d for d in (_tarih(b.get("tarih")) for b in belgeler) if d)
    ozet |= {"ilk": tarihler[0].strftime("%d.%m.%Y") if tarihler else None,
             "son": tarihler[-1].strftime("%d.%m.%Y") if tarihler else None, "olmayan": olmayan}
    bildir(f"{ozet['fis']} masraf + {ozet['akaryakit']} akaryakıt + {ozet['eski_oran']} eski oranlı fiş · {ozet['hata']} hatalı · {ozet['uyari']} uyarılı"
           + (f" · {ozet['tekrar']} tekrar çıkarıldı" if ozet["tekrar"] else ""))
    if ozet["tasan"]:
        bildir(f"UYARI: şablon 80 satır; {len(ozet['tasan'])} fiş sığmadı: {', '.join(ozet['tasan'][:5])}")
    return ozet
