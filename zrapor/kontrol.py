"""Z raporu kontrolleri ve çoklu okumaların birleştirilmesi.

Her rapora `_sorunlar` listesi eklenir: `{"alan", "seviye", "mesaj"}`.

seviye:
    "hata"     -> rakamlar birbirini tutmuyor, mutlaka fişe bakılmalı (Excel'de kırmızı)
    "uyari"    -> olağan dışı ama mümkün ya da tahmini (Excel'de sarı)
    "duzeltme" -> okuma hatası komşu Z'lerden kanıtlanıp düzeltildi (Excel'de mavi)

Yedi kontrol grubu:
    1. Ciro dengesi: departman toplamı = net satış = mali veri toplamı = KDV dökümü toplamı
    2. Ödemeler toplamı = net satış
    3. KDV formülü: %X KDV = toplam × X / (100 + X); KDV satırları toplamı = TOPKDV
    4. Banka dökümü toplamı = kredi kartı tahsilatı
    5. Adetler: ödeme adetleri = müşteri fiş adedi; iptal adedi varsa tutarı da olmalı
    6. Kümülatif zincir: bugünkü KÜM. TOPLAM − dünkü = bugünkü ciro (KDV için de)
    7. Z no / tarih ardışıklığı ve fiş adedi aykırılığı
"""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from statistics import median, mode

from .model import RAPOR_ALANLARI

TOL = 0.01       # toplamlar kuruşu kuruşuna tutmalı
KDV_TOL = 0.05   # cihaz KDV'yi kesip yuvarlayabiliyor
KUM_TOL = 0.015  # kümülatif KDV cihazda kuruş altı hassasiyetle tutuluyor, 1 kuruş kayabiliyor

KARSILASTIR = ("tarih", "fis_no", "net_satis", "mali_topkdv", "kum_toplam", "kum_topkdv",
               "fis_iptal_adet", "fis_iptal_tutar", "musteri_fis_adedi", "mali_fis_adedi",
               "slip_fis_adedi", "toplam_fis_adedi", "kasa_nakit")
ZINCIRLI = {"net_satis", "mali_topkdv", "kum_toplam", "kum_topkdv"}  # zincir doğrularsa uyarı düşer


# --- yardımcılar ---------------------------------------------------------------------------
def _t(x) -> float:
    """None'u 0 sayan float çevirici."""
    return float(x or 0)


def _tarih(s) -> date | None:
    """'GG.AA.YYYY' -> date; bozuksa None."""
    try:
        return datetime.strptime(s, "%d.%m.%Y").date()
    except (TypeError, ValueError):
        return None


def _znum(r: dict) -> int | None:
    """Z numarasını tamsayı olarak verir ('0130' -> 130)."""
    try:
        return int(str(r.get("z_no")).strip())
    except (TypeError, ValueError):
        return None


def odeme_turu(tur: str | None) -> str:
    """Fişteki ödeme etiketini 'nakit' | 'kart' | 'diger' sınıfına çevirir."""
    t = (tur or "").upper().replace("İ", "I").replace(".", "").replace(" ", "")
    if "NAKIT" in t:
        return "nakit"
    if t.startswith("KKART") or "KREDI" in t or t == "KART":
        return "kart"
    return "diger"


def odeme(r: dict, tur: str) -> float | None:
    """Bir türdeki (nakit | kart | diger) ödemelerin toplamı; o türde ödeme yoksa None."""
    vals = [o["tutar"] for o in r.get("odemeler") or [] if odeme_turu(o.get("tur")) == tur]
    return sum(_t(v) for v in vals) if vals else None


def tl(x) -> str:
    """Türk para biçimi: 1234.5 -> '1.234,50'."""
    return f"{_t(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _ekle(r: dict, alan: str, seviye: str, mesaj: str) -> None:
    r["_sorunlar"].append({"alan": alan, "seviye": seviye, "mesaj": mesaj})


# --- tek fiş kontrolleri ---------------------------------------------------------------------
def _okuma_kontrolu(r: dict, ekle) -> None:
    """Zorunlu alanlar, modelin okuma notları, kadraj ve tarih biçimi."""
    for alan in ("tarih", "z_no", "net_satis", "mali_topkdv", "kum_toplam", "kum_topkdv"):
        if r.get(alan) in (None, ""):
            ekle(alan, "hata", f"{alan} okunamadi")
    for not_ in r.get("emin_olunmayan_alanlar") or []:
        m = re.match(r"\s*([a-z_]+)", not_)
        ekle(m.group(1) if m and m.group(1) in RAPOR_ALANLARI else "diger", "uyari",
             f"okuma notu: {not_}")
    if r.get("tam_gorunuyor") is False:
        ekle("z_no", "uyari", "fisin tamami fotografta degil")
    if r.get("tarih") and not _tarih(r["tarih"]):
        ekle("tarih", "hata", f"tarih bicimi bozuk: {r['tarih']}")


def _ciro_dengesi(r: dict, net: float, ekle) -> None:
    """Kontrol 1-2: departman toplamı, ödemeler toplamı ve mali veri toplamı = net satış."""
    dep = r.get("departmanlar") or []
    if dep and None not in [d.get("tutar") for d in dep] and \
            abs(sum(_t(d.get("tutar")) for d in dep) - net) > TOL:
        ekle("net_satis", "hata",
             f"departman toplami {tl(sum(_t(d.get('tutar')) for d in dep))} ≠ net satis {tl(net)}")
    od = r.get("odemeler") or []
    if any(o.get("tutar") is None for o in od):
        ekle("odemeler", "hata", "odeme tutari okunamadi")
    elif od and abs(sum(_t(o.get("tutar")) for o in od) - net) > TOL:
        ekle("odemeler", "hata",
             f"odemeler toplami {tl(sum(_t(o.get('tutar')) for o in od))} ≠ net satis {tl(net)}")
    if r.get("mali_toplam") is not None and abs(_t(r["mali_toplam"]) - net) > TOL:
        ekle("mali_toplam", "hata", f"mali veri toplam {tl(r['mali_toplam'])} ≠ net satis {tl(net)}")


def _kdv_kontrolu(r: dict, net: float, ekle) -> None:
    """Kontrol 1 ve 3: KDV dökümü toplamı = net satış, %X formülü, satırlar = TOPKDV."""
    kd = r.get("kdv_dokumu") or []
    if not kd:
        ekle("kdv_dokumu", "hata", "vergi dokumu okunamadi")
        return
    if None not in [k.get("toplam") for k in kd] and \
            abs(sum(_t(k.get("toplam")) for k in kd) - net) > TOL:
        ekle("kdv_dokumu", "hata", "vergi dokumu toplami net satisi tutmuyor")
    for k in kd:
        if k.get("kdv") is None or k.get("toplam") is None:
            continue  # okunamayan alan zaten raporlandı
        o = _t(k.get("oran"))
        beklenen = _t(k.get("toplam")) * o / (100 + o)
        if abs(beklenen - _t(k.get("kdv"))) > KDV_TOL:
            ekle(f"kdv_{o:g}", "hata", f"%{o:g} KDV {tl(k.get('kdv'))}, hesaplanan {tl(beklenen)}")
    if len(kd) > 1 and any(k.get("kdv") is None for k in kd):
        ekle("kdv_dokumu", "hata", "oran KDV'si okunamadi")
    elif r.get("mali_topkdv") is not None and None not in [k.get("kdv") for k in kd] and \
            abs(sum(_t(k.get("kdv")) for k in kd) - _t(r["mali_topkdv"])) > TOL:
        ekle("mali_topkdv", "hata", "KDV satirlari toplami TOPKDV'yi tutmuyor")
    oranlar = {_t(k.get("oran")) for k in kd}
    for d in r.get("departmanlar") or []:
        if d.get("kdv_orani") is not None and _t(d["kdv_orani"]) not in oranlar:
            ekle("departmanlar", "hata", f"{d.get('ad')} %{d['kdv_orani']:g} vergi dokumunde yok")


def _banka_ve_adetler(r: dict, ekle) -> None:
    """Kontrol 4-5: banka dökümü = kart; ödeme adetleri = müşteri fiş adedi; iptal tutarı."""
    kart = odeme(r, "kart")
    banka = r.get("banka_dokumu") or []
    if kart is not None and banka and abs(sum(_t(b.get("tutar")) for b in banka) - kart) > TOL:
        ekle("banka_dokumu", "uyari", "banka dokumu kredi karti toplamini tutmuyor")
    od = r.get("odemeler") or []
    adet = [o.get("adet") for o in od]
    if od and None not in adet and r.get("musteri_fis_adedi") is not None \
            and sum(adet) != r["musteri_fis_adedi"]:
        ekle("musteri_fis_adedi", "uyari",
             f"odeme adetleri toplami {sum(adet)} ≠ musteri fis adedi {r['musteri_fis_adedi']}")
    if r.get("fis_iptal_adet") and r.get("fis_iptal_tutar") in (None, 0):
        ekle("fis_iptal_tutar", "uyari", "iptal adedi var ama tutar yok")


def tek_rapor(r: dict) -> list[dict]:
    """Tek bir Z raporunun kendi içindeki kontrolleri (1-5); bulunan sorunları döndürür."""
    s: list[dict] = []

    def ekle(alan, seviye, mesaj):
        s.append({"alan": alan, "seviye": seviye, "mesaj": mesaj})

    _okuma_kontrolu(r, ekle)
    net = r.get("net_satis")
    if net is None:
        return s
    _ciro_dengesi(r, net, ekle)
    _kdv_kontrolu(r, net, ekle)
    _banka_ve_adetler(r, ekle)
    return s


# --- doldurma ve kümülatif zincir ----------------------------------------------------------
def _ardisik(a: dict | None, b: dict | None) -> bool:
    """b, a'dan hemen sonraki Z mi."""
    return a is not None and b is not None and None not in (_znum(a), _znum(b)) \
        and _znum(b) == _znum(a) + 1


def _dus(r: dict, alan: str) -> None:
    """Doldurulan alanın 'okunamadı' / okuma notu / iki-okuma uyarılarını kaldırır."""
    r["_sorunlar"] = [x for x in r["_sorunlar"] if not (x["alan"] == alan and (
        x["mesaj"].endswith("okunamadi") or x["mesaj"].startswith("okuma notu")
        or x["mesaj"].startswith("iki okuma farkli")))]


def doldur(r: dict) -> None:
    """Aynı fişin diğer rakamlarından KESİN hesaplanabilen tek eksik alanları doldurur."""
    net = r.get("net_satis")
    kd = r.get("kdv_dokumu") or []
    if len(kd) == 1 and kd[0].get("toplam") is None and net is not None:
        kd[0]["toplam"] = net  # tek oranlı fişte vergi dökümü toplamı = net satış
    if r.get("mali_toplam") is None and net is not None:
        r["mali_toplam"] = net  # MALİ VERİ TOPLAM her zaman net satışa eşit
        _dus(r, "mali_toplam")

    od = r.get("odemeler") or []
    eksik = [o for o in od if o.get("tutar") is None]
    if net is not None and len(eksik) == 1:
        e = eksik[0]
        e["tutar"] = round(net - sum(_t(o["tutar"]) for o in od if o is not e), 2)
        banka = [b.get("tutar") for b in r.get("banka_dokumu") or []]
        kanit = odeme_turu(e.get("tur")) == "kart" and banka and None not in banka \
            and abs(sum(_t(x) for x in banka) - e["tutar"]) <= TOL
        _dus(r, "odemeler")
        _ekle(r, "odemeler", "duzeltme" if kanit else "uyari",
              f"{e.get('tur')} tutari okunamadi; net satis − diger odemeler = {tl(e['tutar'])}"
              + (" (banka dokumu de ayni)" if kanit else " — fisten teyit edin"))

    eksik = [k for k in kd if k.get("kdv") is None]
    if len(kd) > 1 and len(eksik) == 1 and r.get("mali_topkdv") is not None:
        e = eksik[0]
        e["kdv"] = round(_t(r["mali_topkdv"]) - sum(_t(k["kdv"]) for k in kd if k is not e), 2)
        o = _t(e.get("oran"))
        uyum = e.get("toplam") is not None and abs(_t(e["toplam"]) * o / (100 + o) - e["kdv"]) <= KDV_TOL
        _dus(r, "kdv_dokumu")
        _ekle(r, "kdv_dokumu", "duzeltme" if uyum else "uyari",
              f"%{o:g} KDV okunamadi; TOPKDV − diger oranlar = {tl(e['kdv'])}"
              + ("" if uyum else " — oranla uyusmuyor, fisten teyit edin"))

    if r.get("mali_topkdv") is None and kd and None not in [k.get("kdv") for k in kd]:
        r["mali_topkdv"] = round(sum(_t(k["kdv"]) for k in kd), 2)
        _dus(r, "mali_topkdv")
        _ekle(r, "mali_topkdv", "duzeltme",
              f"TOPKDV okunamadi; vergi dokumu satirlarindan {tl(r['mali_topkdv'])}")


def _kumeler(degerler: list[tuple[int, float]], tol: float) -> list[list[tuple[int, float]]]:
    """(indeks, sabit) çiftlerini sabite göre tol içinde kümeler, büyükten küçüğe."""
    kumeler: list[list] = []
    for i, c in sorted(degerler, key=lambda x: x[1]):
        if kumeler and c - kumeler[-1][-1][1] <= tol:
            kumeler[-1].append((i, c))
        else:
            kumeler.append([(i, c)])
    return sorted(kumeler, key=len, reverse=True)


def _belirsiz(r: dict, alan: str) -> bool:
    """Alanda uyarı var mı (iki okuma farklı, okuma notu, tahmin): zinciri bağlamakta kullanılmaz."""
    return any(x["alan"] == alan and x["seviye"] == "uyari" for x in r["_sorunlar"])


def _parcalara_bol(seg: list[dict], gun: str) -> list[list[dict]]:
    """Günlüğü bilinmeyen ya da şüpheli Z, önceki Z'den bağlanamaz: orada yeni parça başlar."""
    parcalar, cari = [], []
    for r in seg:
        if cari and (r.get(gun) is None or _belirsiz(r, gun)):
            parcalar.append(cari)
            cari = []
        cari.append(r)
    parcalar.append(cari)
    return parcalar


def uzlastir(seg: list[dict], kum: str, gun: str, ad: str, tol: float) -> None:
    """Ardışık Z'lerde kümülatif zinciri bir bütün olarak uzlaştırır (kontrol 6, düzeltme).

    Doğru okunmuş her kümülatif için  K_j − (baştan j'ye günlük toplam) = aynı sabit.
    Çoğunluk sabiti doğrudur; ondan sapan / okunamayan kümülatif tek tek düzeltilir (mavi).
    Sapanlar zincirin bir ucuna kadar aynı sabitte toplanıyorsa sorun kümülatiflerde değil,
    o sınırdaki GÜNLÜK rakamdadır: o zaman hiçbir şey düzeltilmez, günlük rakam işaretlenir.
    """
    for p in _parcalara_bol(seg, gun):
        S = [0.0]
        for r in p[1:]:
            S.append(round(S[-1] + _t(r[gun]), 2))
        bilinen = [(j, round(_t(r[kum]) - S[j], 2)) for j, r in enumerate(p) if r.get(kum) is not None]
        kumeler = _kumeler(bilinen, tol)
        if not kumeler or len(kumeler[0]) < 2 or (len(kumeler) > 1 and len(kumeler[1]) == len(kumeler[0])):
            continue  # çoğunluk yok: ikili kontroller işaretler
        ana = kumeler[0]
        ana_j = {j for j, _ in ana}
        bilinen_j = [j for j, _ in bilinen]

        # günlük rakam hatası deseni: aynı sabitte >=2 sapan, zincirin bir ucunu kaplıyor
        gunluk_hata = set()
        for k in kumeler[1:]:
            kj = sorted(j for j, _ in k)
            uc_son = kj == [j for j in bilinen_j if j >= kj[0]]
            uc_bas = kj == [j for j in bilinen_j if j <= kj[-1]]
            if len(kj) >= 2 and (uc_son or uc_bas):
                gunluk_hata |= set(kj)
                fark = round(k[0][1] - ana[0][1], 2)
                sinir = p[kj[0]] if uc_son else p[min(ana_j)]
                if uc_bas:  # sapan blok baştaysa sınır, ana kümenin ilk Z'si
                    fark = -fark
                _ekle(sinir, gun, "hata",
                      f"kumulatifler bu gunun {ad} rakamini {tl(_t(sinir[gun]) + fark)} gosteriyor "
                      f"(okunan {tl(sinir[gun])}) — gunluk rakamda okuma hatasi olabilir, fise bakin")

        for j, r in enumerate(p):
            if j in ana_j or j in gunluk_hata:
                if j in ana_j:  # zincir doğruladı: iki-okuma uyarısı bilgi taşımıyor
                    r["_sorunlar"] = [x for x in r["_sorunlar"]
                                      if not (x["alan"] == kum and x["mesaj"].startswith("iki okuma farkli"))]
                continue
            capa = min(ana, key=lambda x: abs(x[0] - j))[0]  # en yakın doğru kümülatif (kuruş kayması için)
            beklenen = round(_t(p[capa][kum]) + S[j] - S[capa], 2)
            eski = r.get(kum)
            r[kum] = beklenen
            r["_sorunlar"] = [x for x in r["_sorunlar"] if x["alan"] != kum]
            _ekle(r, kum, "duzeltme",
                  f"kumulatif {ad} fiste {tl(eski) + ' okundu' if eski is not None else 'okunamadi'}; "
                  f"zincirdeki {len(ana)} fisle tutarli deger {tl(beklenen)} — duzeltildi")


def _segmentler(raporlar: list[dict]) -> list[list[dict]]:
    """Z no'su ardışık raporları kesintisiz parçalara ayırır."""
    segmentler, cari = [], []
    for r in raporlar:
        if cari and not _ardisik(cari[-1], r):
            segmentler.append(cari)
            cari = []
        cari.append(r)
    if cari:
        segmentler.append(cari)
    return segmentler


def _gunluk_kdv_doldur(seg: list[dict]) -> None:
    """Günlük TOPKDV hâlâ boşsa: iki kümülatifin farkı (kesin) ya da oran formülü (tahmini)."""
    for i, r in enumerate(seg):
        supheli = r.get("mali_topkdv") is not None and _belirsiz(r, "mali_topkdv")
        if (r.get("mali_topkdv") is not None and not supheli) or r.get("net_satis") is None:
            continue
        onceki = seg[i - 1] if i > 0 else None
        if onceki and None not in (onceki.get("kum_topkdv"), r.get("kum_topkdv")):
            fark = round(_t(r["kum_topkdv"]) - _t(onceki["kum_topkdv"]), 2)
            okunan = r.get("mali_topkdv")
            r["mali_topkdv"] = fark
            if len(r.get("kdv_dokumu") or []) == 1:
                r["kdv_dokumu"][0]["kdv"] = fark
            r["_sorunlar"] = [x for x in r["_sorunlar"] if x["alan"] not in ("mali_topkdv", "kdv_dokumu")]
            _ekle(r, "mali_topkdv", "duzeltme",
                  ("TOPKDV okunamadi" if okunan is None else f"TOPKDV {tl(okunan)} okundu (supheli)")
                  + f"; iki kumulatifin farkindan {tl(fark)} (±0,01)")
            continue
        if supheli:
            continue
        if len(r.get("kdv_dokumu") or []) <= 1:
            kd = (r.get("kdv_dokumu") or [{}])[0]
            o = _t(kd.get("oran")) or _t(((r.get("departmanlar") or [{}])[0]).get("kdv_orani"))
            if o:
                r["mali_topkdv"] = round(_t(r["net_satis"]) * o / (100 + o), 2)
                _dus(r, "mali_topkdv")
                _ekle(r, "mali_topkdv", "uyari",
                      f"TOPKDV okunamadi; %{o:g} ile TAHMINI {tl(r['mali_topkdv'])} — fisten teyit edin")
        if r.get("mali_topkdv") is not None:
            for k in r.get("kdv_dokumu") or []:
                if k.get("kdv") is None and len(r["kdv_dokumu"]) == 1:
                    k["kdv"] = r["mali_topkdv"]


def onar(raporlar: list[dict]) -> None:
    """Z no sırasındaki raporlar: fiş içi doldurma, kümülatif uzlaştırma, günlük KDV doldurma."""
    for r in raporlar:
        doldur(r)
    for seg in _segmentler(raporlar):
        uzlastir(seg, "kum_toplam", "net_satis", "toplam", 0.005)
        uzlastir(seg, "kum_topkdv", "mali_topkdv", "KDV", 0.025)
        _gunluk_kdv_doldur(seg)
        # yeni günlük KDV'lerle kümülatif KDV'yi bir daha uzlaştır (boş kalan uçlar dolar)
        uzlastir(seg, "kum_topkdv", "mali_topkdv", "KDV", 0.025)


def _notlari_sadelestir(r: dict) -> None:
    """Düzeltilen alanın 'okuma notu' uyarıları artık bilgi taşımıyor."""
    duzelen = {s["alan"] for s in r["_sorunlar"] if s["seviye"] == "duzeltme"}
    if "mali_topkdv" in duzelen:
        duzelen |= {"kdv_dokumu", "kum_topkdv"}
    r["_sorunlar"] = [s for s in r["_sorunlar"] if not (
        s["alan"] in duzelen and s["seviye"] == "uyari" and s["mesaj"].startswith("okuma notu"))]


def _ardisiklik_kontrolu(onceki: dict, r: dict) -> None:
    """Kontrol 6-7 (ikili): Z no atlaması/tekrarı, tarih sırası, kümülatif fark = günlük."""
    s = r["_sorunlar"]
    z0, z1 = _znum(onceki), _znum(r)
    if z0 is None or z1 is None:
        return
    if z1 == z0:
        s.append({"alan": "z_no", "seviye": "hata", "mesaj": f"Z {z1} iki kez var"})
        return
    if z1 != z0 + 1:
        eksik = f"Z {z0 + 1}" if z1 == z0 + 2 else f"Z {z0 + 1}–{z1 - 1}"
        ek = ""
        if None not in (onceki.get("kum_toplam"), r.get("kum_toplam"), r.get("net_satis")):
            ek = (f"; kumulatife gore eksik {'Z' + chr(39) + 'nin' if z1 == z0 + 2 else 'Z' + chr(39) + 'lerin'} cirosu "
                  f"{tl(_t(r['kum_toplam']) - _t(r['net_satis']) - _t(onceki['kum_toplam']))}")
        s.append({"alan": "z_no", "seviye": "hata",
                  "mesaj": f"{eksik} fotografta yok{ek} — o fisi de ekleyin"})
        return
    d0, d1 = _tarih(onceki.get("tarih")), _tarih(r.get("tarih"))
    if d0 and d1:
        if d1 <= d0:
            s.append({"alan": "tarih", "seviye": "hata",
                      "mesaj": f"tarih onceki Z'den ({onceki['tarih']}) ileri degil"})
        elif d1 - d0 > timedelta(days=1):
            s.append({"alan": "tarih", "seviye": "uyari",
                      "mesaj": f"{onceki['tarih']} sonrasi {(d1 - d0).days - 1} gun Z yok (kapali gun?)"})
    for kum, gun, ad in (("kum_toplam", "net_satis", "toplam"), ("kum_topkdv", "mali_topkdv", "KDV")):
        if None in (onceki.get(kum), r.get(kum), r.get(gun)):
            continue
        fark = _t(r[kum]) - _t(onceki[kum])
        if abs(fark - _t(r[gun])) > KUM_TOL:
            s.append({"alan": kum, "seviye": "hata",
                      "mesaj": f"kumulatif {ad} farki {tl(fark)} ≠ gunluk {ad} {tl(r[gun])}; "
                               f"gunluk rakam dogruysa kumulatif {tl(_t(onceki[kum]) + _t(r[gun]))} "
                               f"olmali (bu fiste, onceki Z'de ya da gunluk rakamda okuma hatasi)"})


def zincir(raporlar: list[dict]) -> None:
    """Aynı firmanın raporlarını Z no sırasına koyar, onarır ve zincir kontrollerini (6-7) yapar."""
    raporlar.sort(key=lambda r: (_znum(r) is None, _znum(r) or 0))
    onar(raporlar)
    for i, r in enumerate(raporlar):
        _notlari_sadelestir(r)
        # zincirin iki yandan doğruladığı alanlarda "iki okuma farklı" uyarısı bilgi taşımıyor
        komsu = i > 0 and i + 1 < len(raporlar) and _ardisik(raporlar[i - 1], r) \
            and _ardisik(r, raporlar[i + 1])
        zincir_hatasi = any(s["seviye"] == "hata" and s["alan"] in ZINCIRLI for s in r["_sorunlar"]) \
            or (i + 1 < len(raporlar) and any(s["seviye"] == "hata" and s["alan"] in ZINCIRLI
                                              for s in raporlar[i + 1]["_sorunlar"]))
        if komsu and not zincir_hatasi:
            r["_sorunlar"] = [s for s in r["_sorunlar"] if not s.get("_zincirli")]
    fis_adetleri = [r["toplam_fis_adedi"] for r in raporlar if r.get("toplam_fis_adedi")]
    orta = median(fis_adetleri) if fis_adetleri else None
    for onceki, r in zip([None] + raporlar[:-1], raporlar):
        if orta and r.get("toplam_fis_adedi") is not None and r["toplam_fis_adedi"] < 0.3 * orta:
            r["_sorunlar"].append({"alan": "toplam_fis_adedi", "seviye": "uyari",
                                   "mesaj": f"toplam fis adedi {r['toplam_fis_adedi']}, ay ortancasi {orta:g}"})
        if onceki is not None:
            _ardisiklik_kontrolu(onceki, r)


# --- firma gruplama ------------------------------------------------------------------------
def _norm(s) -> str:
    """Büyük harf, Türkçe karakter sadeleştirme, yalnız harf-rakam."""
    return re.sub(r"[^A-Z0-9]", "", (s or "").upper().translate(str.maketrans("İIĞÜŞÖÇ", "IIGUSOC")))


def _cihaz(s) -> str:
    """'AF AV0000000001' -> 'AV0000000001'."""
    n = _norm(s)
    return n[2:] if n.startswith("AF") and len(n) > 10 else n


def _yakin(a: str, b: str) -> bool:
    """Aynı uzunlukta ve en çok 1 hane farklı (tek hane okuma hatası)."""
    return bool(a) and len(a) == len(b) and sum(x != y for x, y in zip(a, b)) <= 1


def grupla(raporlar: list[dict]) -> dict[str, list[dict]]:
    """Aynı mükellefin raporlarını gruplar.

    Cihaz sicil no ya da VKN (1 hane okuma farkı tolere) ortaksa, ya da unvan aynı ve VKN'lerden
    biri boşsa birleşir. Başlık alanları (unvan, VKN...) çoğunluk oyuyla seçilir.
    """
    ana = list(range(len(raporlar)))

    def bul(i):
        while ana[i] != i:
            ana[i] = ana[ana[i]]
            i = ana[i]
        return i

    for i, a in enumerate(raporlar):
        for j in range(i):
            b = raporlar[j]
            ca, cb = _cihaz(a.get("cihaz_sicil_no")), _cihaz(b.get("cihaz_sicil_no"))
            va, vb = _norm(a.get("vkn")), _norm(b.get("vkn"))
            ua, ub = _norm(a.get("firma_unvan")), _norm(b.get("firma_unvan"))
            if (ca and ca == cb) or _yakin(va, vb) or (ua and ua == ub and (not va or not vb)):
                ana[bul(i)] = bul(j)
    gruplar: dict[int, list[dict]] = {}
    for i, r in enumerate(raporlar):
        gruplar.setdefault(bul(i), []).append(r)

    firmalar: dict[str, list[dict]] = {}
    for liste in gruplar.values():
        for r in liste:
            if r.get("cihaz_sicil_no"):
                r["cihaz_sicil_no"] = _cihaz(r["cihaz_sicil_no"])
        for alan in ("firma_unvan", "firma_sahibi", "vkn", "vergi_dairesi", "cihaz_sicil_no"):
            degerler = [r.get(alan) for r in liste if r.get(alan)]
            if not degerler:
                continue
            cogunluk = mode(degerler)
            farkli = sorted({r[alan] for r in liste if r.get(alan) and r[alan] != cogunluk})
            if farkli and alan in ("vkn", "cihaz_sicil_no"):
                # tek bir not yeter; her fişi sarıya boyamak gürültü olur
                liste[0]["_sorunlar"].append({"alan": alan, "seviye": "uyari", "mesaj":
                    f"{alan} {sum(r.get(alan) in farkli for r in liste)} fiste {', '.join(farkli)} "
                    f"okundu; cogunluk {cogunluk} kullanildi"})
            for r in liste:
                if r.get(alan):
                    r[alan] = cogunluk
        firmalar[liste[0].get("vkn") or liste[0].get("firma_unvan") or f"firma{len(firmalar)}"] = liste
    return firmalar


# --- çoklu okuma ve ana giriş ----------------------------------------------------------------
def birlestir(okumalar: list[list[dict]]) -> list[dict]:
    """Aynı fotoğrafın bağımsız okumalarını uzlaştırır (ilk okuma esas).

    Farklı çıkan alan uyarı olur (`_okuma_farklari`); ilk okumanın göremediğini ikinci gördüyse
    değer alınır. Yalnız bir okumada görülen fiş de uyarı olur.
    """
    if len(okumalar) == 1:
        return okumalar[0]
    esas, digerleri = okumalar[0], okumalar[1:]
    sozluk = [{_znum(r): r for r in o} for o in digerleri]
    sonuc = []
    for r in esas:
        r["_okuma_farklari"] = []
        for d in sozluk:
            e = d.pop(_znum(r), None)
            if e is None:
                r["_okuma_farklari"].append({"alan": "z_no", "seviye": "uyari",
                                             "mesaj": "bu Z yalniz bir okumada bulundu"})
                continue
            cift = [(a, r.get(a), e.get(a)) for a in KARSILASTIR]
            for tur in ("nakit", "kart", "diger"):
                cift.append((f"odeme_{tur}", odeme(r, tur), odeme(e, tur)))
            for alan, x, y in cift:
                if x == y or (isinstance(x, (int, float)) and isinstance(y, (int, float))
                              and abs(x - y) <= TOL):
                    continue
                if x is None:  # ilk okuma göremedi, ikinci gördü
                    if alan in r:
                        r[alan] = y
                    x = "okunamadi"
                r["_okuma_farklari"].append({
                    "alan": "odemeler" if alan.startswith("odeme_") else alan,
                    "seviye": "uyari", "_zincirli": alan in ZINCIRLI,
                    "mesaj": f"iki okuma farkli: {alan} = {x} / {'okunamadi' if y is None else y}"})
        sonuc.append(r)
    for d in sozluk:  # yalnız sonraki okumalarda bulunanlar
        for e in d.values():
            e["_okuma_farklari"] = [{"alan": "z_no", "seviye": "uyari",
                                     "mesaj": "bu Z yalniz bir okumada bulundu"}]
            sonuc.append(e)
    return sonuc


def _tekrarlari_ayikla(liste: list[dict]) -> list[dict]:
    """Aynı Z iki fotoğrafta varsa daha az hatalı olanı tutar; rakamlar farklıysa hata yazar."""
    tekil: dict = {}
    for r in liste:
        z = _znum(r)
        k = z if z is not None else id(r)
        if k not in tekil:
            tekil[k] = r
            continue
        eski = tekil[k]
        farkli = [a for a in ("net_satis", "kum_toplam", "mali_topkdv", "tarih") if eski.get(a) != r.get(a)]
        iyi, kotu = sorted((eski, r), key=lambda x: sum(1 for s in x["_sorunlar"] if s["seviye"] == "hata"))
        if farkli:
            iyi["_sorunlar"].append({"alan": farkli[0], "seviye": "hata",
                                     "mesaj": f"ayni Z iki fotografta farkli okundu ({', '.join(farkli)}): "
                                              f"{iyi.get('_kaynak')} / {kotu.get('_kaynak')}"})
        tekil[k] = iyi
    return list(tekil.values())


def kontrol_et(raporlar: list[dict]) -> dict[str, list[dict]]:
    """Ana giriş: tek fiş kontrolleri, firmaya (VKN) göre gruplama, tekrar ayıklama, zincir.

    Döner: {firma anahtarı (VKN): Z no sırasında raporlar}; her raporda `_sorunlar` vardır.
    """
    for r in raporlar:
        r["_sorunlar"] = tek_rapor(r) + r.pop("_okuma_farklari", [])
    firmalar = grupla(raporlar)
    for anahtar, liste in firmalar.items():
        firmalar[anahtar] = _tekrarlari_ayikla(liste)
        zincir(firmalar[anahtar])
    return firmalar


def ozet(firmalar: dict[str, list[dict]]) -> list[dict]:
    """Firma başına kısa özet: Z sayısı, dönem, hata/uyarı/düzeltme sayıları."""
    sonuc = []
    for rs in firmalar.values():
        say = {sv: sum(1 for r in rs for s in r["_sorunlar"] if s["seviye"] == sv)
               for sv in ("hata", "uyari", "duzeltme")}
        tarihler = sorted((r["tarih"] for r in rs if _tarih(r.get("tarih"))), key=_tarih)
        sonuc.append({"firma": rs[0].get("firma_unvan"), "vkn": rs[0].get("vkn"), "z": len(rs),
                      "ilk": tarihler[0] if tarihler else None,
                      "son": tarihler[-1] if tarihler else None, **say})
    return sonuc
