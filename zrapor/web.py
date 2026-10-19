"""Telefon yükleme sayfası: fotoğraf yükle -> kontrollü Excel ve Luca dosyalarını indir.

    ZRAPOR_SIFRE=...   (zorunlu, en az 8 karakter; ortam değişkeni ya da .env)
    ZRAPOR_ARKA_UC=... (isteğe bağlı: api | cli | gemini | karma; varsayılan api)
    z-rapor web [--port 8765] [--veri ./calisma] [--arka-uc karma]

İşler sırayla işlenir (API/abonelik limitleri için tek işçi); her iş `<veri>/isler/<id>/` altında:
fotolar/, cikti/, durum.json. Sayfa kapansa da iş sürer; sunucu yeniden başlarsa yarım işler
tekrar sıraya girer. Firma hesap kodları `<veri>/hesaplar.json` dosyasında tutulur.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import queue
import re
import secrets
import shutil
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

try:
    import uvicorn
    from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
    from fastapi.responses import FileResponse, HTMLResponse
except ImportError as e:  # pragma: no cover
    raise SystemExit("Web sayfası için: pip install 'z-rapor-okuyucu[web]'") from e
from PIL import Image

from . import luca, masraf
from .ayarlar import VARSAYILAN_ARKA_UC, VARSAYILAN_CALISMA, VARSAYILAN_MODEL, VARSAYILAN_PORT
from .isleme import isle
from .okuma import ARKA_UCLAR

SAYFA = Path(__file__).parent / "static" / "web.html"
CEREZ = "zr"
MAKS_FOTO, MAKS_MB = 40, 25
DK_PER_CAGRI = 8  # kaba süre tahmini: bir okuma çağrısı ~8 dk'ya kadar sürebilir; 3 çağrı paralel
HESAP_ALANLARI = ("kasa", "pos", "diger_odeme", "aciklama", "fis_no")
ORANLAR = ("1", "10", "20")


@dataclass
class Ayar:
    """Sunucunun çalışma ayarları (`calistir` doldurur)."""
    veri: Path = VARSAYILAN_CALISMA
    sifre: str = ""
    model: str = VARSAYILAN_MODEL
    arka_uc: str = VARSAYILAN_ARKA_UC
    giris_denemeleri: dict[str, list[float]] = field(default_factory=dict)

    @property
    def isler_klasoru(self) -> Path:
        return self.veri / "isler"

    @property
    def hesap_dosyasi(self) -> Path:
        return self.veri / "hesaplar.json"


AYAR = Ayar()
app = FastAPI(docs_url=None, redoc_url=None)
kuyruk: queue.Queue[str] = queue.Queue()
kilit = threading.Lock()
isler: dict[str, dict] = {}


# --- yetki -------------------------------------------------------------------------------------
def _jeton() -> str:
    """Paroladan türetilen oturum çerezi değeri (parola değişince eski oturumlar düşer)."""
    return hmac.new(AYAR.sifre.encode(), b"zrapor-oturum", hashlib.sha256).hexdigest()


def yetki(req: Request) -> None:
    """Geçerli oturum çerezi yoksa 401."""
    if not hmac.compare_digest(req.cookies.get(CEREZ, ""), _jeton()):
        raise HTTPException(401, "giris gerekli")


@app.post("/api/giris")
def giris(req: Request, resp: Response, sifre: str = Form(...)):
    """Parolayla giriş; 5 dakikada 8 yanlış denemeden sonra geçici kilit."""
    # Bir tünel/ters vekil arkasında bütün istekler aynı adresten gelir: sınır o zaman ortak olur.
    ip = req.client.host if req.client else "?"
    simdi = time.time()
    son = [t for t in AYAR.giris_denemeleri.get(ip, []) if simdi - t < 300]
    if len(son) >= 8:
        raise HTTPException(429, "cok fazla deneme, 5 dk bekleyin")
    if not hmac.compare_digest(sifre.encode(), AYAR.sifre.encode()):
        AYAR.giris_denemeleri[ip] = son + [simdi]
        raise HTTPException(403, "sifre yanlis")
    resp.set_cookie(CEREZ, _jeton(), max_age=60 * 60 * 24 * 90, httponly=True, samesite="strict",
                    secure=req.url.scheme == "https" or req.headers.get("x-forwarded-proto") == "https")
    return {"ok": True}


# --- iş kaydı ----------------------------------------------------------------------------------
def _kaydet(is_: dict) -> None:
    (AYAR.isler_klasoru / is_["id"] / "durum.json").write_text(
        json.dumps(is_, ensure_ascii=False, indent=1), encoding="utf-8")


def _guncelle(is_id: str, **alanlar) -> None:
    with kilit:
        is_ = isler[is_id]
        is_.update(alanlar)
        _kaydet(is_)


def _yukle_eski_isler() -> None:
    """Diskteki işleri belleğe alır; yarım kalanları yeniden sıraya koyar."""
    for d in sorted(AYAR.isler_klasoru.glob("*/durum.json")):
        is_ = json.loads(d.read_text(encoding="utf-8"))
        if is_["durum"] in ("sirada", "isleniyor"):
            is_["durum"] = "sirada"
            is_["mesajlar"].append("sunucu yeniden basladi, is tekrar siraya alindi")
            kuyruk.put(is_["id"])
        isler[is_["id"]] = is_


def _dosya_adi(s: str) -> str:
    return re.sub(r"[^\w\-. ]", "", s, flags=re.UNICODE).strip().replace(" ", "_")[:60] or "z_rapor"


def _masraf_isi(is_id: str, is_: dict, fotolar: list[Path], gecici: Path, bildir) -> None:
    mo = masraf.isle(fotolar, gecici, model=AYAR.model, arka_uc=AYAR.arka_uc, okuma=is_["okuma"],
                     hazirlayan=is_.get("hazirlayan") or "", bildir=bildir)
    ad = _dosya_adi(is_.get("etiket") or "Masraf_Listesi")
    if mo.get("ilk"):
        ad += f"_{mo['ilk']}_{mo['son']}".replace(".", "-")
    hedef = gecici.with_name(f"{ad}.xlsx")
    gecici.replace(hedef)
    _guncelle(is_id, durum="bitti", masraf_ozet=mo, excel=hedef.name,
              bitti=datetime.now().isoformat(timespec="seconds"))


def _z_isi(is_id: str, is_: dict, fotolar: list[Path], gecici: Path, bildir) -> None:
    ozet = isle(fotolar, gecici, model=AYAR.model, arka_uc=AYAR.arka_uc, okuma=is_["okuma"], bildir=bildir)
    o = ozet[0] if len(ozet) == 1 else None
    ad = _dosya_adi(is_.get("etiket") or (o and o["firma"]) or "Z_raporlari")
    if o and o["ilk"]:
        ad += f"_{o['ilk']}_{o['son']}".replace(".", "-")
    hedef = gecici.with_name(f"{ad}.xlsx")
    gecici.replace(hedef)
    _guncelle(is_id, durum="bitti", ozet=ozet, excel=hedef.name,
              bitti=datetime.now().isoformat(timespec="seconds"))
    try:
        _luca_uret(is_id)
    except Exception as e:  # Luca dosyası üretilemese de Excel hazır
        bildir(f"Luca dosyasi uretilemedi: {e}")


def isci() -> None:
    """Kuyruktaki işleri sırayla işler (tek işçi)."""
    while True:
        is_id = kuyruk.get()
        is_ = isler[is_id]
        _guncelle(is_id, durum="isleniyor", basladi=datetime.now().isoformat(timespec="seconds"))
        klasor = AYAR.isler_klasoru / is_id
        fotolar = sorted((klasor / "fotolar").iterdir())
        gecici = klasor / "cikti" / "z_rapor.xlsx"

        def bildir(m, _id=is_id):
            with kilit:
                isler[_id]["mesajlar"].append(f"{datetime.now():%H:%M} {m}")
                _kaydet(isler[_id])

        try:
            if is_.get("tur") == "masraf":
                _masraf_isi(is_id, is_, fotolar, gecici, bildir)
            else:
                _z_isi(is_id, is_, fotolar, gecici, bildir)
        except Exception as e:
            bildir(f"HATA: {e}")
            _guncelle(is_id, durum="hata", hata=str(e)[:500])
        finally:
            kuyruk.task_done()


def _raporlar(is_id: str) -> dict:
    veri = AYAR.isler_klasoru / is_id / "cikti" / "raporlar.json"
    if not veri.exists():
        raise HTTPException(409, "Bu isin kontrol verisi yok")
    return json.loads(veri.read_text(encoding="utf-8"))


def _luca_uret(is_id: str) -> Path:
    """Kontrol edilmiş veriden güncel hesap kodlarıyla Luca bilanço fiş aktarım dosyası üretir."""
    is_ = isler[is_id]
    hedef = AYAR.isler_klasoru / is_id / "cikti" / (Path(is_["excel"]).stem + "_LUCA_bilanco.xlsx")
    o = luca.bilanco_yaz(_raporlar(is_id), hedef, luca.hesaplar_oku(AYAR.hesap_dosyasi))
    _guncelle(is_id, luca=o | {"dosya": hedef.name})
    return hedef


# --- API -----------------------------------------------------------------------------------------
def _kamu(is_: dict) -> dict:
    """İşin sayfaya gönderilen alanları + sıradaki yeri."""
    sira = None
    if is_["durum"] == "sirada":
        bekleyen = sorted((i for i in isler.values() if i["durum"] == "sirada"), key=lambda i: i["olustu"])
        sira = [i["id"] for i in bekleyen].index(is_["id"]) + 1
    return {k: is_.get(k) for k in ("id", "etiket", "durum", "olustu", "basladi", "bitti", "foto",
                                    "okuma", "mesajlar", "ozet", "excel", "hata", "tahmini_dk",
                                    "luca", "luca_isletme", "tur", "masraf_ozet", "hazirlayan")} | {"sira": sira}


def _bitmis_is(is_id: str) -> dict:
    is_ = isler.get(is_id)
    if not is_ or is_["durum"] != "bitti":
        raise HTTPException(404, "Excel hazir degil")
    return is_


@app.post("/api/is")
async def is_olustur(req: Request, fotolar: list[UploadFile] = File(...), etiket: str = Form(""),
                     okuma: int = Form(2), tur: str = Form("z"), hazirlayan: str = Form("")):
    """Yüklenen fotoğraflarla yeni iş açar ve sıraya koyar."""
    yetki(req)
    if not 1 <= len(fotolar) <= MAKS_FOTO:
        raise HTTPException(400, f"1-{MAKS_FOTO} fotograf yukleyin")
    is_id = datetime.now().strftime("%Y%m%d-%H%M%S-") + secrets.token_hex(3)
    klasor = AYAR.isler_klasoru / is_id / "fotolar"
    klasor.mkdir(parents=True)
    try:
        for i, f in enumerate(fotolar, 1):
            veri = await f.read()
            if len(veri) > MAKS_MB * 1024 * 1024:
                raise HTTPException(400, f"{f.filename}: {MAKS_MB} MB'tan buyuk")
            yol = klasor / f"{i:02d}.img"
            yol.write_bytes(veri)
            try:
                with Image.open(yol) as im:
                    im.verify()
                    bicim = (im.format or "jpeg").lower()
            except Exception:
                raise HTTPException(400, f"{f.filename}: resim olarak acilamadi (HEIC ise JPEG'e cevirin)")
            yol.rename(yol.with_suffix(".jpg" if bicim == "jpeg" else f".{bicim}"))
    except HTTPException:
        shutil.rmtree(AYAR.isler_klasoru / is_id, ignore_errors=True)  # yarım iş bırakma
        raise
    okuma = 1 if okuma == 1 else 2
    is_ = {"id": is_id, "etiket": etiket.strip()[:60], "durum": "sirada", "foto": len(fotolar),
           "tur": "masraf" if tur == "masraf" else "z", "hazirlayan": hazirlayan.strip()[:40],
           "okuma": okuma, "olustu": datetime.now().isoformat(timespec="seconds"), "mesajlar": [],
           "tahmini_dk": -(-len(fotolar) * okuma // 3) * DK_PER_CAGRI}
    (AYAR.isler_klasoru / is_id / "cikti").mkdir()
    with kilit:
        isler[is_id] = is_
        _kaydet(is_)
    kuyruk.put(is_id)
    return _kamu(is_)


@app.get("/api/is/{is_id}")
def is_durumu(is_id: str, req: Request):
    yetki(req)
    if is_id not in isler:
        raise HTTPException(404, "is bulunamadi")
    return _kamu(isler[is_id])


@app.get("/api/isler")
def is_listesi(req: Request):
    yetki(req)
    return [_kamu(i) for i in sorted(isler.values(), key=lambda i: i["olustu"], reverse=True)[:30]]


@app.get("/api/is/{is_id}/excel")
def excel_indir(is_id: str, req: Request):
    yetki(req)
    is_ = _bitmis_is(is_id)
    return FileResponse(AYAR.isler_klasoru / is_id / "cikti" / is_["excel"], filename=is_["excel"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/is/{is_id}/luca")
def luca_indir(is_id: str, req: Request):
    """Bilanço fiş aktarım dosyası; her indirmede güncel hesap kodlarıyla yeniden üretilir."""
    yetki(req)
    _bitmis_is(is_id)
    hedef = _luca_uret(is_id)
    return FileResponse(hedef, filename=hedef.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/api/is/{is_id}/luca-isletme")
def luca_isletme_indir(is_id: str, req: Request):
    """İşletme defteri CSV'si; her indirmede güncel kodlarla yeniden üretilir."""
    yetki(req)
    is_ = _bitmis_is(is_id)
    hedef = AYAR.isler_klasoru / is_id / "cikti" / (Path(is_["excel"]).stem + "_LUCA_isletme.csv")
    o = luca.isletme_yaz(_raporlar(is_id), hedef, luca.hesaplar_oku(AYAR.hesap_dosyasi))
    _guncelle(is_id, luca_isletme=o | {"dosya": hedef.name})
    return FileResponse(hedef, filename=hedef.name, media_type="text/csv; charset=windows-1254")


@app.get("/api/hesaplar")
def hesaplar_listesi(req: Request):
    """İşlerde görülen firmalar + güncel Luca hesap kodları."""
    yetki(req)
    firmalar = {}
    for i in isler.values():
        for o in i.get("ozet") or []:
            if o.get("vkn") and o.get("firma"):
                firmalar[o["vkn"]] = o["firma"]
    tum = luca.hesaplar_oku(AYAR.hesap_dosyasi)
    return [{"vkn": v, "firma": f, "kodlar": luca.firma_hesaplari(v, tum),
             "ozel": tum.get(luca.vkn_anahtari(v), {})}
            for v, f in sorted(firmalar.items(), key=lambda x: x[1])]


@app.post("/api/hesaplar/{vkn}")
async def hesap_kaydet(vkn: str, req: Request):
    """Bir firmanın hesap kodlarını kaydeder."""
    yetki(req)
    govde = await req.json()
    temiz = {k: str(govde[k]).strip()[:40] for k in HESAP_ALANLARI if str(govde.get(k) or "").strip()}
    for grup in ("satis", "kdv"):
        g = {o: str((govde.get(grup) or {}).get(o) or "").strip()[:40] for o in ORANLAR}
        if any(g.values()):
            temiz[grup] = {o: v for o, v in g.items() if v}
    fk = str((govde.get("isletme") or {}).get("faaliyet_kodu") or "").strip()[:20]
    if fk:
        temiz["isletme"] = {"faaliyet_kodu": fk}
    if temiz.get("fis_no") not in (None, "z_no", "bos"):
        raise HTTPException(400, "fis_no: z_no ya da bos")
    with kilit:
        tum = luca.hesaplar_oku(AYAR.hesap_dosyasi)
        tum[luca.vkn_anahtari(vkn)] = temiz
        luca.hesaplar_yaz(tum, AYAR.hesap_dosyasi)
    return {"ok": True, "kodlar": luca.firma_hesaplari(vkn, tum)}


@app.get("/", response_class=HTMLResponse)
def sayfa():
    return SAYFA.read_text(encoding="utf-8")


def calistir(port: int = VARSAYILAN_PORT, host: str = "127.0.0.1", veri: Path = VARSAYILAN_CALISMA,
             model: str = VARSAYILAN_MODEL, arka_uc: str = VARSAYILAN_ARKA_UC) -> None:
    """Sunucuyu başlatır. Parola ZRAPOR_SIFRE ortam değişkeninden (ya da .env'den) okunur."""
    AYAR.sifre = os.environ.get("ZRAPOR_SIFRE", "")
    if len(AYAR.sifre) < 8:
        raise SystemExit("ZRAPOR_SIFRE ortam degiskeni (ya da .env) en az 8 karakter olmali")
    if arka_uc not in ARKA_UCLAR:
        raise SystemExit(f"bilinmeyen arka uc: {arka_uc} ({' | '.join(ARKA_UCLAR)})")
    AYAR.veri, AYAR.model, AYAR.arka_uc = Path(veri), model, arka_uc
    AYAR.isler_klasoru.mkdir(parents=True, exist_ok=True)
    _yukle_eski_isler()
    threading.Thread(target=isci, daemon=True).start()
    uvicorn.run(app, host=host, port=port)
