"""Fotoğraf -> yapılandırılmış JSON (görsel dil modeliyle okuma).

Fotoğraf önce büyütülmüş, üst üste binen parçalara bölünür (sıkıştırılmış fotoğraflarda rakamlar
küçülüyor), sonra tam fotoğraf + parçalar tek seferde modele verilir. Model her fişi bir kez,
şemadaki alanlarla döndürür.

Arka uçlar:
    api    -> Anthropic Messages API (ANTHROPIC_API_KEY; ortam değişkeni ya da .env)   [varsayılan]
    cli    -> `claude -p` (Claude Code oturumu; API anahtarı gerekmez)
    gemini -> Google Gemini API (GEMINI_API_KEY); tam fotoğraf tek görsel olarak gider
    karma  -> çapraz okuma: tek numaralı okumalar `api`, çift numaralı okumalar `gemini`.
              İki farklı model aynı hatayı nadiren yapar; farklar birleştirmede uyarıya düşer.
              `okuma=1` ise yalnız `api` çalışır.

Sonuçlar `calisma` klasöründe fotoğrafın içerik özetiyle önbelleklenir; aynı fotoğraf aynı model
ve okuma numarasıyla ikinci kez modele gitmez (`yeniden=True` ile zorlanır).
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import shutil
import subprocess
import time
from pathlib import Path

from PIL import Image, ImageOps

from .ayarlar import (API_MAKS_TOKEN, GEMINI_URL, VARSAYILAN_ARKA_UC, VARSAYILAN_MODEL, api_modeli,
                      gemini_modeli)
from .model import SEMA

ARKA_UCLAR = ("api", "cli", "gemini", "karma")
GEMINI_EK_NOT = ("\n\nBu okumada yalniz fotografin tamami var (parca yok); rakamlari tam fotograftan, "
                 "her haneye dikkat ederek oku.")

TALIMAT = """Sen bir mali musavir burosu icin yazar kasa Z RAPORU fislerini okuyorsun. Okudugun
rakamlar beyannameye girecek: TAHMIN ETME. Okuyamadigin degeri null yaz ve alan adini
emin_olunmayan_alanlar'a ekle (kisa: "kum_topkdv: son iki hane kirpik, gorunen 135.094,3_").
Bir degeri BASKA FISLERDEN HESAPLAYARAK, ornuntuden ya da "olmasi gereken"den DOLDURMA, basili
rakami da zincire uysun diye DEGISTIRME: capraz kontrolu ve duzeltmeyi bizim kodumuz yapiyor.
Sadece fiste gordugunu yaz. Termal baskida 0/8, 6/0, 5/9, 1/7 karisir: her haneye dikkat et.

YALNIZ "Z RAPORU" / "Z GUNLUK RAPORU" basligi ve Z SAYAC / Z NO'su olan gun sonu raporlarini
raporlar'a yaz. Musteri satis fisi (urun satirlari olan market fisi), e-arsiv fatura, kredi karti
POS slibi, X raporu gibi belgeler Z raporu DEGILDIR: onlari raporlar'a KOYMA, z_raporu_olmayan'a
turu ve saticisiyla yaz. Fotograf yan ya da ters cekilmis olabilir; once fisin yonunu bul.

Gorseller: ilki fotografin tamami, digerleri ayni fotografin buyutulmus, ust uste binen
parcalari. Fisleri tam fotograftan tani, rakamlari parcalardan oku. Fotograftaki HER Z raporunu
TAM BIR KEZ dondur (parcalarda tekrar gorunse bile). Fisin bir kismi kadrajin disindaysa
okuyabildigini yaz, tam_gorunuyor=false yap.

Fis duzeni (cihaza gore degisebilir): bir etiket satirinin saginda ADET, hemen alt satirinda
TUTAR olur. Ornek:
  Yiyecek  %10   12        -> departman ad=Yiyecek, kdv_orani=10, adet=12
          *8.110,00        -> tutar=8110.00
  NET SATIS *8.110,00
  NAKIT            1       -> odeme tur=NAKIT adet=1, alttaki tutar
  KASA NAKIT  *350,00      -> kasa_nakit
  K.KARTI         11       -> odeme tur=K.KARTI adet=11, alttaki tutar
  Ornekbank  *7.760,00     -> banka_dokumu
  FIS IPTAL        1       -> fis_iptal_adet, alttaki tutar fis_iptal_tutar
"Z SAYAC" yanindaki numara z_no'dur (fisin en altindaki "Z NO" ile ayni). Ustteki "FIS NO" fis_no.
VERGI DOKUMU'ndeki her "%X TOPLAM / TOPKDV" ciftini kdv_dokumu'ne yaz. MALI VERI TOPLAM/TOPKDV ->
mali_toplam/mali_topkdv; KUM. TOPLAM/KUM. TOPKDV -> kum_toplam/kum_topkdv. "AF AV..." satiri
cihaz_sicil_no. Birden fazla departman (Yiyecek, Icecek...) varsa hepsini ayri yaz.

Sayilar: Turk bicimi "*1.326.747,00" -> 1326747.00 (nokta binlik, virgul ondalik, * isaretini at).
Tarih GG.AA.YYYY, z_no/fis_no fisteki gibi sifirli metin ("0130")."""


def dosya_ozeti(p: Path) -> str:
    """Fotoğraf içeriğinin kısa SHA-1 özeti (önbellek anahtarı)."""
    return hashlib.sha1(p.read_bytes()).hexdigest()[:12]


def parcala(foto: Path, hedef: Path) -> list[Path]:
    """Tam görüntü + 3x3 üst üste binen parça üretir. Küçük fotoğraflar 2x büyütülür."""
    hedef.mkdir(parents=True, exist_ok=True)
    im = ImageOps.exif_transpose(Image.open(foto)).convert("RGB")
    W, H = im.size
    tam = im.copy()
    tam.thumbnail((1600, 1600))
    yollar = [hedef / "00_tam.jpg"]
    tam.save(yollar[0], quality=92)
    olcek = 2 if max(W, H) < 2400 else 1
    tw, th = int(W * 0.45), int(H * 0.55)
    for yi, y0 in enumerate((0, (H - th) // 2, H - th)):
        for xi, x0 in enumerate((0, (W - tw) // 2, W - tw)):
            p = im.crop((x0, y0, x0 + tw, y0 + th))
            if olcek > 1:
                p = p.resize((tw * olcek, th * olcek), Image.LANCZOS)
            p.thumbnail((1800, 1800))
            yol = hedef / f"p{yi}{xi}.jpg"
            p.save(yol, quality=92)
            yollar.append(yol)
    return yollar


def tam_gorsel(foto: Path, hedef: Path, maks: int = 3072) -> list[Path]:
    """Yönü düzeltilmiş tek tam görüntü (JPEG). Gemini arka ucu parçasız okur."""
    hedef.mkdir(parents=True, exist_ok=True)
    im = ImageOps.exif_transpose(Image.open(foto)).convert("RGB")
    im.thumbnail((maks, maks))
    yol = hedef / "00_tam.jpg"
    im.save(yol, quality=92)
    return [yol]


def gemini_semasi(sema: dict) -> dict:
    """JSON Schema -> Gemini `response_schema` (OpenAPI alt kümesi).

    `{"type": ["number", "null"]}` -> `{"type": "NUMBER", "nullable": true}`; desteklenmeyen
    anahtarlar atılır.
    """
    out: dict = {}
    tur = sema.get("type")
    if isinstance(tur, list):
        asil = [t for t in tur if t != "null"]
        if "null" in tur:
            out["nullable"] = True
        tur = asil[0] if asil else "string"
    if tur:
        out["type"] = str(tur).upper()
    for k in ("description", "enum", "format"):
        if k in sema:
            out[k] = sema[k]
    if "properties" in sema:
        out["properties"] = {ad: gemini_semasi(alt) for ad, alt in sema["properties"].items()}
    if "required" in sema:
        out["required"] = list(sema["required"])
    if "items" in sema:
        out["items"] = gemini_semasi(sema["items"])
    return out


def _gemini(gorseller: list[Path], model: str, sema: dict, talimat: str) -> dict:
    """Google Gemini API (generateContent): görseller inline_data, çıktı şemalı JSON.

    Anahtar `x-goog-api-key` başlığıyla gider (adreste görünmez). 429/5xx'te 3 kez dener.
    """
    import httpx  # anthropic'in de bağımlılığı; pyproject'te açıkça yazılı

    anahtar = os.environ.get("GEMINI_API_KEY")
    if not anahtar:
        raise RuntimeError("GEMINI_API_KEY tanimli degil (ortam degiskeni ya da .env)")
    parcalar: list[dict] = [{"inline_data": {"mime_type": "image/jpeg",
                                             "data": base64.b64encode(g.read_bytes()).decode()}}
                            for g in gorseller]
    parcalar.append({"text": talimat})
    govde = {"contents": [{"role": "user", "parts": parcalar}],
             "generationConfig": {"response_mime_type": "application/json",
                                  "response_schema": gemini_semasi(sema), "temperature": 0}}
    son_hata: Exception | None = None
    for deneme in range(3):
        try:
            r = httpx.post(GEMINI_URL.format(model=model), json=govde, timeout=600,
                           headers={"x-goog-api-key": anahtar, "Content-Type": "application/json"})
        except httpx.HTTPError as e:
            son_hata = e
        else:
            if r.status_code == 200:
                aday = (r.json().get("candidates") or [{}])[0]
                metin = "".join(p.get("text", "") for p in (aday.get("content") or {}).get("parts", []))
                if not metin:
                    raise RuntimeError(f"Gemini bos cevap (finishReason={aday.get('finishReason')})")
                return json.loads(metin)
            son_hata = RuntimeError(f"Gemini {r.status_code}: {r.text[:300]}")
            if r.status_code not in (429, 500, 502, 503, 504):
                raise son_hata
        if deneme < 2:
            time.sleep(2 * (deneme + 1))
    raise RuntimeError(f"Gemini 3 denemede basarisiz: {son_hata}")


def arka_uc_coz(arka_uc: str, deneme: int) -> str:
    """`karma`yı okuma numarasına göre gerçek arka uca çevirir (tek -> api, çift -> gemini)."""
    if arka_uc not in ARKA_UCLAR:
        raise ValueError(f"bilinmeyen arka uc: {arka_uc} ({' | '.join(ARKA_UCLAR)})")
    if arka_uc == "karma":
        return "api" if deneme % 2 == 1 else "gemini"
    return arka_uc


def _cli(gorseller: list[Path], model: str, sema: dict, talimat: str) -> dict:
    """`claude -p` ile okur (Claude Code oturumu). Görselleri Read aracıyla açtırır."""
    yedek = Path.home() / ".local" / "bin" / ("claude.exe" if os.name == "nt" else "claude")
    exe = shutil.which("claude") or (str(yedek) if yedek.exists() else "claude")
    liste = "\n".join(f"- {g.resolve()}" for g in gorseller)
    istem = (f"{talimat}\n\nSu gorselleri Read araci ile sirayla ac ve oku:\n{liste}\n\n"
             "Sonucu yapilandirilmis cikti olarak dondur.")
    cmd = [exe, "-p", "--model", model, "--allowedTools", "Read",
           "--add-dir", str(gorseller[0].parent.resolve()), "--output-format", "json",
           "--json-schema", json.dumps(sema, ensure_ascii=False), "--no-session-persistence"]
    # Ortamda bir API anahtarı tanımlıysa `claude -p` onu oturum girişine tercih eder.
    # Bu arka uç oturumu kullansın diye anahtar değişkenleri alt sürece aktarılmaz.
    ortam = {k: v for k, v in os.environ.items()
             if k not in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_BASE_URL")}
    r = subprocess.run(cmd, input=istem, capture_output=True, text=True, encoding="utf-8",
                       cwd=gorseller[0].parent, timeout=900, env=ortam)
    if r.returncode != 0:
        raise RuntimeError(f"claude -p hata ({r.returncode}): {r.stderr[-400:]} | {r.stdout[-400:]}")
    cevap = json.loads(r.stdout)
    out = cevap.get("structured_output")
    if out is None:
        raise RuntimeError(f"yapilandirilmis cikti yok: {str(cevap.get('result'))[:500]}")
    return out


def _api(gorseller: list[Path], model: str, sema: dict, talimat: str) -> dict:
    """Anthropic Messages API: görseller + talimat, şemalı zorunlu araç çağrısıyla JSON çıktı."""
    import anthropic  # yalnız api arka ucunda gerekir

    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise RuntimeError("ANTHROPIC_API_KEY tanimli degil (ortam degiskeni ya da .env)")
    icerik: list[dict] = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                           "data": base64.b64encode(g.read_bytes()).decode()}} for g in gorseller]
    icerik.append({"type": "text", "text": talimat})
    r = anthropic.Anthropic().messages.create(
        model=model, max_tokens=API_MAKS_TOKEN,
        tools=[{"name": "sonuc", "description": "Okunan belgeler", "input_schema": sema}],
        tool_choice={"type": "tool", "name": "sonuc"},
        messages=[{"role": "user", "content": icerik}])
    for b in r.content:
        if b.type == "tool_use":
            return b.input
    raise RuntimeError(f"API cevabinda arac cagrisi yok (stop_reason={r.stop_reason})")


def oku_ham(foto: Path, calisma: Path, sema: dict, talimat: str, model: str = VARSAYILAN_MODEL,
            arka_uc: str = VARSAYILAN_ARKA_UC, yeniden: bool = False, deneme: int = 1,
            onek: str = "") -> dict:
    """Bir fotoğrafı verilen şema ve talimatla okur; ham model çıktısını önbellekten ya da modelden döndürür.

    deneme: aynı fotoğrafın bağımsız okuma numarası (çift okuma için ayrı önbellek kaydı).
    arka_uc="karma": tek numaralı okuma Anthropic API, çift numaralı okuma Gemini.
    """
    arka_uc = arka_uc_coz(arka_uc, deneme)
    if arka_uc == "gemini":
        model = gemini_modeli(model)
    calisma.mkdir(parents=True, exist_ok=True)
    anahtar = f"{onek}{foto.stem}_{dosya_ozeti(foto)}_{model}_{arka_uc}" + (f"_o{deneme}" if deneme > 1 else "")
    onbellek = calisma / f"{anahtar}.json"
    if onbellek.exists() and not yeniden:
        return json.loads(onbellek.read_text(encoding="utf-8"))
    if arka_uc == "gemini":
        sonuc = _gemini(tam_gorsel(foto, calisma / anahtar), model, sema, talimat + GEMINI_EK_NOT)
    elif arka_uc == "api":
        sonuc = _api(parcala(foto, calisma / anahtar), api_modeli(model), sema, talimat)
    else:
        sonuc = _cli(parcala(foto, calisma / anahtar), model, sema, talimat)
    sonuc = dict(sonuc, _kaynak=foto.name)  # `z-rapor yenile` fotoğraf adını önbellekten alır
    onbellek.write_text(json.dumps(sonuc, ensure_ascii=False, indent=1), encoding="utf-8")
    return sonuc


def oku(foto: Path, calisma: Path, model: str = VARSAYILAN_MODEL, arka_uc: str = VARSAYILAN_ARKA_UC,
        yeniden: bool = False, deneme: int = 1) -> list[dict]:
    """Bir fotoğraftaki tüm Z raporları.

    Z raporu olmayan belgeler listede `{"_z_degil": [...]}` kaydı olarak döner; isleme ayıklar.
    """
    return raporlara(oku_ham(foto, calisma, SEMA, TALIMAT, model, arka_uc, yeniden, deneme), foto.name)


def raporlara(sonuc: dict, kaynak: str) -> list[dict]:
    """Ham model çıktısı -> rapor listesi (Z olmayan belgeler `_z_degil` kaydı olarak)."""
    raporlar = [dict(r, _kaynak=kaynak) for r in sonuc.get("raporlar", [])]
    if sonuc.get("z_raporu_olmayan"):
        raporlar.append({"_z_degil": sonuc["z_raporu_olmayan"], "_kaynak": kaynak})
    return raporlar
