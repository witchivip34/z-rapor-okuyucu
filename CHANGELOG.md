# Değişiklik günlüğü

## 0.1.0 — 2026-10-19

İlk herkese açık sürüm.

- `zrapor` paketi ve `z-rapor` komutu: `oku`, `dogrula`, `yenile`, `luca`, `masraf`, `web`.
- Okuma: Anthropic Messages API (varsayılan), `claude -p` ve Google Gemini API (`gemini`) arka uçları;
  tam görüntü + 9 parça; çoklu bağımsız okuma; içerik özetli yerel önbellek.
- Çapraz okuma (`--arka-uc karma`): 1. okuma Anthropic, 2. okuma Gemini; iki modelin farkları uyarı olur.
  `oku`, `masraf` ve telefon sayfasında (`ZRAPOR_ARKA_UC`) çalışır. Gemini modeli `ayarlar.GEMINI_MODEL`.
- `z-rapor yenile`: model çağrısı yapmadan `_okumalar/` önbelleğinden kontrol, Excel ve `raporlar.json`
  (ya da telefon sayfasının bir işini `--veri ... --is <id>` ile) yeniden üretir.
- Kontroller: 7 kontrol grubu, kümülatif zincir uzlaştırma (kanıtlanan okuma hatası düzeltilir),
  firma (VKN) gruplama, tekrar ayıklama.
- Excel: Özet, Detay, Departmanlar, Ödemeler, Kontrol Listesi; kırmızı/sarı/mavi işaretleme.
- Luca: bilanço fiş aktarım şablonu ve işletme defteri CSV'si; firma başına hesap kodu dosyası.
  Her iki şablon Luca'da gerçek dönem verisiyle içe aktarılarak doğrulandı.
- Masraf fişi listesi ve parola korumalı telefon yükleme sayfası; `--host` ile ağa açma,
  `--guvenilir-vekil` ile ters vekil/tünel arkasında gerçek istemci adresine göre giriş sınırı.
- `deploy/`: Cloudflare Tunnel, Caddy, systemd ve başlatma betiği örnekleri (yer tutuculu).
- API'siz sentetik örnekler ve pytest testleri (Anthropic ve Gemini yanıtları sahte).
