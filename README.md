# z-rapor-okuyucu

> Reads photos of Turkish cash-register "Z reports" (end-of-day fiscal receipts) with a vision-language model,
> runs every figure through 7 arithmetic and consistency checks, and writes a colour-coded Excel file plus
> import files for the Luca accounting software. Python, MIT. English summary: [README.en.md](README.en.md).

Yazar kasa Z raporu fotoğraflarını bir görsel dil modeliyle okur, rakamları 7 aritmetik ve tutarlılık
kontrolünden geçirir, kontrollü Excel ve Luca içe aktarım dosyaları üretir.

## Sorun

Mali müşavirlik bürolarına her ay müşterilerden kâğıt Z raporları gelir: lokanta, market, kafe... Her gün
bir fiş, her fişte departman, KDV dökümü, ödeme kırılımı ve kümülatif toplamlar. Bunlar elle Excel'e ya da
muhasebe programına girilir. Elle girişte en sık hatalar ciroda tek hane kayması, iptallerin atlanması ve
eksik günün fark edilmemesidir; bunlar beyanname öncesi yakalanmazsa sonradan bulmak zordur.

## Ne yapar

1. Fotoğrafı büyütüp üst üste binen 9 parçaya böler; tam görüntü + parçalar tek seferde modele gider.
2. Model her Z raporunu şemadaki alanlarla döndürür. Okuyamadığını `null` bırakır, tahmin etmez.
3. Her fotoğraf varsayılan olarak **iki kez bağımsız** okunur; iki okuma arasındaki fark uyarı olur.
   `karma` modunda iki okuma iki farklı modelle (Anthropic + Google Gemini) yapılır ve birbirini denetler.
4. Rakamlar 7 kontrolden geçer. Aynı fişin diğer rakamlarından ya da önceki/sonraki Z'nin kümülatif
   toplamlarından **kanıtlanabilen** okuma hataları düzeltilir ve işaretlenir.
5. Excel yazılır (Özet, Detay, Departmanlar, Ödemeler, Kontrol Listesi); istenirse Luca aktarım dosyası.

Ayrıca: masraf fişlerinden "Masraf Listesi" Excel'i, telefondan yükleme yapılan bir web sayfası ve
model çağrısı yapmadan önbellekten yeniden üretme (`z-rapor yenile`).

## Kurulum

Python 3.10 veya üstü:

```bash
git clone https://github.com/witchivip34/z-rapor-okuyucu.git && cd z-rapor-okuyucu
python -m venv .venv && . .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -e .                 # temel
pip install -e ".[web,test]"     # telefon sayfası ve testler de
```

## Hızlı deneme (API'siz)

`ornekler/okuma_ornek.json` uydurma bir işletmenin ("ÖRNEK LOKANTA", VKN 1234567890) 5 ardışık günlük
hazır okumasıdır. Z 0103'te departman tutarı bilerek bozuk, Z 0104'te kümülatif toplam bilerek yanlış
okunmuş. Model çağrısı yapmadan kontrol ve Excel:

```bash
z-rapor dogrula ornekler/okuma_ornek.json -o calisma/ornek.xlsx
# ÖRNEK LOKANTA (1234567890): 5 Z · 1 hata · 0 uyarı · 1 düzeltme
```

Excel'de Z 0103'ün cirosu kırmızı (departman toplamı net satışı tutmuyor), Z 0104'ün kümülatif toplamı
mavi (komşu Z'lerin zinciriyle düzeltildi). `calisma/raporlar.json` kontrol edilmiş veridir; Luca
dosyaları bundan üretilir.

## API ile okuma

```bash
cp ornekler/.env.example .env     # ANTHROPIC_API_KEY=... (karma için GEMINI_API_KEY=... de) doldurun
python ornekler/uret_fis.py       # ornekler/z_ornek.png (uydurma fiş) zaten var; yeniden çizer
z-rapor oku ornekler/z_ornek.png -o calisma/deneme.xlsx
z-rapor oku fotolar/ -o calisma/eylul.xlsx --okuma 2 --paralel 3
```

| Seçenek | Varsayılan | Açıklama |
|---|---|---|
| `--arka-uc` | `api` | `api`: Anthropic API (`ANTHROPIC_API_KEY`, ortam değişkeni ya da `.env`). `cli`: `claude -p` ile Claude Code oturumu, API anahtarı gerekmez. `gemini`: Google Gemini API (`GEMINI_API_KEY`). `karma`: 1. okuma `api`, 2. okuma `gemini` |
| `--model` | `sonnet` | kısa ad (`sonnet`, `opus`) ya da tam model kimliği; eşlemeler `zrapor/ayarlar.py` içinde. Gemini modeli `GEMINI_MODEL` sabitidir (varsayılan `gemini-2.5-pro`) |
| `--okuma` | `2` | her fotoğraf kaç kez bağımsız okunsun |
| `--paralel` | `3` | aynı anda kaç model çağrısı |
| `--yeniden` | kapalı | önbelleği yok say |

### Çapraz okuma (`karma`)

İki anahtar da varsa (`ANTHROPIC_API_KEY` ve `GEMINI_API_KEY`) `--arka-uc karma` önerilir: iki farklı
model aynı hatayı yapmaz, farklar uyarıya düşer. 1. okuma Anthropic API'siyle (tam görüntü + 9 parça),
2. okuma Google'ın Gemini API'siyle (tam görüntü, şemalı JSON çıktı) yapılır; birleştirmede iki okuma
arasındaki her fark sarı uyarı olur, birinin okuyamadığı alan diğerinden dolar. `--okuma 1` verilirse
yalnız Anthropic okuması yapılır. `oku` ve `masraf` komutlarında ve telefon sayfasında çalışır; telefon
sayfası için `.env`'e `ZRAPOR_ARKA_UC=karma` yazın ya da `z-rapor web --arka-uc karma` ile başlatın
(varsayılan `api`; sayfadaki "Hızlı mod" tek okumadır, yalnız Anthropic).

```bash
z-rapor oku fotolar/ -o calisma/eylul.xlsx --arka-uc karma
z-rapor masraf masraf_fotolari/ -o masraf_listesi.xlsx --arka-uc karma
```

Okumalar çıktının yanındaki `_okumalar/` klasöründe fotoğrafın içerik özetiyle önbelleklenir; aynı
fotoğraf ikinci kez API'ye gitmez.

### Önbellekten yeniden üretme (`yenile`)

Kontrol ya da Excel kodu güncellendiyse fotoğrafları yeniden okutmak gerekmez. `yenile` model çağrısı
yapmadan `_okumalar/` önbelleğindeki okumalardan kontrolü, Excel'i ve `raporlar.json`'u yeniden üretir:

```bash
z-rapor yenile calisma/eylul.xlsx                        # önbellek: calisma/_okumalar/
z-rapor yenile masraf_listesi.xlsx --tur masraf --hazirlayan "Ad Soyad"
z-rapor yenile --veri ./calisma --is 20260919-101500-abc123   # telefon sayfasının bir işi
```

Aynı fotoğrafın aynı okuma numarası önbellekte birden çok kez varsa en yenisi kullanılır. Telefon
sayfası çalışıyorsa yeni özeti görmesi için sunucuyu yeniden başlatın.

**Maliyet notu.** Bir okuma = 1 tam görsel + 9 parça, görsel başına yaklaşık 1.600 giriş token'ı: okuma
başına kabaca 16-17 bin giriş ve 2-3 bin çıkış token'ı. Varsayılan çift okumayla fotoğraf başına bunun iki
katı. Bir karede 3 fiş varsa maliyet fişe bölünür. Güncel token fiyatı için Anthropic'in fiyat sayfasına
bakın; tahmin sabitleri `zrapor/ayarlar.py` içinde. `karma` modunda 2. okuma Gemini'ye tek tam görüntü
olarak gider; onun maliyeti Google'ın fiyat sayfasına göredir.

## Telefon sayfası

```bash
# .env içinde ZRAPOR_SIFRE=en-az-8-karakter
z-rapor web --port 8765 --veri ./calisma
```

Tarayıcıda `http://127.0.0.1:8765`: parolayla giriş → fotoğraf çek ya da seç (çözünürlük rozeti: sarı =
muhtemelen mesajlaşma uygulaması küçültmüş) → "Excel'e çevir". İş sırayla işlenir, sayfa kapansa da sürer;
"İşler" listesinden Excel, Luca bilanço ve Luca işletme dosyaları indirilir. Firma başına hesap kodları
sayfadan girilir ve `<veri>/hesaplar.json` dosyasına yazılır. Sayfa yalnız yerelde açılır; dışarıdan
erişmek istersen Cloudflare Tunnel ya da ngrok ile kendin aç.

## Dışarı açma

Sayfa varsayılan olarak yalnız bu bilgisayarda (`127.0.0.1`) dinler. Her seçenekte parola
(`ZRAPOR_SIFRE`, en az 8 karakter) zorunludur ve aynı adresten 5 dakikada 8 yanlış girişten sonra giriş
geçici olarak kilitlenir. Ters vekil ya da tünel arkasında bütün istekler `127.0.0.1`'den gelir; sınırın
gerçek istemci adresine göre işlemesi için `--guvenilir-vekil` verin (adres `X-Forwarded-For`'dan okunur).
Vekil yokken bu bayrağı açmayın, başlık sahtelenebilir. Örnek dosyalar `deploy/` klasöründe; yer
tutucuları (`/YOL/...`, `ORNEK_TUNEL`, `zrapor.ornek.com`) kendi değerlerinizle değiştirin.

**Yerelde (aynı Wi-Fi'daki telefon):**

```bash
z-rapor web --host 0.0.0.0 --port 8765 --veri ./calisma
# telefondan: http://<bilgisayarın yerel IP'si>:8765  (Windows: ipconfig, Linux/macOS: ip a)
```

**Kendi alan adında (sunucu + Caddy, otomatik HTTPS):**

```bash
cp deploy/Caddyfile.ornek Caddyfile                 # zrapor.ornek.com -> kendi alan adınız
sudo cp deploy/zrapor-web.service.ornek /etc/systemd/system/zrapor-web.service   # /YOL'ları düzeltin
sudo systemctl daemon-reload && sudo systemctl enable --now zrapor-web
caddy run --config Caddyfile                        # 127.0.0.1:8765'e ters vekil
```

**Tünelle (port açmadan, Cloudflare Tunnel):**

```bash
cloudflared tunnel login && cloudflared tunnel create ORNEK_TUNEL
cloudflared tunnel route dns ORNEK_TUNEL zrapor.ornek.com
cp deploy/cloudflared.ornek.yml config.yml          # tünel adı, kimlik dosyası, alan adı
cp deploy/baslat.sh.ornek baslat.sh                 # Windows: deploy/baslat.cmd.ornek -> baslat.cmd
./baslat.sh                                         # sunucu + tünel, --guvenilir-vekil ile
```

Kopyaladığınız kişisel dosyalar (`Caddyfile`, `config.yml`, `baslat.sh`, `baslat.cmd` ve `deploy/` altındaki
`.ornek` uzantısız kopyalar) `.gitignore` ile depo dışında kalır.

## Luca aktarımı

```bash
z-rapor luca calisma/raporlar.json --tur bilanco --hesaplar hesaplar.json -o luca_bilanco.xlsx
z-rapor luca calisma/raporlar.json --tur isletme --hesaplar hesaplar.json -o luca_isletme.csv
```

- **Bilanço** (`zrapor/sablonlar/bilanco_fis_aktarim_sablon.xlsx`, Luca "Fiş Aktarım Şablon"): her Z bir
  mahsup fişi. KDV oranı başına 600 (matrah, alacak) + 391 (KDV, alacak); kart → 108 POS, nakit → 100 Kasa
  (borç). Belge türü ZR, evrak no = Z no. Fişler her zaman dengeli (borç = alacak).
- **İşletme defteri** (`zrapor/sablonlar/isletme_hizli_fis_aktarim.csv`, cp1254, `;`, 36 kolon, ondalık
  virgül): Z başına KDV oranı başına bir satır. TUTAR = matrah, KREDİLİ TUTAR = kartla tahsilat (çok oranlı
  Z'de oranlara toplam payıyla bölünür). Faaliyet kodu firma başına verilir.
- **Hesap kodları**: varsayılanlar `zrapor/luca.py` içinde (`100.01.001`, `108.01.001`, `600.01.00x`,
  `391.01.00x`). Her firmanın hesap planı farklıdır; kendi kodlarınızı VKN başına bir JSON dosyasında verin
  (biçim: `ornekler/hesaplar.ornek.json`). Kontrol edilmemiş kod kullanılırsa komut bunu söyler.
- Çözülmemiş kırmızısı olan Z dosyaya **girmez**, nedeniyle listelenir.

Her iki şablon da Luca'da gerçek dönem verisiyle içe aktarılarak doğrulandı. Hesap kodları her büroda
farklıdır; ilk aktarımı kendi deneme firmanızda yapıp kontrol edin.

## 7 kontrol

| # | Kontrol | Bozulursa |
|---|---|---|
| 1 | Departman toplamı = net satış = mali veri toplamı = KDV dökümü toplamı | kırmızı |
| 2 | Ödemeler toplamı (nakit + kart + diğer) = net satış | kırmızı |
| 3 | %X KDV = toplam × X / (100 + X); KDV satırları toplamı = TOPKDV | kırmızı |
| 4 | Banka dökümü toplamı = kredi kartı tahsilatı | sarı |
| 5 | Ödeme adetleri = müşteri fiş adedi; iptal adedi varsa tutarı da var | sarı |
| 6 | Kümülatif zincir: bugünkü KÜM. TOPLAM − dünkü = bugünkü ciro (KDV için de) | kırmızı ya da mavi düzeltme |
| 7 | Z no / tarih ardışıklığı (eksik Z'nin cirosu kümülatiften hesaplanır), fiş adedi aykırılığı | kırmızı / sarı |

Model fişte görmediğini yazmaz; türetme ve düzeltme yalnız kodda (`zrapor/kontrol.py`) yapılır ve her
birinin açıklaması hücre notuna ve "Kontrol Listesi" sayfasına yazılır.

## Renk kodları

| Renk | Anlamı |
|---|---|
| **Kırmızı** | Rakamlar birbirini tutmuyor; fişe bakın |
| **Sarı** | Olağan dışı ya da tahmini; kontrol edin |
| **Mavi** | Okuma hatası önceki ve sonraki Z'den kanıtlanıp düzeltildi |

## Fotoğraf çekme kuralları

Okuma hatalarının çoğu fotoğraftan gelir. Fişler birbirini örttüğünde ya da fotoğraf sıkıştırıldığında
rakamlar karışır.

1. Bir karede **en çok 3 fiş**, fişler **birbirine değmesin**.
2. Düz zemin, gölgesiz ışık, telefon fişe paralel (yukarıdan).
3. Fotoğrafı mesajlaşma uygulamasından **"Belge" olarak** gönderin (sıkıştırılmasın).
4. Ayın son Z'sinden sonraki ilk Z'yi de çekin: zincir son günü de doğrular.

## Masraf listesi

```bash
z-rapor masraf masraf_fotolari/ -o masraf_listesi.xlsx --hazirlayan "Ad Soyad"
```

Market, akaryakıt, restoran fişleri ve e-arşiv faturalardan "MASRAF LİSTESİ" şablonu
(`zrapor/sablonlar/masraf_listesi.xlsx`) doldurulur: tarih, fiş no, satıcı ve %1/%10/%20 KDV tutarları;
matrah ve toplamlar şablondaki formüllerle hesaplanır. Akaryakıt fişleri ayrı sayfaya, eski oranlı (%8/%18)
fişler ayrı sayfaya gider. POS slibi, Z raporu gibi masraf olmayan belgeler atlanır ve listelenir; aynı fiş
iki kez gelirse biri çıkarılır. KDV formülü ve toplam tutarlılığı aynı renk koduyla işaretlenir.
Şablon 80 satırlıktır; sığmayan fişler bildirilir.

## Sınırlamalar

- Model termal baskıda 0/8, 6/0, 5/9, 1/7 hanelerini karıştırabilir. Kontroller çoğunu yakalar ama
  **işaretli her hücre fişle karşılaştırılmalıdır**; araç muhasebecinin yerine geçmez.
- Kümülatif düzeltme en az iki tutarlı komşu Z gerektirir; tek başına gelen fişte yalnız işaretleme yapılır.
- Cihaz düzenleri farklıdır; talimat yaygın düzene göre yazıldı. Farklı bir yazar kasada ilk ayı dikkatle
  kontrol edin.
- HEIC fotoğrafları web sayfası kabul etmez (JPEG'e çevirin).

## Gizlilik

Fotoğraflar yalnızca seçtiğiniz arka uca gider (`api`: Anthropic API; `cli`: kendi Claude Code
oturumunuz; `gemini`: Google Gemini API; `karma`: Anthropic ve Google). Okuma önbelleği, Excel ve hesap kodları yerel diskte kalır. `.gitignore` fotoğrafları,
`calisma/`, `_okumalar/` ve `.env` dosyasını depo dışında tutar; gerçek müşteri verisini depoya koymayın.

## Geliştirme

```bash
pip install -e ".[web,test]"
python -m pytest -q
```

Paket düzeni: `model.py` (şema) · `okuma.py` (görsel → JSON; `api`, `cli`, `gemini`, `karma`) ·
`kontrol.py` (7 kontrol, birleştirme, zincir) · `excel.py` · `luca.py` · `masraf.py` · `yenile.py`
(önbellekten yeniden üretme) · `web.py` + `static/web.html` · `cli.py` · `deploy/` (dışarı açma örnekleri).
Testler model çağırmaz: Anthropic ve Gemini yanıtları sahtelenir.

## Lisans

MIT © 2026 Ahmet Said Kara. Ayrıntı: [LICENSE](LICENSE).
