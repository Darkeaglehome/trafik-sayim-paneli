# Trafik Sayım ve Yoğunluk Paneli

Kavşak bazlı **canlı trafik sayımı ve yoğunluk izleme** uygulaması. Harita üzerinde işaretlediğiniz kavşak kolları için gerçek yol mesafesi/süresi (OSRM) veya canlı trafik hızı (TomTom / HERE / Google) çekilir; bu hızdan yola çıkılarak koldan geçen araç sayısı, kuyruk yükü ve saatlik toplamlar hesaplanır. Sonuçlar SQLite'a kaydedilir ve Excel raporu (pivot tablo dahil) olarak indirilebilir.

- Web paneli: Flask + Leaflet harita (`app.py`, `templates/index.html`)
- Masaüstü sürümü: Tkinter arayüzü (`trafik_arayuz.py`)
- Geliştirici: [Darkeaglehome](https://github.com/Darkeaglehome)

> English version: [README.en.md](README.en.md)

---

## İçindekiler

- [Özellikler](#özellikler)
- [Nasıl Kullanılır?](#nasıl-kullanılır)
- [Kurulum](#kurulum)
- [Çalıştırma](#çalıştırma)
- [Trafik Sağlayıcıları](#trafik-sağlayıcıları)
- [Hesaplama Modeli](#hesaplama-modeli)
- [API Uç Noktaları](#api-uç-noktaları)
- [Veritabanı](#veritabanı)
- [Excel Çıktısı](#excel-çıktısı)
- [Canlıya Alma (Deploy)](#canlıya-alma-deploy)
- [Proje Yapısı](#proje-yapısı)
- [Lisans](#lisans)

---

## Özellikler

- **Haritadan kol seçimi:** Yolun üzerine tıklayarak kavşak kolu ekleme; kollar otomatik olarak pusula yönüne göre (Kuzey, GüneyDoğu, Batı_2 …) adlandırılır.
- **Gerçek yol verisi:** OSRM Table API ile tüm kollar arasında N×N mesafe/süre matrisi, rota geometrisi ile birlikte alınır.
- **Canlı trafik desteği:** TomTom, HERE ve Google Routes için API anahtarı girilirse gerçek anlık hız kullanılır; anahtar yoksa ücretsiz OSRM verisiyle çalışır.
- **Araç sayımı:** Hıza bağlı temel diyagram (Greenshields) modeliyle anlık geçen araç, şerit sayısı, kuyruk yükü ve saatlik toplamlar.
- **Sinyalize kavşak anahtarı:** Kavşakta trafik sinyali varsa yeşil süre payı (%45) uygulanır, yoksa araçlar kesintisiz geçer.
- **Saatlik sayaç:** Paneldeki "Bu Saat Geçen Araç" değeri ile Excel'deki değerler birebir aynı kaynaktan üretilir.
- **Excel raporu:** Saat bazlı sayım, cadde bazlı özet, genel özet, ham trafik logları ve gerçek Excel pivot tablosu.
- **Dayanıklı izleme döngüsü:** Temizle/Durdur işlemleri eski arka plan işlerini `generation` sayacı ile iptal eder; veri geri gelmez.

---

## Nasıl Kullanılır?

1. **Paneli açın** → `http://127.0.0.1:5000`
2. **Kavşağı bulun:** Enlem/boylam girin veya "Yer Ara" kutusuna adres yazın (örn. *Sancaktepe Atatürk Caddesi*).
3. **Kolları işaretleyin:** Haritada **yolun/şeridin üzerine, o kolun gidiş yönünde tıklamanız yeterli.** Her tıklama bir nokta ekler ve noktanın altındaki cadde adı otomatik bulunur.
   - Bir kavşağın 3–4 kolunu temsil edecek şekilde, her kolun üzerine bir kez tıklayın.
   - Aynı noktayı iki kez tıklarsanız mükerrer kayıt oluşmaz.
   - Yanlış tıkladıysanız "İşaretli Noktalar" satırındaki **Temizle** ile noktaları silin.
4. **"Kolları Ekle"** → işaretlenen noktalar kol listesine eklenir, isimleri gerçek yönlerine göre atanır.
5. **Sağlayıcıyı seçin:** Ücretsiz kullanım için `OSRM` bırakın; TomTom/HERE/Google seçerseniz API anahtarını girin.
6. **"İzlemeyi Başlat"** → kollar için yol ağı hesaplanır ve izleme döngüsü başlar. Tablolar ve harita 5–60 saniyede bir kendini günceller.
7. **"Durdur"** izlemeyi duraklatır (seçili kollar korunur), **"Temizle"** her şeyi sıfırlar.
8. **"Excel'e Aktar"** veya **"Raporu Kaydet"** ile `.xlsx` dosyasını indirin.

> Kısaca: **haritada yolun şerit yönüne tıklamak yeterli** — gerisini uygulama yapar.

Masaüstü sürümü için: `python trafik_arayuz.py` (Tkinter + osmnx gerekir; izleme mantığı aynıdır).

---

## Kurulum

Gereksinim: **Python 3.10+**

```bash
git clone https://github.com/Darkeaglehome/trafik-sayim-paneli.git
cd trafik-sayim-paneli

python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

pip install -r requirements.txt
```

`trafik_arayuz.py` (masaüstü sürümü) ek olarak `osmnx` ve Tkinter ister:

```bash
pip install osmnx
```

---

## Çalıştırma

```bash
python app.py
```

Ardından tarayıcıdan `http://127.0.0.1:5000` adresini açın.

Ortam değişkenleri:

| Değişken | Varsayılan | Açıklama |
|---|---|---|
| `PORT` | `5000` | Dinlenecek port |
| `HOST` | `0.0.0.0` | Dinlenecek adres |
| `FLASK_DEBUG` | `0` | `1` ise debug modu |
| `DATA_DIR` | proje klasörü | SQLite veritabanının tutulacağı klasör |

```bash
# örnek
PORT=8000 DATA_DIR=/var/data python app.py
```

---

## Trafik Sağlayıcıları

| Sağlayıcı | Anahtar | Veri | Not |
|---|---|---|---|
| OSRM | Gerekmez | Yol mesafesi, süresi, geometrisi | Ücretsiz genel sunucu: `router.project-osrm.org` |
| TomTom | Gerekir | Anlık hız + serbest akış hızı | Flow Segment Data API |
| HERE | Gerekir | Bbox içindeki akım hızı | Traffic Flow API |
| Google | Gerekir | Rota süresi/mesafesinden türetilen hız | Routes API (`TRAFFIC_AWARE`) |

Anahtar girilmezse veya sağlayıcı yanıt vermezse uygulama otomatik olarak OSRM verisine düşer. API anahtarları sunucuda saklanmaz, yalnızca istek sırasında kullanılır.

---

## Hesaplama Modeli

Koldan geçen araç sayısı, gerçek hızdan türetilir (Greenshields temel diyagramı):

```
r        = anlık_hız / serbest_akış_hızı        (0..1)
yoğunluk = k_jam * (1 - r),  k_jam = 160 araç/km/şerit
akış     = yoğunluk * anlık_hız * 3.6           (araç/saat/şerit)
geçen    = akış / 3600 * 10 sn * şerit * yeşil_payı * saat_faktörü
```

- Yol yavaşladıkça yoğunluk artar, akış önce yükselir sonra kilitlenmeye yakın düşer.
- Fiziksel üst sınır: kolda aynı anda bulunabilecek araçtan fazlası geçemez.
- `yeşil_payı`: sinyalize kavşakta `0.45`, sinyalsizde `1.0`.
- `saat_faktörü`: puant saatlerde (07–09, 17–19) `1.3`, gece `0.6`.

---

## API Uç Noktaları

| Metot | Yol | Açıklama |
|---|---|---|
| `GET` | `/` | Web paneli |
| `POST` | `/api/detect` | İşaretlenen noktalardan kolları tespit eder (izlemeyi başlatmaz) |
| `POST` | `/api/add_arm` | Tek bir kol ekler |
| `POST` | `/api/start` | İzlemeyi başlatır (`coords`, `provider`, `api_key`, `signalized`) |
| `POST` | `/api/stop` | İzlemeyi durdurur (kollar korunur) |
| `POST` | `/api/clear` | Tüm durumu sıfırlar |
| `POST` | `/api/signalize` | Çalışırken sinyalize durumunu değiştirir |
| `GET` | `/api/status` | Canlı durum, kollar, sayaçlar, saatlik veriler |
| `GET` | `/api/export` | Detaylı Excel (log + saatlik + pivot) |
| `GET` | `/api/report` | Saatlik özet Excel raporu |

Örnek:

```bash
curl -X POST http://127.0.0.1:5000/api/start \
  -H "Content-Type: application/json" \
  -d '{"coords": [[41.0083, 28.9780], [41.0080, 28.9788]], "provider": "osrm", "signalized": true}'
```

---

## Veritabanı

`osrm_intersection_traffic.db` — SQLite, tek tablo:

```sql
CREATE TABLE live_traffic_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    arm_name TEXT NOT NULL,
    distance_m REAL,
    duration_s REAL,
    estimated_queue_load INTEGER,
    data_source TEXT,
    vehicle_count INTEGER
);
```

Kayıtlar 60 saniyelik pencerelerde birikir; `vehicle_count` o pencerede gerçekten geçen araç toplamıdır. Yeni izleme başlatıldığında tablo temizlenir.

---

## Excel Çıktısı

`/api/export` ve `/api/report` şu sayfaları üretir:

1. **Saat Bazlı Sayım** — kollar × saat aralıkları, genel toplam satırıyla (panelle birebir aynı).
2. **Cadde Bazlı Özet** — Tarih | Saat Aralığı | Cadde Adı | Araç Sayısı.
3. **Özet** — kayıt sayısı, toplam araç, kollar, ilk/son kayıt.
4. **Trafik Logları** — 60 saniyelik ham kayıtlar.
5. **Pivot** — gerçek Excel pivot tablosu (satır: saat aralığı, sütun: cadde, değer: toplam araç).

> Pivot sayfası Excel COM üzerinden eklenir; bu özellik yalnızca **Windows + Excel kurulu** ortamda oluşur. Diğer platformlarda dosya pivot olmadan, diğer tüm sayfalarla birlikte üretilir.

---

## Canlıya Alma (Deploy)

Uygulama tek süreçte çalışır (içinde arka plan izleme thread'i ve SQLite var), bu yüzden **tek worker** ile yayınlanmalıdır.

### Render / benzeri PaaS

Depoda hazır `render.yaml` ve `Procfile` bulunur:

- Build: `pip install -r requirements.txt`
- Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 8 --timeout 300`
- Kalıcı disk: `/var/data` → `DATA_DIR=/var/data`

Adımlar:

1. GitHub'da `Darkeaglehome` hesabında bir depo oluşturup kodu gönderin.
2. Barındırma panelinde **New Web Service** → depoyu bağlayın.
3. Start komutunu ve `DATA_DIR` değişkenini ayarlayın, kalıcı disk ekleyin.
4. Deploy sonrası verilen `https://...` adresini açın.

### Kendi sunucunuz (VPS)

```bash
pip install -r requirements.txt
DATA_DIR=/var/data gunicorn app:app \
  --bind 0.0.0.0:8000 --workers 1 --threads 8 --timeout 300
```

Önüne Nginx reverse proxy koyun. SQLite ve arka plan thread'i nedeniyle `--workers` değerini **1**'den fazla yapmayın; aksi halde her worker kendi ayrı durumunu tutar.

---

## Proje Yapısı

```
.
├── app.py                     # Flask web sunucusu + izleme döngüsü + Excel üretimi
├── trafik_arayuz.py           # Tkinter masaüstü sürümü
├── templates/index.html       # Leaflet haritalı web paneli
├── osrm_intersection_traffic.db
├── requirements.txt
├── Procfile                   # PaaS start komutu
├── render.yaml                # Render blueprint
└── README.md
```

---

## Lisans

MIT © [Darkeaglehome](https://github.com/Darkeaglehome)
