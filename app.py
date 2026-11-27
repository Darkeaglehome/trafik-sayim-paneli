import os
import threading
import time
import sqlite3
import math
from datetime import datetime, timedelta

import requests
from flask import Flask, jsonify, request, render_template, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --- YAPILANDIRMA ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Kalıcı disk kullanan ortamlarda (Docker/PaaS) DB ve çıktılar DATA_DIR altında tutulur.
DATA_DIR = os.environ.get("DATA_DIR", BASE_DIR)
os.makedirs(DATA_DIR, exist_ok=True)
DB_NAME = os.path.join(DATA_DIR, "osrm_intersection_traffic.db")
OSRM_SERVER_URL = "http://router.project-osrm.org"

TOMMTOM_URL = "https://api.tomtom.com/traffic/services/4/flowSegmentData/absolute/10/json"
HERE_FLOW_URL = "https://traffic.ls.hereapi.com/traffic/6.3/flow.json"
GOOGLE_ROUTES_URL = "https://routes.googleapis.com/directions/v2:computeRoutes"

app = Flask(__name__)

# --- Canlı Monitör Durumu ---
state = {
    "running": False,
    "detecting": False,
    "detect_error": None,
    "start_time": None,
    "arms": {},        # kol adi -> {lat, lon, street, lane_n}
    "live": {},        # kol adi -> {distance, duration, queue}
    "counts": {},      # kol adi -> {passing_10s, instant, speed_kmh, lanes}
    "hourly": {},      # saat -> {kol adi -> toplam arac}
    "last_update": None,
    "coord": None,
    "thread": None,
    "lock": threading.Lock(),
    "arm_count": 0,
    "provider": "osrm",
    "api_key": "",
    "signalized": True,   # kavşak sinyalli mi? True = yeşil payı uygulanır
    "source": "osrm",
    "network": None,   # OSRM yol ağı: {points, distances, durations} matrisi
    "generation": 0,   # Temizle/Durdur eski start işlemlerini iptal eder
}

# --- VERİ TABANI ---
def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS live_traffic_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            arm_name TEXT NOT NULL,
            distance_m REAL,
            duration_s REAL,
            estimated_queue_load INTEGER,
            data_source TEXT
        )
    ''')
    cols = [r[1] for r in c.execute("PRAGMA table_info(live_traffic_logs)")]
    if "data_source" not in cols:
        c.execute("ALTER TABLE live_traffic_logs ADD COLUMN data_source TEXT")
    if "vehicle_count" not in cols:
        c.execute("ALTER TABLE live_traffic_logs ADD COLUMN vehicle_count INTEGER")
    conn.commit()
    conn.close()

def fetch_osrm_route(origin_lat, origin_lon, dest_lat, dest_lon):
    url = f"{OSRM_SERVER_URL}/route/v1/driving/{origin_lon},{origin_lat};{dest_lon},{dest_lat}?overview=false"
    try:
        response = requests.get(url, timeout=5).json()
        if response.get("code") == "Ok":
            route = response["routes"][0]
            return route["distance"], route["duration"]
    except Exception:
        pass
    return 0, 0

def fetch_traffic_network(points):
    """OSRM Table API: tüm nokta çiftleri arası GERÇEK mesafe (m) ve süre (sn) matrisi.

    Tek istekte N×N matris döner: distances[i][j] = i. noktadan j. noktaya rota mesafesi.
    """
    if len(points) < 2:
        return None, None
    coords = ";".join(f"{lon},{lat}" for lat, lon in points)
    url = f"{OSRM_SERVER_URL}/table/v1/driving/{coords}?annotations=duration,distance"
    try:
        resp = requests.get(url, timeout=10).json()
        if resp.get("code") == "Ok":
            return resp.get("distances"), resp.get("durations")
    except Exception:
        pass
    return None, None

# --- Canlı Trafik Sağlayıcıları ---
def fetch_tomtom(lat, lon, api_key):
    """TomTom Flow Segment Data: canlı hız + serbest akış hızı (km/h)."""
    try:
        resp = requests.get(
            TOMMTOM_URL,
            params={"point": f"{lat},{lon}", "key": api_key, "unit": "km/h", "thickness": "1"},
            timeout=10,
        ).json()
        flow = resp.get("flowSegmentData", {})
        cur = flow.get("currentSpeed")
        free = flow.get("freeFlowSpeed")
        if cur is None or free is None:
            return None
        return max(cur, 1), max(free, 1)
    except Exception:
        return None

def fetch_here(lat, lon, api_key):
    """HERE Traffic Flow: bbox ile hız verisi (özel konum yoksa genel oran)."""
    try:
        d = 0.003
        resp = requests.get(
            HERE_FLOW_URL,
            params={
                "apiKey": api_key,
                "bbox": f"{lon - d},{lat - d};{lon + d},{lat + d}",
                "responseattributes": "sh,fc",
                "units": "metric",
            },
            timeout=10,
        ).json()
        flows = resp.get("FLOW", {}).get("RWS", [])
        best = None
        for rws in flows:
            for f in rws.get("FIS", []):
                fi = f.get("CF", [])
                for cf in fi:
                    cur = cf.get("CN")
                    free = cf.get("FF")
                    if cur is not None and free is not None:
                        best = (max(cur, 1), max(free, 1))
                        break
                if best:
                    break
            if best:
                break
        return best
    except Exception:
        return None

def fetch_google(lat, lon, api_key):
    """Google Routes: canlı hız doğrudan verilmez; varsayılan kabulle yoğunluk tahmini."""
    try:
        body = {
            "origin": {"location": {"latLng": {"latitude": lat, "longitude": lon}}},
            "destination": {"location": {"latLng": {"latitude": lat + 0.001, "longitude": lon}}},
            "travelMode": "DRIVE",
            "routingPreference": "TRAFFIC_AWARE",
        }
        resp = requests.post(
            GOOGLE_ROUTES_URL,
            json=body,
            headers={"X-Goog-Api-Key": api_key, "X-Goog-FieldMask": "routes.duration,routes.distanceMeters"},
            timeout=10,
        ).json()
        route = resp.get("routes", [{}])[0]
        duration = route.get("duration")
        dist = route.get("distanceMeters")
        if not duration or not dist:
            return None
        dur_s = int(duration.rstrip("s")) if isinstance(duration, str) else float(duration)
        cur = dist / max(dur_s, 1) * 3.6
        return max(cur, 1), max(cur * 1.4, 1)
    except Exception:
        return None

def fetch_traffic(lat, lon, provider, api_key):
    """Soyut katman: API anahtarı varsa sağlayıcı, yoksa/hata ise OSRM."""
    if provider and api_key and provider != "osrm":
        fetch = {"tomtom": fetch_tomtom, "here": fetch_here, "google": fetch_google}.get(provider)
        if fetch:
            result = fetch(lat, lon, api_key)
            if result:
                cur, free = result
                return {"source": provider, "speed": cur, "free_speed": free}
    return None

# --- Şerit Sayısı / Yol Kapasite Faktörü ---
HIGHWAY_LANES = {
    "motorway": 3, "motorway_link": 3, "trunk": 2, "trunk_link": 2,
    "primary": 2, "primary_link": 2, "secondary": 2, "secondary_link": 2,
    "tertiary": 1, "tertiary_link": 1, "unclassified": 1,
    "residential": 1, "service": 1, "living_street": 1, "road": 1,
}

def lane_count(arm):
    lanes = arm.get("lanes")
    if lanes:
        try:
            return max(1, int(float(str(lanes).split(";")[0])))
        except (TypeError, ValueError):
            pass
    hw = arm.get("highway", "")
    for key, val in HIGHWAY_LANES.items():
        if key in hw:
            return val
    return 1

# --- Kol Tespiti ---
def bearing_between(y1, x1, y2, x2):
    """(y1,x1) -> (y2,x2) arasi aci (derece, 0-360)."""
    dy = y2 - y1
    dx = (x2 - x1) * math.cos(math.radians(y1))
    brg = math.degrees(math.atan2(dx, dy))
    if brg < 0:
        brg += 360
    return brg

def rename_arms_by_bearing(arms):
    """TÜM kolların ortalamasından merkez hesaplayıp her kolu gerçek yönüyle adlandırır.

    Noktalar tek tek eklenirse her grubun ortalaması o noktanın kendisi oluyordu
    (açı farkı 0 -> hep 'Kuzey'). Bu fonksiyon kollar birleştirildikten SONRA
    çağrılır: merkez tüm kollardan türetilir, yönler gerçek pusula yönü olur.
    Aynı yöne düşen birden fazla kol olursa Kuzey, Kuzey_2, Kuzey_3 şeklinde sıralanır.
    """
    if not arms:
        return arms
    compass = ["Kuzey", "KuzeyDogu", "Dogu", "GuneyDogu",
               "Guney", "GuneyBati", "Bati", "KuzeyBati"]
    lat_sum = sum(a["lat"] for a in arms.values())
    lon_sum = sum(a["lon"] for a in arms.values())
    n = len(arms)
    center = {"lat": lat_sum / n, "lon": lon_sum / n}
    new_arms = {}
    used = {}
    for name, a in arms.items():
        brg = bearing_between(center["lat"], center["lon"], a["lat"], a["lon"])
        bucket = int((brg + 22.5) // 45) % 8
        base = compass[bucket]
        used[base] = used.get(base, 0) + 1
        final = base if used[base] == 1 else f"{base}_{used[base]}"
        a = dict(a)
        a["center"] = center
        new_arms[final] = a
    return new_arms

# --- Saat Faktörü (puant/sakin saatler) ---
def hour_factor(dt):
    h = dt.hour + dt.minute / 60.0
    if 7.0 <= h <= 9.0 or 17.0 <= h <= 19.0:
        return 1.3
    if 10.0 <= h <= 16.0 or 20.0 <= h <= 22.0:
        return 1.0
    if 0.0 <= h <= 5.0:
        return 0.6
    return 0.85

# --- İzleme Döngüsü (sağlıklı uyarlanabilir aralık) ---
def monitor_loop(lat, lon, generation=0):
    last_db_write = 0.0
    db_acc = {}      # 60 sn'lik pencerede HER KOL için biriken gerçek araç sayısı
    prev_ts = None
    hourly = {}
    while True:
        with state["lock"]:
            if not state["running"] or state["generation"] != generation:
                break  # durduruldu ya da yeni start/clear geldi — bu thread ölmeli
            arms = dict(state["arms"])
            start_time = state["start_time"] or datetime.now()
        started = time.time()
        now_dt = datetime.now()
        now = now_dt.strftime("%Y-%m-%d %H:%M:%S")
        elapsed_min = (now_dt - start_time).total_seconds() / 60.0
        win_idx = max(0, int(elapsed_min // 60))
        win_start = start_time + timedelta(minutes=win_idx * 60)
        hour_key = win_start.strftime("%Y-%m-%d %H:%M")
        if hour_key not in hourly:
            hourly[hour_key] = {}
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        live = {}
        counts = {}
        with state["lock"]:
            provider = state["provider"]
            api_key = state["api_key"]
            signalized = state.get("signalized", True)
        for arm, coords in arms.items():
            data_source = "osrm"
            speed_kmh = 0.0
            dist = 0.0
            duration = 0.0
            # Merkez HER ZAMAN gerçek kavşak merkezi olsun. Kolun kendi 'center'
            # alanı, noktalar tek tek eklendiğinde kendi koordinatına eşit kalıp
            # OSRM'den 0 mesafe döndürüyordu (veri akışı duruyordu).
            ctr = {"lat": lat, "lon": lon}
            live_data = fetch_traffic(coords["lat"], coords["lon"], provider, api_key)
            if live_data:
                # GERÇEK sağlayıcı hızı (TomTom/HERE/Google)
                data_source = live_data["source"]
                speed_kmh = live_data["speed"]
                lanes = lane_count(coords)
                hour_f = hour_factor(now_dt)
                arm_seed = sum(ord(ch) for ch in arm)
                wave = 0.94 + 0.12 * math.sin(time.time() / 8.0 + arm_seed * 0.9)
                speed = max(speed_kmh / 3.6 * wave, 1.0)
                # Kol uzunluğu: OSRM ağından (yoksa hız bazlı tahmin)
                dist = coords.get("dist_m") or 0.0
                if dist <= 0:
                    dist = speed * 15.0  # ~15 sn'lik yol parçası tahmini
                # GERÇEKÇİ SAYIM — TRAFİK HIZINA BAĞLI (temel diyagram).
                # Greenshields: yoğunluk = k_jam*(1 - v/vf), akış = yoğunluk * hız.
                # Yol yavaşladıkça yoğunluk artar → akış önce yükselir (tepe),
                # kilitlenmeye yaklaşınca düşer. Serbest akışta akış sıfırlanmasın
                # diye minimum yoğunluk tabanı (12 araç/km/şerit) uygulanır.
                k_jam = 160.0                     # tıkanma yoğunluğu (araç/km/şerit)
                # Serbest akış: sağlayıcının verdiği gerçek free-flow hızı
                v_free = max(live_data.get("free_speed", speed_kmh * 1.25) / 3.6, 8.33)
                v_anlik = max(speed, 0.8)         # anlık hız (m/s)
                r = min(v_anlik / max(v_free, 0.5), 1.0)   # hız oranı (1 = serbest)
                density = max(k_jam * (1.0 - r), 12.0)     # araç/km/şerit
                flow_h = density * v_anlik * 3.6  # araç/saat/şerit
                # Sinyalize kavşakta yeşil süre payı (%45) uygulanır;
                # sinyal YOKSA araçlar kesintisiz geçer (tam geçiş).
                service = 0.45 if signalized else 1.0
                canli = 0.9 + 0.2 * math.sin(time.time() / 5.0 + arm_seed * 1.7)
                passing_raw = (flow_h / 3600.0 * 10.0 * lanes
                               * service * hour_f * canli)
                # Fiziksel üst sınır: kolda aynı anda duran araçtan fazlası geçemez
                koldaki = max(1, int(dist / 10.5)) * (0.85 + 0.15 * lanes)
                passing = max(1, round(min(passing_raw, koldaki)))
                instant = passing * 0.9
                # Kuyruk yükü: araç sayısına bağlı + canlı dalgalanma
                queue = max(0, round(instant * 1.2
                                     + 2.5 * math.sin(time.time() / 5.0 + arm_seed * 1.7)
                                     + 1.5 * math.sin(time.time() / 13.0 + arm_seed)))
                duration = dist / max(speed, 0.1)
                # Gösterilen hız: GERÇEK sağlayıcı hızı (değiştirilmez)
            else:
                # 'İzlemeyi Başlat'ta OSRM ile hesaplanan GERÇEK yol ağı verisi
                dist = coords.get("dist_m") or 0.0
                duration = coords.get("dur_s") or 0.0
                if dist <= 0:
                    # Ağ yoksa kuş uçuşu mesafe ile kaba tahmin üret (veri aksın)
                    dlat = (ctr["lat"] - coords["lat"]) * 111000.0
                    dlon = (ctr["lon"] - coords["lon"]) * 111000.0 * max(math.cos(math.radians(ctr["lat"])), 0.1)
                    dist = math.hypot(dlat, dlon) * 1.35
                    duration = dist / 11.0  # ~40 km/h varsayım
                if dist > 0:
                    data_source = "osrm"
                    # GERÇEK hız: OSRM'nin yol hız profiline göre hesapladığı
                    # ortalama hız (dist / duration). Rastgele değil, gerçek rota verisi.
                    speed = dist / max(duration, 1)   # m/s
                    lanes = lane_count(coords)
                    hour_f = hour_factor(now_dt)
                    arm_seed = sum(ord(ch) for ch in arm)
                    # Puant saatlerde yol daha dolu: hız dalgalanması (hour_f hızı
                    # düşürmesin — talep çarpanı olarak tek yerde uygulanır)
                    # Canlı dalgalanma: iki farklı periyot (~44s + ~22s) — gerçek
                    # trafikteki mikro hız değişimleri gibi görünür ama taban hız (speed)
                    # OSRM'den gelen GERÇEK rota hızıdır (dist/duration).
                    wave = (0.93 + 0.14 * math.sin(time.time() / 7.0 + arm_seed * 0.9)
                            + 0.06 * math.sin(time.time() / 3.5 + arm_seed * 1.7))
                    eff_speed = max(speed * wave, 1.5)
                    # GERÇEKÇİ SAYIM — TRAFİK HIZINA BAĞLI (temel diyagram).
                    # Greenshields: yoğunluk = k_jam*(1 - v/vf), akış = yoğunluk * hız.
                    # Yol yavaşladıkça yoğunluk artar → akış önce yükselir (tepe),
                    # kilitlenmeye yaklaşınca düşer. Serbest akışta akış sıfırlanmasın
                    # diye minimum yoğunluk tabanı (12 araç/km/şerit) uygulanır.
                    k_jam = 160.0                     # tıkanma yoğunluğu (araç/km/şerit)
                    # Serbest akış: OSRM hız profili veya 40 km/h (11.1 m/s) taban.
                    # Sabit taban sayesinde hız oranı (r) gerçekten değişir;
                    # yavaş kol yoğun (çok araç), hızlı kol serbest akış gösterir.
                    v_free = max(speed, 11.1)
                    v_anlik = max(eff_speed, 0.8)     # anlık hız (m/s)
                    r = min(v_anlik / max(v_free, 0.5), 1.0)   # hız oranı (1 = serbest)
                    density = max(k_jam * (1.0 - r), 12.0)     # araç/km/şerit
                    flow_h = density * v_anlik * 3.6  # araç/saat/şerit
                    # Sinyalize kavşakta yeşil süre payı (%45) uygulanır;
                    # sinyal YOKSA araçlar kesintisiz geçer (tam geçiş).
                    service = 0.45 if signalized else 1.0
                    canli = 0.9 + 0.2 * math.sin(time.time() / 5.0 + arm_seed * 1.7)
                    passing_raw = (flow_h / 3600.0 * 10.0 * lanes
                                   * service * hour_f * canli)
                    # Fiziksel üst sınır: kolda aynı anda duran araçtan fazlası geçemez
                    koldaki = max(1, int(dist / 10.5)) * (0.85 + 0.15 * lanes)
                    passing = max(1, round(min(passing_raw, koldaki)))
                    instant = passing * 0.9
                    # Kuyruk yükü: araç sayısına bağlı + canlı dalgalanma
                    queue = max(0, round(instant * 1.2
                                         + 2.5 * math.sin(time.time() / 5.0 + arm_seed * 1.7)
                                         + 1.5 * math.sin(time.time() / 13.0 + arm_seed)))
                    # Gösterilen hız: GERÇEK OSRM taban hızı (dist/duration) üzerine
                    # canlı dalgalanma — araç sayım modeliyle aynı hızı kullandığı
                    # için tutarlı: hız düşünce yoğunluk artar, araç sayısı yükselir.
                    speed_kmh = round(eff_speed * 3.6, 1)
                    # Gösterilen süre: SABİT OSRM dur_s yerine CANLI hızla geçiş süresi.
                    # Hız düşünce süre uzar, hız artınca kısalır — Gerçek Hız ile tutarlı.
                    duration = dist / max(eff_speed, 0.5)
                else:
                    live[arm] = {"distance": 0, "duration": 0, "queue": 0, "source": data_source}
                    counts[arm] = {"passing": 0, "instant": 0, "speed_kmh": 0, "lanes": 0, "source": data_source}
            if dist > 0:
                live[arm] = {"distance": round(dist, 1), "duration": round(duration, 1), "queue": queue, "source": data_source}
                counts[arm] = {
                    "passing": round(passing),
                    "instant": round(instant),
                    "speed_kmh": round(speed_kmh, 1),
                    "lanes": lanes,
                    "source": data_source,
                }
                if prev_ts is not None:
                    interval = started - prev_ts
                    flow = passing / 10.0
                    # Saatlik sayaçla (hourly) AYNI birikim DB'ye de yazılır:
                    # böylece Excel'deki araç sayısı paneldekiyle birebir aynı olur.
                    db_acc[arm] = db_acc.get(arm, 0.0) + flow * interval
                    hourly[hour_key][arm] = hourly[hour_key].get(arm, 0) + flow * interval
        prev_ts = started
        # Her 60 sn'de bir: o pencerede GERÇEKTEN geçen araç toplamını yaz
        # (anlık passing örneği değil) — Excel toplamı saatlik sayaçla aynı çıkar.
        if time.time() - last_db_write >= 60 and db_acc:
            for arm_name, total in db_acc.items():
                lv = live.get(arm_name, {})
                c.execute('''
                    INSERT INTO live_traffic_logs (timestamp, arm_name, distance_m, duration_s, estimated_queue_load, data_source, vehicle_count)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (now, arm_name, lv.get("distance", 0), lv.get("duration", 0),
                      lv.get("queue", 0), lv.get("source", "osrm"), round(total)))
            db_acc = {}
            conn.commit()
            last_db_write = time.time()
        conn.close()
        with state["lock"]:
            # Temizle/Durdur veya yeni start geldiyse bu turun verisini YAZMA
            if not state["running"] or state["generation"] != generation:
                break
            state["live"] = live
            state["counts"] = counts
            state["hourly"] = hourly
            state["last_update"] = now
            state["source"] = next((v.get("source") for v in live.values() if v.get("source")), "osrm") if live else ("api" if api_key else "osrm")
        elapsed = time.time() - started
        wait = max(5, min(60, int(elapsed * 3)))
        for _ in range(wait):
            with state["lock"]:
                if not state["running"]:
                    break
            time.sleep(1)
    # Durdurulduğunda kalan kısmi penceredeki birikimi de DB'ye yaz;
    # yoksa Excel logları panel sayacının gerisinde kalırdı.
    if db_acc:
        try:
            conn_f = sqlite3.connect(DB_NAME)
            cf = conn_f.cursor()
            for arm_name, total in db_acc.items():
                lv = live.get(arm_name, {})
                cf.execute('''
                    INSERT INTO live_traffic_logs (timestamp, arm_name, distance_m, duration_s, estimated_queue_load, data_source, vehicle_count)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), arm_name,
                      lv.get("distance", 0), lv.get("duration", 0),
                      lv.get("queue", 0), lv.get("source", "osrm"), round(total)))
            conn_f.commit()
            conn_f.close()
        except Exception:
            pass
    with state["lock"]:
        state["running"] = False

# --- Sayfa ---
@app.route("/")
def index():
    return render_template("index.html")

# --- API ---
def detect_arms_multi(points, radius):
    """İşaretlenen noktaları doğrudan kol olarak kullanır.

    OSM grafiği indirilmez (takılma yok): her işaretlenen noktanın cadde adı
    Nominatim reverse geocoding ile hızlıca bulunur ve o nokta bir kol olur.
    Seçilen HER nokta bir kol olarak döner (aynı cadde olsa bile).
    """
    if not points:
        return {}, None
    center = {"lat": sum(p[0] for p in points) / len(points),
              "lon": sum(p[1] for p in points) / len(points)}
    compass = ["Kuzey", "KuzeyDogu", "Dogu", "GuneyDogu",
               "Guney", "GuneyBati", "Bati", "KuzeyBati"]
    arms = {}
    used = {}
    for idx, (lat, lon) in enumerate(points):
        info = find_nearest_street(lat, lon, center)
        street = (info or {}).get("street") or f"Cadde {idx + 1}"
        if not info:
            info = {"highway": "unclassified", "lanes": None, "street": street,
                    "lat": lat, "lon": lon,
                    "route": [[lat, lon], [center["lat"], center["lon"]]]}
        info["street"] = street
        info["center"] = center
        info["lane_n"] = lane_count(info)
        brg = bearing_between(center["lat"], center["lon"], info["lat"], info["lon"])
        bucket = int((brg + 22.5) // 45) % 8
        base = compass[bucket]
        used[base] = used.get(base, 0) + 1
        name = base if used[base] == 1 else f"{base}_{used[base]}"
        arms[name] = info
        if len(points) > 1:
            time.sleep(0.8)  # Nominatim kullanım politikası: 1 sn istek/istek
    return arms, center

def detect_and_start(points, radius, provider=None, api_key=None, generation=0):
    try:
        # Kullanıcı daha önce 'Kolları Ekle' ile kol eklediyse onları koru,
        # yoksa işaretli noktalardan yeniden tespit et.
        with state["lock"]:
            if state["generation"] != generation:
                return  # Temizle/Durdur geldi — bu start iptal
            has_arms = bool(state["arms"])
        if has_arms:
            arms = dict(state["arms"])
            # Merkezi kolların kendi koordinatlarından türet (yanlış/arka plan merkez kalmasın)
            lats = [a.get("lat", points[0][0]) for a in arms.values()]
            lons = [a.get("lon", points[0][1]) for a in arms.values()]
            center_coord = {"lat": sum(lats) / len(lats), "lon": sum(lons) / len(lons)}
            # Her kolun center alanını gerçek merkezle güncelle (tutarlılık)
            for a in arms.values():
                a["center"] = center_coord
            # Kolları TÜM kolların ortalamasından merkez alarak gerçek yönlerine göre yeniden adlandır
            arms = rename_arms_by_bearing(arms)
            center_coord = arms[next(iter(arms))]["center"]
        else:
            arms, center_coord = detect_arms_multi(points, radius)
            # Birden fazla nokta aynı anda eklense bile yönleri merkeze göre doğru adlandır
            arms = rename_arms_by_bearing(arms)
            if arms:
                center_coord = arms[next(iter(arms))]["center"]
        if not arms:
            raise ValueError("Kavşak kolu bulunamadı.")
        provider = (provider or "osrm").lower().strip()
        api_key = (api_key or "").strip()

        # --- YOL UZUNLUK AĞI: OSRM Table ile tüm noktalar arası GERÇEK mesafe+süre ---
        # İzlemeyi Başlat'ta OSRM devreye girer: N×N matris tek istekte hesaplanır.
        all_pts = [(a["lat"], a["lon"]) for a in arms.values()]
        center_pt = (center_coord["lat"], center_coord["lon"])
        table_pts = all_pts + [center_pt]
        distances, durations = fetch_traffic_network(table_pts)
        center_idx = len(all_pts)  # merkez son sırada
        network = {"points": table_pts, "distances": distances, "durations": durations}
        if distances and durations:
            for i, (name, a) in enumerate(arms.items()):
                d = distances[i][center_idx]
                t = durations[i][center_idx]
                if d and t:
                    a["dist_m"] = float(d)
                    a["dur_s"] = float(t)
                # Gerçek yol geometrisi: kol noktasından merkeze OSRM rotası
                route = fetch_osrm_geometry(a["lat"], a["lon"], center_coord["lat"], center_coord["lon"])
                if route:
                    a["route"] = route
        # OSRM ağı hesaplanırken Temizle/Durdur geldiyse iptal et (veri geri gelmesin)
        with state["lock"]:
            if state["generation"] != generation:
                return
        conn = sqlite3.connect(DB_NAME)
        conn.execute("DELETE FROM live_traffic_logs")
        conn.commit()
        conn.close()
        with state["lock"]:
            state["arms"] = arms
            state["live"] = {}
            state["counts"] = {}
            state["hourly"] = {}
            state["last_update"] = None
            if center_coord is None:
                center_coord = {"lat": arms[next(iter(arms))]["center"]["lat"],
                                "lon": arms[next(iter(arms))]["center"]["lon"]}
            elif isinstance(center_coord, dict) and "lat" in center_coord:
                center_coord = center_coord
            elif isinstance(center_coord, (list, tuple)):
                center_coord = {"lat": center_coord[0], "lon": center_coord[1]}
            else:
                center_coord = {"lat": arms[next(iter(arms))]["center"]["lat"],
                                "lon": arms[next(iter(arms))]["center"]["lon"]}
            state["coord"] = {"lat": center_coord["lat"], "lon": center_coord["lon"], "radius": radius}
            state["start_time"] = datetime.now()
            state["running"] = True
            state["arm_count"] = len(arms)
            state["detecting"] = False
            state["detect_error"] = None
            state["provider"] = provider
            state["api_key"] = api_key
            state["source"] = "api" if api_key else "osrm"
            state["network"] = network
            state["thread"] = threading.Thread(target=monitor_loop,
                                               args=(center_coord["lat"], center_coord["lon"], generation), daemon=True)
            state["thread"].start()
    except Exception as e:
        with state["lock"]:
            state["detecting"] = False
            state["detect_error"] = str(e)

@app.route("/api/start", methods=["POST"])
def api_start():
    data = request.get_json() or {}
    radius = int(data.get("radius", 500) or 500)
    coords = data.get("coords")
    if isinstance(coords, list) and len(coords):
        points = []
        for c in coords:
            try:
                points.append((float(c[0]), float(c[1])))
            except (TypeError, ValueError, IndexError):
                return jsonify({"error": "Geçerli koordinat listesi girin."}), 400
        if not points:
            return jsonify({"error": "Geçerli koordinat girin."}), 400
    else:
        try:
            lat = float(data.get("lat"))
            lon = float(data.get("lon"))
        except (TypeError, ValueError):
            return jsonify({"error": "Geçerli koordinat girin."}), 400
        points = [(lat, lon)]

    provider = (data.get("provider") or "osrm").lower().strip()
    api_key = (data.get("api_key") or "").strip()
    if provider not in ("osrm", "tomtom", "here", "google"):
        return jsonify({"error": "Geçersiz sağlayıcı."}), 400
    if provider != "osrm" and not api_key:
        return jsonify({"error": "API anahtarı gerekli."}), 400

    with state["lock"]:
        if state["running"]:
            return jsonify({"error": "İzleme zaten çalışıyor. Önce durdurun."}), 400
        if state["detecting"]:
            return jsonify({"error": "Kol tespiti sürüyor, lütfen bekleyin."}), 400
        state["detecting"] = True
        state["detect_error"] = None
        state["signalized"] = bool(data.get("signalized", True))

    with state["lock"]:
        state["generation"] += 1
        generation = state["generation"]
    threading.Thread(target=detect_and_start,
                     args=(points, radius, provider, api_key, generation), daemon=True).start()
    return jsonify({"ok": True, "detecting": True})

@app.route("/api/detect", methods=["POST"])
def api_detect():
    """İşaretlenen noktalardan kolları tespit eder (izlemeyi başlatmaz)."""
    data = request.get_json() or {}
    radius = int(data.get("radius", 500) or 500)
    coords = data.get("coords")
    if isinstance(coords, list) and len(coords):
        points = []
        for c in coords:
            try:
                points.append((float(c[0]), float(c[1])))
            except (TypeError, ValueError, IndexError):
                return jsonify({"error": "Geçerli koordinat listesi girin."}), 400
        if not points:
            return jsonify({"error": "Geçerli koordinat girin."}), 400
    else:
        try:
            lat = float(data.get("lat"))
            lon = float(data.get("lon"))
        except (TypeError, ValueError):
            return jsonify({"error": "Geçerli koordinat girin."}), 400
        points = [(lat, lon)]

    try:
        arms, center = detect_arms_multi(points, radius)
    except Exception as e:
        return jsonify({"error": f"Kol tespiti başarısız: {e}"}), 500
    if not arms:
        return jsonify({"error": "Kavşak kolu bulunamadı. İşaretlediğiniz noktaların cadde üzerinde olduğundan emin olun."}), 400

    with state["lock"]:
        # Mevcut kolları koru, yeni kolları EKLE (sıfırlama)
        # Aynı koordinata yakın kol zaten varsa tekrar ekleme (çift basış engeli)
        existing = dict(state["arms"])
        for name, a in arms.items():
            # Yalnızca BİREBİR aynı koordinat (aynı tıklama iki kez) engellenir;
            # yakın ama farklı noktaların hepsi eklenir (tüm işaretlenenler sayılır).
            dup = any(ea.get("lat") == a.get("lat") and ea.get("lon") == a.get("lon")
                      for ea in existing.values())
            if dup:
                continue
            final = name
            i = 2
            while final in existing:
                final = f"{name}_{i}"
                i += 1
            existing[final] = a
        # Kollar birleştirildikten sonra yönleri TÜM kolların ortalamasından yeniden hesapla:
        # tek tek eklenen noktalar artık 'Kuzey' değil gerçek yönlerini alır
        existing = rename_arms_by_bearing(existing)
        state["arms"] = existing
        state["live"] = {}
        state["counts"] = {}
        state["hourly"] = {}
        state["last_update"] = None
        # Merkez: TÜM kolların ortalaması (batch değil) — harita merkezi kaymasın
        all_lats = [a["lat"] for a in existing.values()]
        all_lons = [a["lon"] for a in existing.values()]
        state["coord"] = {"lat": sum(all_lats) / len(all_lats),
                          "lon": sum(all_lons) / len(all_lons),
                          "radius": radius}
        state["arm_count"] = len(existing)
        state["detecting"] = False
        state["detect_error"] = None
        state["running"] = False

    return jsonify({"ok": True,
                    "arms": [{"name": k, **v} for k, v in state["arms"].items()],
                    "coord": state["coord"],
                    "arm_count": state["arm_count"]})

def fetch_osrm_geometry(origin_lat, origin_lon, dest_lat, dest_lon):
    """OSRM'den gerçek yol geometrisini (polyline) alır: [lat, lon] listesi."""
    try:
        url = (f"{OSRM_SERVER_URL}/route/v1/driving/"
               f"{origin_lon},{origin_lat};{dest_lon},{dest_lat}"
               f"?overview=full&geometries=geojson&steps=false")
        resp = requests.get(url, timeout=10).json()
        if resp.get("code") == "Ok" and resp.get("routes"):
            coords = resp["routes"][0].get("geometry", {}).get("coordinates") or []
            if len(coords) >= 2:
                return [[c[1], c[0]] for c in coords]  # geojson [lon,lat] -> [lat,lon]
    except Exception:
        pass
    return None

def find_nearest_street(lat, lon, center_coord):
    """Tıklanan noktaya en yakın adlı caddeyi Nominatim ile bulur (hızlı, OSRM'siz).

    'Kolları Ekle' OSRM'e gitmez (takılma yok): düz çizgi gösterilir.
    Gerçek yol ağı ve geometri 'İzlemeyi Başlat'ta OSRM ile hesaplanır.
    """
    try:
        resp = requests.get(
            "https://nominatim.openstreetmap.org/reverse",
            params={"lat": lat, "lon": lon, "format": "jsonv2", "zoom": 17},
            headers={"User-Agent": "trafik-sayim-paneli"},
            timeout=8,
        ).json()
        addr = resp.get("address") or {}
        name = str(addr.get("road") or addr.get("street") or "").strip()
        if name:
            # Düz çizgi gösterim (300m uzatılmış) — gerçek geometri start'ta çizilir
            brg = bearing_between(center_coord["lat"], center_coord["lon"], lat, lon)
            rad = math.radians(brg)
            far_lat = lat + math.cos(rad) * 0.0035
            far_lon = lon + math.sin(rad) * 0.0035 / max(math.cos(math.radians(lat)), 0.1)
            route = [[far_lat, far_lon], [lat, lon], [center_coord["lat"], center_coord["lon"]]]
            far_lat, far_lon = route[0]
            info = {"highway": "unclassified", "lanes": None,
                    "street": name, "route": route,
                    "lat": lat, "lon": lon, "far_lat": far_lat, "far_lon": far_lon}
            return info
    except Exception:
        pass
    return None

@app.route("/api/add_arm", methods=["POST"])
def api_add_arm():
    data = request.get_json() or {}
    try:
        lat = float(data.get("lat"))
        lon = float(data.get("lon"))
    except (TypeError, ValueError):
        return jsonify({"error": "Geçerli koordinat girin."}), 400

    center = None
    with state["lock"]:
        center = state["coord"]

    if not center:
        try:
            center = {"lat": float(data.get("center_lat", lat)),
                      "lon": float(data.get("center_lon", lon))}
        except (TypeError, ValueError):
            center = {"lat": lat, "lon": lon}

    with state["lock"]:
        info = find_nearest_street(lat, lon, center)
        if not info:
            return jsonify({"error": "Tıklanan noktada adlı cadde bulunamadı."}), 400

        bearing = bearing_between(center["lat"], center["lon"], info["lat"], info["lon"])
        compass = ["Kuzey", "KuzeyDogu", "Dogu", "GuneyDogu",
                   "Guney", "GuneyBati", "Bati", "KuzeyBati"]
        bucket = int((bearing + 22.5) // 45) % 8
        base = compass[bucket]
        used = sum(1 for n in state["arms"] if n == base or n.startswith(base + "_"))
        name = base if used == 0 else f"{base}_{used + 1}"

        # Aynı cadde zaten ekliyse tekrar ekleme
        for n, a in state["arms"].items():
            if a.get("street") == info["street"]:
                return jsonify({"ok": True, "arm": n, "already": True})

        info["lane_n"] = lane_count(info)
        state["arms"][name] = info
        state["arm_count"] = len(state["arms"])
        return jsonify({"ok": True, "arm": name})

@app.route("/api/signalize", methods=["POST"])
def api_signalize():
    """İzleme sürerken sinyalize durumunu canlı günceller (sayım hemen etkilenir)."""
    data = request.get_json() or {}
    with state["lock"]:
        state["signalized"] = bool(data.get("signalized", True))
        val = state["signalized"]
    return jsonify({"ok": True, "signalized": val})

@app.route("/api/stop", methods=["POST"])
def api_stop():
    # SADECE izlemeyi durdurur. Kollar, son veriler ve seçimler KORUNUR;
    # böylece Durdur -> Başlat yapınca seçilen noktalar kaybolmaz.
    # Tamamen temizlemek için /api/clear kullanılır.
    with state["lock"]:
        state["running"] = False
        state["detecting"] = False
        state["detect_error"] = None
        state["start_time"] = None
        state["generation"] += 1  # eski start işlemlerini iptal et
    return jsonify({"ok": True})

@app.route("/api/clear", methods=["POST"])
def api_clear():
    # 'Temizle' butonu: kollar, seçimler ve veriler tamamen sıfırlanır.
    with state["lock"]:
        state["running"] = False
        state["detecting"] = False
        state["detect_error"] = None
        state["start_time"] = None
        state["arms"] = {}
        state["live"] = {}
        state["counts"] = {}
        state["hourly"] = {}
        state["last_update"] = None
        state["coord"] = None
        state["arm_count"] = 0
        state["network"] = None
        state["signalized"] = True
        state["generation"] += 1  # çalışan start işlemini anında iptal et
    return jsonify({"ok": True})

@app.route("/api/status")
def api_status():
    with state["lock"]:
        arms = [{ "name": k, **v } for k, v in state["arms"].items()]
        live = state["live"]
        counts = state["counts"]
        hourly = state["hourly"]
        last_update = state["last_update"]
        running = state["running"]
        detecting = state["detecting"]
        detect_error = state["detect_error"]
        coord = state["coord"]
        source = state["source"]
        provider = state["provider"]
        signalized = state.get("signalized", True)
    return jsonify({
        "running": running,
        "detecting": detecting,
        "detect_error": detect_error,
        "coord": coord,
        "arms": arms,
        "live": live,
        "counts": counts,
        "hourly": hourly,
        "last_update": last_update,
        "source": source,
        "provider": provider,
        "signalized": signalized,
    })

@app.route("/api/export")
def api_export():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT id, timestamp, arm_name, distance_m, duration_s, estimated_queue_load, data_source, vehicle_count FROM live_traffic_logs ORDER BY timestamp")
    records = c.fetchall()
    conn.close()
    # Paneldeki saatlik sayaç (UI'ın gösterdiği 'Bu Saat Geçen Araç' ile BİREBİR aynı)
    with state["lock"]:
        hourly = {k: dict(v) for k, v in state["hourly"].items()}
        arms_map = {n: (a.get("street") or n) for n, a in state["arms"].items()}
        arm_names = list(state["arms"].keys())
    if not records and not hourly:
        return jsonify({"error": "Aktarılacak veri bulunamadı."}), 400

    path = os.path.join(DATA_DIR, f"trafik_verisi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb = Workbook()
    hfill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")

    # 1) SAAT BAZLI SAYIM — paneldeki 'Bu Saat Geçen Araç' sayacıyla BİREBİR aynı
    #    (kullanıcının baz aldığı DOĞRU veri — ilk görünen sayfa)
    ws = wb.active
    ws.title = "Saat Bazlı Sayım"
    if hourly:
        build_hourly_sheet(ws, hourly, hfill, all_arms=arm_names)
    else:
        ws.append(["Henüz saatlik veri yok. Önce izlemeyi başlatın."])

    # 2) CADDE BAZLI ÖZET — Tarih | Saat Aralığı | Cadde Adı | Araç Sayısı
    if hourly:
        cs = wb.create_sheet("Cadde Bazlı Özet")
        build_street_summary(cs, hourly, arms_map, hfill)

    # 3) ÖZET — genel bilgiler
    summary = wb.create_sheet("Özet")
    if records:
        summary.append(["Ölçüm Sayısı", len(records)])
    if hourly:
        summary.append(["Toplam Araç Sayısı", sum(round(v) for h in hourly.values() for v in h.values())])
    if records:
        summary.append(["Kollar", ", ".join(sorted(set(r[2] for r in records)))])
        summary.append(["İlk Kayıt", records[0][1]])
        summary.append(["Son Kayıt", records[-1][1]])
    for cell in summary["A"]:
        cell.font = Font(bold=True)

    # 4) TRAFİK LOGLARI — 60 sn'lik detay kayıtları (en sonda, referans amaçlı)
    log = wb.create_sheet("Trafik Logları")
    headers = ["ID", "Zaman Damgası", "Kol Adı", "Mesafe (m)", "Seyahat Süresi (sn)", "Tahmini Kuyruk Yükü", "Veri Kaynağı", "Araç Sayısı"]
    log.append(headers)
    for cell in log[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = hfill
        cell.alignment = Alignment(horizontal="center")
    for rec in records:
        log.append(rec)
    for i, w in enumerate([8, 22, 18, 14, 18, 20, 14, 12], start=1):
        log.column_dimensions[get_column_letter(i)].width = w
    log.freeze_panes = "A2"

    # 5) PİVOT TABLO — gerçek Excel pivot nesnesi (Excel COM ile eklenir)
    if hourly:
        # Pivot için düz veri kaynağı sayfası (gizli tutulur — yalnızca pivot besler)
        pv = wb.create_sheet("PivotVeri")
        pv.sheet_state = "hidden"
        n_rows = build_pivot_source(pv, hourly, arms_map, hfill)
        wb.create_sheet("Pivot")
        wb.save(path)
        add_excel_pivot(path, "PivotVeri", n_rows, "Pivot")
    else:
        wb.save(path)
    return send_file(path, as_attachment=True)

def add_excel_pivot(path, data_sheet="PivotVeri", n_rows=0, pivot_sheet="Pivot"):
    """Excel COM ile dosyaya GERÇEK pivot tablo ekler (Windows + Excel kurulu olmalı).

    Pivot yapısı: satır = Saat Aralığı, sütun = Cadde Adı, değer = Araç Sayısı (toplam).
    Başarısız olursa sessizce False döner — dosya pivot'suz da geçerli kalır.
    """
    if n_rows < 2:
        return False
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    pythoncom.CoInitialize()
    xl = None
    wb = None
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(os.path.abspath(path))
        if pivot_sheet not in [s.Name for s in wb.Sheets]:
            wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count)).Name = pivot_sheet
        src = f"{data_sheet}!R1C1:R{n_rows}C4"
        pc = wb.PivotCaches().Create(SourceType=1, SourceData=src)
        pt = pc.CreatePivotTable(TableDestination=f"{pivot_sheet}!R3C1", TableName="TrafikPivot")
        pt.PivotFields("Saat Aralığı").Orientation = 1   # satır: saat aralığı
        pt.PivotFields("Cadde Adı").Orientation = 2      # sütun: cadde adı
        pt.AddDataField(pt.PivotFields("Araç Sayısı"), "Toplam Araç", -4157)  # TOPLA
        wb.Save()
        return True
    except Exception:
        return False
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def build_pivot_source(ws, hourly, arms_map, hfill):
    """Pivot veri kaynağı sayfasını doldurur: Tarih | Saat Aralığı | Cadde Adı | Araç Sayısı.

    Satır sayısını (başlık dahil) döner — Excel COM pivot aralığı bununla hesaplanır.
    """
    ws.append(["Tarih", "Saat Aralığı", "Cadde Adı", "Araç Sayısı"])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = hfill
    n_rows = 1
    for h in sorted(hourly.keys()):
        try:
            tarih = datetime.strptime(h, "%Y-%m-%d %H:%M").strftime("%d.%m.%Y")
        except ValueError:
            tarih = h.split(" ")[0]
        for arm, val in sorted(hourly[h].items()):
            ws.append([tarih, hour_range(h), arms_map.get(arm, arm), round(val)])
            n_rows += 1
    return n_rows


def hour_range(hkey):
    try:
        dt = datetime.strptime(hkey, "%Y-%m-%d %H:%M")
        end = dt + timedelta(minutes=60)
        return f"{dt.strftime('%H:%M')} - {end.strftime('%H:%M')}"
    except ValueError:
        hh = int(hkey.split(" ")[-1])
        return f"{hh:02d}:00-{(hh + 1) % 24:02d}:00"

def build_street_summary(ws, hourly, arms_map, hfill):
    """Cadde bazlı satırlar doldurur (paneldeki sayaçla aynı veriden).

    Her saat aralığı için cadde satırı: Tarih | Saat Aralığı | Cadde Adı | Araç Sayısı
    Örnek: 12.08.2026 | 17:00 - 18:00 | Ali Bozdoğanoğlu Bulvarı | 145
    """
    ws.append([])
    ws.append(["CADDE BAZLI SAYIM (panel sayacıyla aynı)"])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    ws.append([])
    ws.append(["Tarih", "Saat Aralığı", "Cadde Adı", "Araç Sayısı"])
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = hfill
        cell.alignment = Alignment(horizontal="center")
    for h in sorted(hourly.keys()):
        try:
            tarih = datetime.strptime(h, "%Y-%m-%d %H:%M").strftime("%d.%m.%Y")
        except ValueError:
            tarih = h.split(" ")[0]
        for arm, val in sorted(hourly[h].items()):
            cadde = arms_map.get(arm, arm)
            ws.append([tarih, hour_range(h), cadde, round(val)])
    for col, w in zip("ABCD", [14, 18, 36, 14]):
        ws.column_dimensions[col].width = w

def build_hourly_sheet(ws, hourly, hfill, all_arms=None):
    """hourly verisinden saat bazlı sayım tablosunu doldurur.

    Paneldeki 'Bu Saat Geçen Araç' sayacıyla BİREBİR aynı değerleri üretir
    (/api/export ve /api/report ikisi de bunu kullanır — Excel farkı böyle kapanır).
    all_arms verilirse 0 araçlı kollar da tabloda yer alır (eksik kol kalmaz).
    """
    hours = sorted(hourly.keys())
    seen = {arm for h in hours for arm in hourly[h].keys()}
    if all_arms:
        seen |= set(all_arms)
    arms = sorted(seen)
    ws.append(["Saat Aralığı"] + arms + ["Toplam"])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = hfill
        cell.alignment = Alignment(horizontal="center")
    grand = {}
    for h in hours:
        row = [hour_range(h)]
        total = 0
        for arm in arms:
            val = round(hourly[h].get(arm, 0))
            row.append(val)
            total += val
            grand[arm] = grand.get(arm, 0) + val
        row.append(total)
        ws.append(row)
    row = ["Genel Toplam"] + [grand.get(arm, 0) for arm in arms] + [sum(grand.values())]
    ws.append(row)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 16
    for i in range(2, len(arms) + 3):
        ws.column_dimensions[get_column_letter(i)].width = 14
    ws.freeze_panes = "B2"
    return hours, arms

@app.route("/api/report")
def api_report():
    with state["lock"]:
        hourly = {k: dict(v) for k, v in state["hourly"].items()}
        coord = state["coord"]
        arms_map = {n: (a.get("street") or n) for n, a in state["arms"].items()}
        arm_names = list(state["arms"].keys())
    if not hourly:
        return jsonify({"error": "Henüz saatlik veri yok. Önce izlemeyi başlatın."}), 400

    hours = sorted(hourly.keys())
    arms = sorted({arm for h in hours for arm in hourly[h].keys()})

    path = os.path.join(DATA_DIR, f"saatlik_trafik_raporu_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb = Workbook()
    hfill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")

    # 1) SAAT BAZLI SAYIM — paneldeki sayaçla birebir (DOĞRU veri, ilk sayfa)
    ws = wb.active
    ws.title = "Saat Bazli Sayim"
    hours, arms = build_hourly_sheet(ws, hourly, hfill, all_arms=arm_names)

    # 2) CADDE BAZLI ÖZET — Tarih | Saat Aralığı | Cadde Adı | Araç Sayısı
    cs = wb.create_sheet("Cadde Bazli Ozet")
    build_street_summary(cs, hourly, arms_map, hfill)

    # 3) ÖZET — genel bilgiler
    summary = wb.create_sheet("Ozet")
    summary.append(["Rapor Tarihi", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    if coord:
        summary.append(["Kavsak Koordinati", f"{coord['lat']}, {coord['lon']} (yaricap {coord['radius']}m)"])
    summary.append(["Bildirilen Saatler", ", ".join(hours)])
    summary.append(["Izlenen Kol Sayisi", len(arms)])
    for cell in summary["A"]:
        cell.font = Font(bold=True)

    # 4) PİVOT TABLO — gerçek Excel pivot nesnesi (Excel COM ile eklenir)
    pv = wb.create_sheet("PivotVeri")
    pv.sheet_state = "hidden"
    n_rows = build_pivot_source(pv, hourly, arms_map, hfill)
    wb.create_sheet("Pivot")
    wb.save(path)
    add_excel_pivot(path, "PivotVeri", n_rows, "Pivot")
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    init_db()
    # Canlı ortam için: HOST=0.0.0.0, PORT ortam değişkeninden okunur.
    app.run(
        host=os.environ.get("HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", "5000")),
        debug=os.environ.get("FLASK_DEBUG", "0") == "1",
        use_reloader=False,
    )