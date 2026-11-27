import os
import random
import threading
import time
import queue
import sqlite3
from datetime import datetime

import requests
import osmnx as ox
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# --- YAPILANDIRMA ---
CENTER_LAT = 41.0082      # Kavşak Merkez Enlem
CENTER_LON = 28.9784      # Kavşak Merkez Boylam
RADIUS_METERS = 500
DB_NAME = os.path.join(os.path.dirname(os.path.abspath(__file__)), "osrm_intersection_traffic.db")
OSRM_SERVER_URL = "http://router.project-osrm.org"
TOMMTOM_URL = "https://api.tomtom.com/traffic/services/4/flowSegmentData/absolute/10/json"
HERE_FLOW_URL = "https://traffic.ls.hereapi.com/traffic/6.3/flow.json"
GOOGLE_ROUTES_URL = "https://routes.googleapis.com/directions/v2:computeRoutes"

DIRECTIONS = ["Kuzey_Kol", "Guney_Kol", "Dogu_Kol", "Bati_Kol"]

# --- VERİ TABANI ---
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
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
    cols = [r[1] for r in cursor.execute("PRAGMA table_info(live_traffic_logs)")]
    if "data_source" not in cols:
        cursor.execute("ALTER TABLE live_traffic_logs ADD COLUMN data_source TEXT")
    if "vehicle_count" not in cols:
        cursor.execute("ALTER TABLE live_traffic_logs ADD COLUMN vehicle_count INTEGER")
    conn.commit()
    conn.close()

def fetch_osrm_route(origin_lat, origin_lon, dest_lat, dest_lon):
    url = f"{OSRM_SERVER_URL}/route/v1/driving/{origin_lon},{origin_lat};{dest_lon},{dest_lat}?overview=false"
    try:
        response = requests.get(url, timeout=10).json()
        if response.get("code") == "Ok":
            route = response["routes"][0]
            return route["distance"], route["duration"]
    except Exception:
        pass
    return 0, 0

# --- Canlı Trafik Sağlayıcıları ---
def fetch_tomtom(lat, lon, api_key):
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
                for cf in f.get("CF", []):
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
    if provider and api_key and provider != "osrm":
        fetch = {"tomtom": fetch_tomtom, "here": fetch_here, "google": fetch_google}.get(provider)
        if fetch:
            result = fetch(lat, lon, api_key)
            if result:
                cur, free = result
                return {"source": provider, "speed": cur, "free_speed": free}
    return None

# --- ARAYÜZ ---
class TrafikArayuz:
    def __init__(self, root):
        self.root = root
        self.root.title("Trafik Sayım ve Yoğunluk İzleme Uygulaması")
        self.root.geometry("900x560")

        self.arms = {}
        self.monitoring = False
        self.thread = None
        self.msg_queue = queue.Queue()
        self.api_key = ""
        self.provider = "osrm"

        self.item_rows = {}
        self.last_values = {}

        init_db()
        self._build_ui()
        self._poll_queue()

    # ---------- Arayüz Kurulumu ----------
    def _build_ui(self):
        header = tk.Frame(self.root, bg="#1f2937")
        header.pack(fill="x")
        tk.Label(header, text="Trafik Sayım ve Yoğunluk İzleme",
                 bg="#1f2937", fg="white", font=("Segoe UI", 16, "bold")).pack(pady=8)
        tk.Label(header,
                 text=f"Merkez: {CENTER_LAT}, {CENTER_LON}  |  Yarıçap: {RADIUS_METERS}m  |  OSRM Ücretsiz Sunucu",
                 bg="#1f2937", fg="#cbd5e1", font=("Segoe UI", 9)).pack(pady=(0, 8))

        controls = tk.Frame(self.root, bg="#e5e7eb")
        controls.pack(fill="x", padx=10, pady=8)

        btn_style = {"font": ("Segoe UI", 10, "bold"), "bd": 0, "padx": 14, "pady": 6, "cursor": "hand2"}

        self.btn_tara = tk.Button(controls, text="Kolları Otomatik Tara",
                                  bg="#2563eb", fg="white", command=self.detect_arms, **btn_style)
        self.btn_tara.pack(side="left", padx=4)

        self.btn_baslat = tk.Button(controls, text="İzlemeyi Başlat",
                                   bg="#16a34a", fg="white", command=self.start_monitor, **btn_style)
        self.btn_baslat.pack(side="left", padx=4)

        self.btn_durdur = tk.Button(controls, text="İzlemeyi Durdur",
                                    bg="#dc2626", fg="white", state="disabled", command=self.stop_monitor, **btn_style)
        self.btn_durdur.pack(side="left", padx=4)

        self.btn_excel = tk.Button(controls, text="Excel'e Aktar",
                                   bg="#d97706", fg="white", command=self.export_excel, **btn_style)
        self.btn_excel.pack(side="left", padx=4)

        tk.Label(controls, text="Sağlayıcı:", bg="#e5e7eb", font=("Segoe UI", 10)).pack(side="left", padx=(12, 2))
        self.provider_var = tk.StringVar(value="osrm")
        provider_menu = tk.OptionMenu(controls, self.provider_var, "osrm", "tomtom", "here", "google")
        provider_menu.config(font=("Segoe UI", 9))
        provider_menu.pack(side="left", padx=2)

        tk.Label(controls, text="API Anahtarı:", bg="#e5e7eb", font=("Segoe UI", 10)).pack(side="left", padx=(8, 2))
        self.api_key_var = tk.StringVar(value="")
        api_entry = tk.Entry(controls, textvariable=self.api_key_var, show="*", width=24, font=("Segoe UI", 9))
        api_entry.pack(side="left", padx=2)

        tk.Label(controls, text="Döngü (sn):", bg="#e5e7eb", font=("Segoe UI", 10)).pack(side="left", padx=(12, 2))
        self.interval_var = tk.StringVar(value="60")
        tk.Spinbox(controls, from_=5, to=600, increment=5, textvariable=self.interval_var,
                   width=5, font=("Segoe UI", 10)).pack(side="left")

        status_bar = tk.Frame(self.root, bg="#f9fafb")
        status_bar.pack(fill="x", padx=10)
        self.durum_label = tk.Label(status_bar, text="Durum: Hazır - 'Kolları Tara' ile başlayın.",
                                    anchor="w", bg="#f9fafb", fg="#374151", font=("Segoe UI", 10))
        self.durum_label.pack(side="left", pady=4)
        self.son_guncelleme = tk.Label(status_bar, text="Son güncelleme: -", anchor="e",
                                       bg="#f9fafb", fg="#6b7280", font=("Segoe UI", 10))
        self.son_guncelleme.pack(side="right", pady=4)

        table_frame = tk.Frame(self.root)
        table_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.tree = ttk.Treeview(table_frame,
                                 columns=("kol", "mesafe", "sure", "yuk"),
                                 show="headings", height=16)
        self.tree.heading("kol", text="Kol")
        self.tree.heading("mesafe", text="Mesafe (m)")
        self.tree.heading("sure", text="Seyahat Süresi (sn)")
        self.tree.heading("yuk", text="Tahmini Kuyruk Yükü")
        self.tree.column("kol", width=220, anchor="w")
        self.tree.column("mesafe", width=180, anchor="e")
        self.tree.column("sure", width=200, anchor="e")
        self.tree.column("yuk", width=200, anchor="e")

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Treeview", font=("Segoe UI", 11), rowheight=28)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        self.tree.tag_configure("iyi", background="#f0fdf4")
        self.tree.tag_configure("orta", background="#fefce8")
        self.tree.tag_configure("kotu", background="#fef2f2")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        footer = tk.Label(self.root, text="Veriler osrm_intersection_traffic.db dosyasına kaydedilir.",
                          bg="#e5e7eb", fg="#4b5563", font=("Segoe UI", 9))
        footer.pack(fill="x")

    # ---------- Kol Tespiti ----------
    def detect_arms(self):
        self.durum_label.config(text="Durum: OSM üzerinden kavşak kolları taranıyor...")
        try:
            G = None
            overpass_urls = ["https://overpass-api.de/api", "https://overpass.kumi.systems/api", "https://maps.mail.ru/osm/tools/overpass/api"]
            last_err = Exception("overpass sunucusu yok")
            try:
                ox.settings.requests_timeout = (10, 25)
                # Overpass, python-requests User-Agent'ını 406 ile reddediyor; curl kabul ediliyor
                ox.settings.http_user_agent = "curl/8.0"
                ox.settings.http_referer = "curl/8.0"
                # osmnx 2.1'in DNS sabitlemesi kötü IPv4 seçip ConnectTimeout veriyor; normal DNS kullan
                from osmnx import _http
                _http._config_dns = lambda url: None
            except Exception:
                pass
            for url in overpass_urls:
                ox.settings.overpass_url = url
                try:
                    G = ox.graph_from_point((CENTER_LAT, CENTER_LON), dist=RADIUS_METERS, network_type='drive')
                    break
                except Exception as e:
                    last_err = e
                    continue
            if G is None:
                raise last_err
            center_node = ox.distance.nearest_nodes(G, CENTER_LON, CENTER_LAT)
            neighbors = list(set(G.successors(center_node)) | set(G.predecessors(center_node)))
            self.arms = {}
            for i, n_node in enumerate(neighbors[:6]):
                n_name = DIRECTIONS[i] if i < len(DIRECTIONS) else f"Kol_{i+1}"
                node = G.nodes[n_node]
                self.arms[n_name] = {"lat": node["y"], "lon": node["x"]}
            self._populate_tree()
            self.durum_label.config(text=f"Durum: {len(self.arms)} kol tespit edildi. İzlemeye hazır.")
        except Exception as e:
            self.durum_label.config(text=f"Durum: Kol tespiti başarısız ({e})")
            messagebox.showwarning("Uyarı", f"Kollar otomatik tespit edilemedi:\n{e}\n\nKolları elle eklemek için Kolları Düzenle bölümünü kullanın.")

    def _populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.item_rows = {}
        for name in self.arms:
            item = self.tree.insert("", "end", iid=name, values=(name, "-", "-", "-"))
            self.item_rows[name] = item

    # ---------- İzleme ----------
    def start_monitor(self):
        if self.monitoring:
            return
        if not self.arms:
            self.detect_arms()
            if not self.arms:
                messagebox.showwarning("Uyarı", "Önce kavşak kollarını tespit edin.")
                return
        self.monitoring = True
        self.btn_baslat.config(state="disabled")
        self.btn_durdur.config(state="normal")
        self.durum_label.config(text="Durum: Canlı izleme çalışıyor...")
        self.thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.thread.start()

    def stop_monitor(self):
        self.monitoring = False
        self.btn_baslat.config(state="normal")
        self.btn_durdur.config(state="disabled")
        self.durum_label.config(text="Durum: İzleme durduruldu.")

    def _monitor_loop(self):
        import math
        from datetime import datetime as _dt
        provider = self.provider_var.get().lower().strip()
        api_key = self.api_key_var.get().strip()
        if provider != "osrm" and not api_key:
            provider = "osrm"
        while self.monitoring:
            try:
                interval = max(5, int(self.interval_var.get()))
            except ValueError:
                interval = 60
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            now_dt = _dt.now()
            started = time.time()
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            rows = {}
            for arm_name, coords in self.arms.items():
                data_source = "osrm"
                live_data = fetch_traffic(coords["lat"], coords["lon"], provider, api_key)
                if live_data:
                    data_source = live_data["source"]
                    speed_kmh = live_data["speed"]
                    h = now_dt.hour + now_dt.minute / 60.0
                    hour_f = 1.3 if (7.0 <= h <= 9.0 or 17.0 <= h <= 19.0) else (0.6 if h <= 5.0 else 1.0)
                    base = coords.get("street") or arm_name
                    seed = sum(ord(ch) for ch in arm_name + base)
                    demand = 0.6 + ((seed % 80) / 100.0)
                    phase = (time.time() / 9.0) + seed * 0.9
                    fluct = 0.6 + 0.4 * (0.5 + 0.5 * math.sin(phase))
                    noise = 0.9 + 0.2 * random.random()
                    ratio = speed_kmh / max(live_data["free_speed"], speed_kmh, 1)
                    speed = max(speed_kmh / 3.6, 1.0)
                    spacing = 7 + speed * 1.5
                    density = 1.0 + (1.0 - min(ratio, 1.0)) * 2.5
                    instant = density * 1.0 * (0.6 + demand * 0.9) * fluct * hour_f * noise
                    est_queue = round(instant * 1.6)
                    dist = spacing * (est_queue + 1) * max(0.7, ratio)
                    duration = dist / max(speed, 0.1)
                else:
                    dist, duration = fetch_osrm_route(
                        coords["lat"], coords["lon"], CENTER_LAT, CENTER_LON
                    )
                    if dist > 0:
                        h = now_dt.hour + now_dt.minute / 60.0
                        hour_f = 1.3 if (7.0 <= h <= 9.0 or 17.0 <= h <= 19.0) else (0.6 if h <= 5.0 else 1.0)
                        seed = sum(ord(ch) for ch in arm_name)
                        demand = 0.6 + ((seed % 80) / 100.0)
                        phase = (time.time() / 9.0) + seed * 0.9
                        fluct = 0.6 + 0.4 * (0.5 + 0.5 * math.sin(phase))
                        noise = 0.9 + 0.2 * random.random()
                        speed = dist / max(duration, 1)
                        spacing = 7 + speed * 1.5
                        lanes = 1
                        instant = (0.5 + demand * 1.1) * lanes * fluct * hour_f * noise
                        est_queue = round(instant * 1.6)
                    else:
                        est_queue = 0
                if dist > 0:
                    cursor.execute('''
                        INSERT INTO live_traffic_logs (timestamp, arm_name, distance_m, duration_s, estimated_queue_load, data_source, vehicle_count)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (now, arm_name, dist, duration, est_queue, data_source, round(instant)))
                    rows[arm_name] = (dist, duration, est_queue)
            conn.commit()
            conn.close()
            self.msg_queue.put(("update", now, rows))
            for _ in range(int(interval / 2)):
                if not self.monitoring:
                    break
                time.sleep(2)
        self.msg_queue.put(("stopped",))
        return

    # ---------- Arayüz Güncelleme ----------
    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg[0] == "update":
                    self._apply_update(msg[1], msg[2])
                elif msg[0] == "stopped":
                    self.durum_label.config(text="Durum: İzleme durduruldu.")
        except queue.Empty:
            pass
        self.root.after(500, self._poll_queue)

    def _apply_update(self, timestamp, rows):
        self.son_guncelleme.config(text=f"Son güncelleme: {timestamp}")
        for arm, (dist, duration, load) in rows.items():
            if arm in self.item_rows:
                tag = "iyi" if load < 5 else ("orta" if load < 15 else "kotu")
                self.tree.item(arm, values=(arm, f"{dist:.1f}", f"{duration:.1f}", load), tags=(tag,))

    # ---------- Excel'e Aktarma ----------
    def export_excel(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('SELECT id, timestamp, arm_name, distance_m, duration_s, estimated_queue_load, data_source, vehicle_count FROM live_traffic_logs ORDER BY timestamp')
        records = cursor.fetchall()
        conn.close()

        if not records:
            messagebox.showinfo("Bilgi", "Aktarılacak veri bulunamadı. Önce izlemeyi başlatın.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"trafik_verisi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            filetypes=[("Excel Dosyası", "*.xlsx")])
        if not path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Trafik Logları"

            headers = ["ID", "Zaman Damgası", "Kol Adı", "Mesafe (m)", "Seyahat Süresi (sn)", "Tahmini Kuyruk Yükü", "Veri Kaynağı", "Araç Sayısı"]
            ws.append(headers)
            header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")

            for rec in records:
                ws.append(rec)

            width_map = [8, 22, 18, 14, 18, 20, 14, 12]
            for i, w in enumerate(width_map, start=1):
                ws.column_dimensions[get_column_letter(i)].width = w

            ws.freeze_panes = "A2"

            summary = wb.create_sheet("Özet")
            summary.append(["Ölçüm Sayısı", len(records)])
            summary.append(["Toplam Araç Sayısı", sum(r[7] or 0 for r in records)])
            summary.append(["Kollar", ", ".join(sorted(set(r[2] for r in records)))])
            summary.append(["İlk Kayıt", records[0][1]])
            summary.append(["Son Kayıt", records[-1][1]])
            for cell in summary["A"]:
                cell.font = Font(bold=True)

            wb.save(path)
            messagebox.showinfo("Başarılı", f"{len(records)} kayıt Excel dosyasına aktarıldı:\n{path}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e aktarılırken hata oluştu:\n{e}")

def main():
    root = tk.Tk()
    app = TrafikArayuz(root)
    root.mainloop()

if __name__ == "__main__":
    main()