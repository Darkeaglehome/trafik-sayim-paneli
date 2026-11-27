# Traffic Counting & Congestion Dashboard

An **intersection-level live traffic counting and congestion monitoring** app. For each intersection arm you mark on the map, the app fetches real road distance/duration (OSRM) or live traffic speed (TomTom / HERE / Google), then derives passing vehicle counts, queue load and hourly totals. Results are stored in SQLite and can be downloaded as an Excel report (including a pivot table).

- Web panel: Flask + Leaflet map (`app.py`, `templates/index.html`)
- Desktop version: Tkinter UI (`trafik_arayuz.py`)
- Author: [Darkeaglehome](https://github.com/Darkeaglehome)

> Türkçe sürüm: [README.md](README.md)

---

## Contents

- [Features](#features)
- [How to Use](#how-to-use)
- [Installation](#installation)
- [Running](#running)
- [Traffic Providers](#traffic-providers)
- [Counting Model](#counting-model)
- [API Endpoints](#api-endpoints)
- [Database](#database)
- [Excel Output](#excel-output)
- [Deployment](#deployment)
- [Project Layout](#project-layout)
- [License](#license)

---

## Features

- **Pick arms on the map:** click the road to add an intersection arm; arms are named automatically by compass bearing (North, SouthEast, West_2, …).
- **Real road data:** OSRM Table API returns the N×N distance/duration matrix plus route geometry in one request.
- **Live traffic support:** plug in a TomTom, HERE or Google Routes API key for real-time speed; without a key the app runs on free OSRM data.
- **Vehicle counting:** speed-based fundamental-diagram (Greenshields) model produces passing vehicles, lane count, queue load and hourly totals.
- **Signalized intersection toggle:** with a traffic signal, a 45% green-time share is applied; otherwise vehicles pass continuously.
- **Hourly counter:** the panel's "vehicles this hour" and the Excel values come from the same source, so they match exactly.
- **Excel report:** hourly counts, street-based summary, general summary, raw traffic logs and a real Excel pivot table.
- **Resilient monitoring loop:** Clear/Stop bump a `generation` counter so stale background jobs abort and never write data back.

---

## How to Use

1. **Open the panel** → `http://127.0.0.1:5000`
2. **Find the intersection:** type latitude/longitude or search an address (e.g. *Sancaktepe Atatürk Caddesi*).
3. **Mark the arms:** **just click on the road / lane in the direction of that arm.** Each click adds a point and the street name under it is resolved automatically.
   - Click once on each of the 3–4 arms of the intersection.
   - Clicking the same spot twice does not create a duplicate.
   - Misclicked? Use **Clear** next to "Marked Points".
4. **"Add Arms"** → marked points become arms, named by their true bearing from the intersection center.
5. **Choose a provider:** leave `OSRM` for free usage, or pick TomTom/HERE/Google and paste an API key.
6. **"Start Monitoring"** → the road network is computed and the monitoring loop begins. Tables and map refresh every 5–60 seconds.
7. **"Stop"** pauses monitoring (marked arms are preserved), **"Clear"** resets everything.
8. **"Export Excel"** or **"Save Report"** downloads the `.xlsx` file.

> In short: **clicking the lane direction on the road is enough** — the app does the rest.

Desktop version: `python trafik_arayuz.py` (requires Tkinter + osmnx; same monitoring logic).

---

## Installation

Requires **Python 3.10+**

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

The desktop version additionally needs `osmnx`:

```bash
pip install osmnx
```

---

## Running

```bash
python app.py
```

Then open `http://127.0.0.1:5000`.

Environment variables:

| Variable | Default | Description |
|---|---|---|
| `PORT` | `5000` | Port to bind |
| `HOST` | `0.0.0.0` | Address to bind |
| `FLASK_DEBUG` | `0` | `1` enables debug mode |
| `DATA_DIR` | project folder | Where the SQLite database is stored |

```bash
PORT=8000 DATA_DIR=/var/data python app.py
```

---

## Traffic Providers

| Provider | Key | Data | Notes |
|---|---|---|---|
| OSRM | Not required | Distance, duration, geometry | Free public server: `router.project-osrm.org` |
| TomTom | Required | Current + free-flow speed | Flow Segment Data API |
| HERE | Required | Flow speed inside bbox | Traffic Flow API |
| Google | Required | Speed derived from route duration | Routes API (`TRAFFIC_AWARE`) |

If no key is given or a provider fails, the app falls back to OSRM. API keys are not persisted server-side; they are only used per request.

---

## Counting Model

Passing vehicles are derived from real speed (Greenshields fundamental diagram):

```
r        = current_speed / free_flow_speed      (0..1)
density  = k_jam * (1 - r),  k_jam = 160 veh/km/lane
flow     = density * current_speed * 3.6        (veh/hour/lane)
passing  = flow / 3600 * 10 s * lanes * green_share * hour_factor
```

- As the road slows, density rises; flow peaks and then drops near gridlock.
- Physical cap: no more vehicles can pass than could fit on the arm at once.
- `green_share`: `0.45` for signalized intersections, `1.0` otherwise.
- `hour_factor`: `1.3` at peak hours (07–09, 17–19), `0.6` at night.

---

## API Endpoints

| Method | Path | Description |
|---|---|---|
| `GET` | `/` | Web panel |
| `POST` | `/api/detect` | Detects arms from marked points (does not start monitoring) |
| `POST` | `/api/add_arm` | Adds a single arm |
| `POST` | `/api/start` | Starts monitoring (`coords`, `provider`, `api_key`, `signalized`) |
| `POST` | `/api/stop` | Stops monitoring (arms preserved) |
| `POST` | `/api/clear` | Resets all state |
| `POST` | `/api/signalize` | Changes signalized state while running |
| `GET` | `/api/status` | Live state, arms, counters, hourly data |
| `GET` | `/api/export` | Detailed Excel (logs + hourly + pivot) |
| `GET` | `/api/report` | Hourly summary Excel report |

Example:

```bash
curl -X POST http://127.0.0.1:5000/api/start \
  -H "Content-Type: application/json" \
  -d '{"coords": [[41.0083, 28.9780], [41.0080, 28.9788]], "provider": "osrm", "signalized": true}'
```

---

## Database

`osrm_intersection_traffic.db` — single SQLite table:

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

Rows accumulate in 60-second windows; `vehicle_count` is the total vehicles that actually passed in that window. The table is cleared when a new monitoring session starts.

---

## Excel Output

`/api/export` and `/api/report` produce these sheets:

1. **Hourly Count** — arms × hour ranges with a grand-total row (identical to the panel).
2. **Street Summary** — Date | Hour Range | Street Name | Vehicle Count.
3. **Summary** — record count, total vehicles, arms, first/last record.
4. **Traffic Logs** — raw 60-second records.
5. **Pivot** — a real Excel pivot table (rows: hour range, columns: street, values: total vehicles).

> The pivot sheet is added through Excel COM and therefore only appears on **Windows with Excel installed**. On other platforms the file is produced without the pivot, with all other sheets intact.

---

## Deployment

The app runs as a single process (background monitoring thread + SQLite), so it must be served with **one worker**.

### Render / PaaS

The repo ships `render.yaml` and `Procfile`:

- Build: `pip install -r requirements.txt`
- Start: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 8 --timeout 300`
- Persistent disk: `/var/data` → `DATA_DIR=/var/data`

Steps:

1. Create a repository under the `Darkeaglehome` account and push the code.
2. In the hosting panel choose **New Web Service** and connect the repo.
3. Set the start command and `DATA_DIR`, attach a persistent disk.
4. Open the generated `https://...` URL after the deploy finishes.

### Your own VPS

```bash
pip install -r requirements.txt
DATA_DIR=/var/data gunicorn app:app \
  --bind 0.0.0.0:8000 --workers 1 --threads 8 --timeout 300
```

Put Nginx in front as a reverse proxy. Because of the in-process state and SQLite, keep `--workers` at **1**; otherwise each worker holds its own separate state.

---

## Project Layout

```
.
├── app.py                     # Flask server + monitoring loop + Excel generation
├── trafik_arayuz.py           # Tkinter desktop version
├── templates/index.html       # Web panel with Leaflet map
├── osrm_intersection_traffic.db
├── requirements.txt
├── Procfile                   # PaaS start command
├── render.yaml                # Render blueprint
└── README.md
```

---

## License

MIT © [Darkeaglehome](https://github.com/Darkeaglehome)
