# Proximity Mapper
**Automated employer & business proximity research for real estate portfolio properties.**

Given any property from your portfolio, this tool automatically finds nearby key employers, anchors, and businesses — computing the exact distance and direction from the site — and exports a GeoJSON you can drag straight into [Felt](https://felt.com) or a CSV for due diligence reports.

---

## What It Does

1. Reads your property portfolio from a CSV, Excel, or PDF file (no hardcoding)
2. Geocodes each property via Google Maps API
3. Searches **Google Places** (or **OpenStreetMap** as a free fallback) for businesses within your chosen radius
4. Computes **distance in miles** and **cardinal direction** from the property for every result
5. Exports a **GeoJSON** (drag into Felt) and a **CSV** (paste into reports)

**17 search categories — all configurable via `data/config.json`:**

| Icon | Category |
|------|----------|
| 🛒 | Big Box Retail (Costco, Walmart, Target, Home Depot) |
| 🏬 | Shopping Mall & Outlets |
| 🏨 | Hospitality (hotels, motels, extended stay) |
| 🏭 | Industrial & Logistics (warehouses, distribution centers) |
| 🏢 | Major Corporate HQ |
| 💻 | Technology & Innovation (data centers, tech parks) |
| 🏥 | Healthcare (hospitals, clinics, pharmacies) |
| 🎓 | School & University |
| 🏛️ | Government & Civic (city hall, courthouse, post office) |
| 🪖 | Military Base |
| 🏟️ | Sports & Entertainment (stadiums, arenas, casinos) |
| 🍔 | Restaurant & QSR |
| 🛍️ | Grocery & Specialty Food |
| ⛽ | Gas & Convenience (truck stops, travel centers) |
| 🏦 | Financial Services |
| 🌳 | Parks & Recreation |
| 🛣️ | Transportation & Infrastructure (airports, transit, intermodal) |

---

## Setup

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/proximity-mapper.git
cd proximity-mapper
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Add your data files to `data/`
```
data/
├── Your_Project_Master.csv     ← portfolio CSV/Excel/PDF export
└── config.json                 ← search categories (edit to customize)
```

The tool auto-detects any CSV, Excel, or PDF in `data/` that contains your property list. It expects columns: `Project Name` and `State`.

### 4. Add your Google Places API key (optional but recommended)
Copy the example env file and fill in your key:
```bash
cp .env.example .env
```
Edit `.env`:
```
GOOGLE_PLACES_API_KEY=your_key_here
```
OpenStreetMap is used automatically as a free fallback if no key is set.

To get a Google Places API key:
1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Enable **Places API** and **Geocoding API**
3. Create an API key under **Credentials**

---

## Usage

### Interactive mode
```bash
python proximity_mapper.py
```

### Command-line mode
```bash
# Default 3-mile radius
python proximity_mapper.py --property "My Property Name"

# Custom radius
python proximity_mapper.py --property "My Property Name" --radius 5

# List all properties loaded from your file
python proximity_mapper.py --list-properties
```

---

## Output

All exports land in `proximity_output/` (git-ignored):

| File | Use |
|------|-----|
| `PropertyName_YYYYMMDD_HHMM.geojson` | Drag into Felt |
| `PropertyName_YYYYMMDD_HHMM.csv` | Paste into due diligence report |

The `distance_label` CSV column is pre-formatted (e.g. `SSE — 2.1 mi`) for direct use in reports.

### Loading into Felt
1. Open your map at [felt.com](https://felt.com)
2. Click **Upload** in the top toolbar
3. Drag the `.geojson` file onto the map — each category has its own color

---

## Customizing Categories

Edit `data/config.json` to add, remove, or change search categories — no code changes needed:

```json
{
  "search_categories": [
    {
      "label": "My Custom Category",
      "icon": "🏗️",
      "color": "#E74C3C",
      "google_types": ["point_of_interest"]
    }
  ]
}
```

Google Place types reference: [developers.google.com/maps/documentation/places/web-service/supported_types](https://developers.google.com/maps/documentation/places/web-service/supported_types)

---

## Project Structure

```
proximity-mapper/
├── proximity_mapper.py     ← main script
├── requirements.txt        ← pip install -r requirements.txt
├── .env.example            ← API key template (copy to .env)
├── .gitignore              ← keeps data/ and outputs/ out of git
├── README.md
├── data/                   ← git-ignored, add your files here
│   ├── Your_Portfolio.csv
│   └── config.json
└── proximity_output/       ← git-ignored, generated exports land here
```

---

## How It Works

```
Portfolio File (CSV/Excel/PDF)
         │
         ▼
   Parse properties → name + state
         │
         ▼
   Google Geocoding API  →  (lat, lon)
         │
         ▼
   For each of 17 categories:
     Google Places API  (if key set)
     └── OR OSM Overpass  (free fallback)
         │
         ▼
   Distance + Direction computed per result
         │
         ▼
   De-duplicate, sort by distance
         │
         ├──▶  GeoJSON  →  Felt
         └──▶  CSV      →  Report
```

---

## Requirements

- Python 3.8+
- See `requirements.txt`

---

## License

MIT
