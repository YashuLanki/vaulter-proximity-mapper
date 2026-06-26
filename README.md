# Vaulter Proximity Mapper

> Automated employer & business proximity research for Vaulter Real Estate Investments portfolio properties.

Instead of manually researching and measuring distances to nearby businesses for each due diligence report, this tool does it in seconds — and exports a GeoJSON you drag straight into [Felt](https://felt.com).

---

## What It Does

1. Takes any property name from the Vaulter Project Master
2. Geocodes the property to lat/lon
3. Searches **Google Places** (or **OpenStreetMap** as a free fallback) for businesses within your chosen radius
4. Computes **distance in miles** and **cardinal direction** from the property for every result
5. Exports a **GeoJSON** (drag into Felt) and a **CSV** (paste into DD reports)

**Categories searched automatically:**

| Icon | Category | Examples |
|------|----------|---------|
| 🛒 | Retail Anchor | Costco, HEB, Home Depot, Target, Walmart |
| 🏨 | Hospitality | TownePlace Suites, Holiday Inn, Best Western |
| 🏭 | Industrial / Logistics | Amazon DC, Goodyear, warehouse parks |
| 🏥 | Healthcare | Hospitals, clinics, pharmacies |
| 🏫 | School / Civic | School campuses, libraries, civic centers |
| 🍔 | Restaurant / QSR | Chick-fil-A, Whataburger, Starbucks, sit-down dining |

---

## Demo Output

```
============================================================
  Vaulter Proximity Mapper
  Property : Pacific & Pinson - Forney
  Radius   : 3.0 miles  (4,828 m)
  Source   : OpenStreetMap (Overpass)
============================================================

  Geocoding: Pacific Ave & Pinson Rd, Forney, TX 75126 ... ✓

  Searching: 🛒  Retail Anchor ...  8 found
  Searching: 🏨  Hospitality ...  4 found
  Searching: 🏭  Industrial / Logistics ...  6 found
  Searching: 🏥  Healthcare ...  3 found
  Searching: 🏫  School / Civic ...  7 found
  Searching: 🍔  Restaurant / QSR ...  14 found

  ──────────────────────────────────────────────────────────
  PROXIMITY SUMMARY — Pacific & Pinson - Forney
  ──────────────────────────────────────────────────────────

  🏥  HEALTHCARE  (3 places)
    Texas Health Forney Hospital       SSE — 1.45 mi
    ...

  🏫  SCHOOL / CIVIC  (7 places)
    Wilson Elementary                  NNE — 0.45 mi
    Keith Bell Opportunity Central     ENE — 0.60 mi
    ...

  EXPORTS
  GeoJSON → proximity_output/Pacific_and_Pinson_-_Forney_20260625_1430.geojson
  CSV     → proximity_output/Pacific_and_Pinson_-_Forney_20260625_1430.csv
```

---

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/YOUR_USERNAME/vaulter-proximity-mapper.git
cd vaulter-proximity-mapper
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. (Optional but recommended) Add a Google Places API key

Google Places returns richer data than OpenStreetMap. OSM is the automatic free fallback if no key is set.

Copy the example env file and fill in your key:

```bash
cp .env.example .env
# then edit .env and paste your key
```

To get a Google Places API key:
1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Enable the **Places API**
3. Create an API key under **Credentials**

---

## Usage

### Interactive mode (easiest)

```bash
python proximity_mapper.py
```

You'll see a numbered list of all portfolio properties. Pick one, enter a radius, done.

### Command-line mode

```bash
# Default 3-mile radius
python proximity_mapper.py --property "Pacific & Pinson - Forney"

# Custom radius
python proximity_mapper.py --property "Pacific & Pinson - Forney" --radius 5

# Partial name match works
python proximity_mapper.py --property "Forney"

# List all registered properties
python proximity_mapper.py --list-properties
```

---

## Loading Results into Felt

1. Open your Felt map at [felt.com](https://felt.com)
2. Click **Upload** in the top toolbar
3. Drag the `.geojson` file from `proximity_output/` onto the map
4. Each category is pre-colored for easy reading:

| Category | Color |
|----------|-------|
| 🛒 Retail Anchor | Red |
| 🏨 Hospitality | Purple |
| 🏭 Industrial / Logistics | Orange |
| 🏥 Healthcare | Green |
| 🏫 School / Civic | Blue |
| 🍔 Restaurant / QSR | Teal |
| ⭐ Subject Property | Pin |

---

## Output Files

All exports land in `proximity_output/` (git-ignored):

| File | Use |
|------|-----|
| `PropertyName_YYYYMMDD_HHMM.geojson` | Drag into Felt |
| `PropertyName_YYYYMMDD_HHMM.csv` | Paste distances into DD report |

The `distance_label` CSV column is pre-formatted (e.g. `SSE — 2.1 mi`) so it pastes directly into DD memo proximity tables.

---

## Adding Properties

Edit `PROPERTY_REGISTRY` at the top of `proximity_mapper.py`:

```python
PROPERTY_REGISTRY = {
    "My New Property": "123 Main St, City, ST 12345",
    # ...
}
```

Use the most specific address available. City + State works as a fallback.

---

## Project Structure

```
vaulter-proximity-mapper/
├── proximity_mapper.py     # Main script
├── requirements.txt        # Python dependencies
├── .env.example            # API key template
├── .gitignore              # Excludes outputs + secrets
└── README.md               # This file
```

---

## How It Works

```
Property Name
     │
     ▼
PROPERTY_REGISTRY  →  address string
     │
     ▼
Geocoder (Nominatim / Google)  →  (lat, lon)
     │
     ▼
For each of 6 categories:
  Google Places API  (if GOOGLE_PLACES_API_KEY is set)
  └── OR OSM Overpass  (free fallback, no key needed)
     │
     ▼
Distance + Direction computed for every result
     │
     ▼
De-duplicate, sort by distance
     │
     ├──▶  GeoJSON  →  Felt
     └──▶  CSV      →  DD Report
```

---

## Requirements

- Python 3.8+
- `geopy`
- `requests`

See `requirements.txt`.
