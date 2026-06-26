"""
proximity_mapper.py — Vaulter Property Proximity Intelligence Tool
==================================================================
Reads properties directly from the Vaulter Project Master (PDF, CSV, or Excel).
Reads search categories from data/config.json — no hardcoding anywhere.

Usage:
  python proximity_mapper.py
  python proximity_mapper.py --property "Pacific & Pinson - Forney"
  python proximity_mapper.py --property "Pacific & Pinson - Forney" --radius 5
  python proximity_mapper.py --list-properties

Place files in the data/ folder:
  data/Vaulter_Project_Master.pdf   (or .csv / .xlsx)
  data/config.json                  (search categories)

Environment variables:
  GOOGLE_PLACES_API_KEY   — Google Places + Geocoding API key

Dependencies:
  pip install -r requirements.txt
"""

import os
import sys
import csv
import json
import math
import time
import glob
import argparse
import textwrap
from datetime import datetime
from typing import Optional

try:
    import requests
    from geopy.distance import geodesic
except ImportError:
    print("Missing dependencies. Run: pip install -r requirements.txt")
    sys.exit(1)

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# PATHS
# ---------------------------------------------------------------------------
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
DATA_DIR    = os.path.join(SCRIPT_DIR, "data")
OUTPUT_DIR  = os.path.join(SCRIPT_DIR, "proximity_output")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(DATA_DIR,   exist_ok=True)

# ---------------------------------------------------------------------------
# STATE NORMALIZATION
# ---------------------------------------------------------------------------
STATE_ABBR = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT",
    "Delaware": "DE", "Florida": "FL", "Georgia": "GA", "Hawaii": "HI",
    "Idaho": "ID", "Illinois": "IL", "Indiana": "IN", "Iowa": "IA",
    "Kansas": "KS", "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME",
    "Maryland": "MD", "Massachusetts": "MA", "Michigan": "MI",
    "Minnesota": "MN", "Mississippi": "MS", "Missouri": "MO",
    "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
    "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM",
    "New York": "NY", "North Carolina": "NC", "North Dakota": "ND",
    "Ohio": "OH", "Oklahoma": "OK", "Oregon": "OR", "Pennsylvania": "PA",
    "Rhode Island": "RI", "South Carolina": "SC", "South Dakota": "SD",
    "Tennessee": "TN", "Texas": "TX", "Utah": "UT", "Vermont": "VT",
    "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
    "Wisconsin": "WI", "Wyoming": "WY",
}
VALID_ABBR = set(STATE_ABBR.values())

def normalize_state(raw: str) -> str:
    raw = raw.strip()
    if raw in VALID_ABBR:
        return raw
    # Full name
    if raw in STATE_ABBR:
        return STATE_ABBR[raw]
    # Truncated (e.g. "Californ", "Colorad", "New Mex")
    for full, abbr in STATE_ABBR.items():
        if full.startswith(raw) or raw.startswith(full[:5]):
            return abbr
    return ""

# ---------------------------------------------------------------------------
# PROJECT MASTER READER
# ---------------------------------------------------------------------------
SKIP_NAMES = {
    "Project Name", "Template", "Project Sponsor", "Dashboard Link",
    "Project Category", "State", "Submitter", "Submission Date",
    "Project Sponsor Approval", "Approved",
}

def _is_valid_property_name(name: str) -> bool:
    if not name or len(name) < 2:
        return False
    if name in SKIP_NAMES:
        return False
    if "@" in name:
        return False
    import re
    if re.match(r"\d{2}/\d{2}/\d{2}", name):
        return False
    # Filter plain "First Last" people names (no digits, no &, no -)
    if re.match(r"^[A-Z][a-z]+ [A-Z][a-z]+$", name) and "&" not in name and "-" not in name:
        return False
    return True

def _parse_pdf(path: str) -> list:
    try:
        import pdfplumber
    except ImportError:
        print("  ⚠ pdfplumber not installed. Run: pip install pdfplumber")
        return []

    properties = {}
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                if not row or len(row) < 2:
                    continue
                name = (row[0] or "").strip()
                if not _is_valid_property_name(name):
                    continue
                # State is in the last column (index 3 or wherever non-empty)
                state = ""
                for col_idx in [3, 2, 1]:
                    if col_idx < len(row):
                        val = (row[col_idx] or "").strip()
                        norm = normalize_state(val)
                        if norm:
                            state = norm
                            break
                if state and name not in properties:
                    properties[name] = state

    return [{"name": k, "state": v} for k, v in properties.items()]

def _parse_csv(path: str) -> list:
    """
    Parse CSV/Excel export from Smartsheet.
    Minimal filtering — the structured format means we trust Project Name + State columns directly.
    Only skips the Template row and rows with no state.
    """
    properties = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (row.get("Project Name") or "").strip()
            state_raw = (row.get("State") or "").strip()
            # Only skip truly invalid rows
            if not name or name == "Template" or "@" in name:
                continue
            state = normalize_state(state_raw)
            if name and state and name not in properties:
                properties[name] = state
    return [{"name": k, "state": v} for k, v in properties.items()]

def _parse_excel(path: str) -> list:
    try:
        import openpyxl
    except ImportError:
        print("  ⚠ openpyxl not installed. Run: pip install openpyxl")
        return []

    properties = {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    header = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            header = [str(c or "").strip() for c in row]
            continue
        if not row:
            continue
        row_dict = dict(zip(header, [str(c or "").strip() for c in row]))
        name = row_dict.get("Project Name", "").strip()
        state_raw = row_dict.get("State", "").strip()
        if not _is_valid_property_name(name):
            continue
        state = normalize_state(state_raw)
        if name and state and name not in properties:
            properties[name] = state
    wb.close()
    return [{"name": k, "state": v} for k, v in properties.items()]

def load_project_master() -> list:
    """
    Find and parse the Project Master from data/.
    Priority: CSV/Excel first (more structured), then PDF.
    Returns list of {"name": ..., "state": ...} dicts.
    """
    # Look for files in data/ dir
    candidates = (
        glob.glob(os.path.join(DATA_DIR, "*.csv")) +
        glob.glob(os.path.join(DATA_DIR, "*.xlsx")) +
        glob.glob(os.path.join(DATA_DIR, "*.xls")) +
        glob.glob(os.path.join(DATA_DIR, "*.pdf"))
    )

    # Filter to likely Project Master files
    pm_files = [f for f in candidates
                if any(kw in os.path.basename(f).lower()
                       for kw in ["project", "master", "vaulter", "portfolio"])]

    if not pm_files:
        # Fall back to any file in data/
        pm_files = [f for f in candidates
                    if not os.path.basename(f).startswith(".")]

    if not pm_files:
        print(f"\n  ✗ No Project Master file found in {DATA_DIR}/")
        print("    Place your Vaulter_Project_Master.pdf (or .csv/.xlsx) in the data/ folder.")
        sys.exit(1)

    # Prefer CSV > Excel > PDF
    def sort_key(f):
        ext = os.path.splitext(f)[1].lower()
        return {".csv": 0, ".xlsx": 1, ".xls": 2, ".pdf": 3}.get(ext, 9)

    pm_files.sort(key=sort_key)
    chosen = pm_files[0]
    ext = os.path.splitext(chosen)[1].lower()

    print(f"  Reading Project Master: {os.path.basename(chosen)}")

    if ext == ".pdf":
        return _parse_pdf(chosen)
    elif ext == ".csv":
        return _parse_csv(chosen)
    elif ext in (".xlsx", ".xls"):
        return _parse_excel(chosen)
    else:
        print(f"  ✗ Unsupported file type: {ext}")
        sys.exit(1)

# ---------------------------------------------------------------------------
# CONFIG READER
# ---------------------------------------------------------------------------
DEFAULT_CATEGORIES = [
    {
        "label": "Retail & Big Box",
        "icon": "🛒",
        "color": "#E74C3C",
        "google_types": ["supermarket", "department_store", "shopping_mall",
                         "home_goods_store", "hardware_store", "warehouse_store"],
    },
    {
        "label": "Hospitality",
        "icon": "🏨",
        "color": "#9B59B6",
        "google_types": ["lodging"],
    },
    {
        "label": "Industrial & Logistics",
        "icon": "🏭",
        "color": "#F39C12",
        "google_types": ["storage", "moving_company"],
    },
    {
        "label": "Healthcare",
        "icon": "🏥",
        "color": "#2ECC71",
        "google_types": ["hospital", "doctor", "pharmacy", "health"],
    },
    {
        "label": "School & University",
        "icon": "🎓",
        "color": "#3498DB",
        "google_types": ["school", "university"],
    },
    {
        "label": "Government & Civic",
        "icon": "🏛️",
        "color": "#5D6D7E",
        "google_types": ["city_hall", "local_government_office"],
    },
    {
        "label": "Restaurant & QSR",
        "icon": "🍔",
        "color": "#1ABC9C",
        "google_types": ["restaurant", "meal_takeaway", "cafe"],
    },
]

def load_categories() -> list:
    """Load categories from data/config.json, or fall back to defaults."""
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                data = json.load(f)
            cats = data.get("search_categories", [])
            if cats:
                return cats
        except Exception as e:
            print(f"  ⚠ Could not read config.json: {e} — using defaults")
    else:
        print(f"  ℹ No config.json found in data/ — using built-in defaults")
    return DEFAULT_CATEGORIES

# ---------------------------------------------------------------------------
# DIRECTION HELPERS
# ---------------------------------------------------------------------------
def bearing_to_cardinal(bearing: float) -> str:
    dirs = ["N","NNE","NE","ENE","E","ESE","SE","SSE",
            "S","SSW","SW","WSW","W","WNW","NW","NNW"]
    return dirs[round(bearing / 22.5) % 16]

def compute_bearing(lat1, lon1, lat2, lon2) -> float:
    lat1, lat2 = math.radians(lat1), math.radians(lat2)
    dlon = math.radians(lon2 - lon1)
    x = math.sin(dlon) * math.cos(lat2)
    y = math.cos(lat1)*math.sin(lat2) - math.sin(lat1)*math.cos(lat2)*math.cos(dlon)
    return (math.degrees(math.atan2(x, y)) + 360) % 360

def distance_and_direction(origin, dest):
    dist = round(geodesic(origin, dest).miles, 2)
    bearing = compute_bearing(origin[0], origin[1], dest[0], dest[1])
    return dist, bearing_to_cardinal(bearing)

# ---------------------------------------------------------------------------
# GEOCODER
# ---------------------------------------------------------------------------
def geocode_address(address: str) -> Optional[tuple]:
    """
    Geocode an address string to (lat, lon).
    Tries Google Geocoding API first, then Nominatim as fallback.
    """
    api_key = os.environ.get("GOOGLE_PLACES_API_KEY", "").strip()

    # Google Geocoding
    if api_key:
        try:
            resp = requests.get(
                "https://maps.googleapis.com/maps/api/geocode/json",
                params={"address": address, "key": api_key},
                timeout=10,
            )
            data = resp.json()
            if data.get("status") == "OK":
                loc = data["results"][0]["geometry"]["location"]
                return (loc["lat"], loc["lng"])
            else:
                print(f"\n  ⚠ Google Geocoding: {data.get('status')} for '{address}'")
        except Exception as e:
            print(f"\n  ⚠ Google Geocoding error: {e}")

    # Nominatim fallback
    try:
        from geopy.geocoders import Nominatim
        geo = Nominatim(user_agent="vaulter_proximity_mapper/2.0")
        loc = geo.geocode(address, timeout=10)
        if loc:
            return (loc.latitude, loc.longitude)
    except Exception as e:
        print(f"\n  ⚠ Nominatim error: {e}")

    return None

def build_geocode_query(name: str, state: str) -> str:
    """
    Build the best possible geocodable query from a property name + state.
    Strips lot sizes, phase numbers, and acreage from names.
    """
    import re
    # Remove trailing numbers like "10", "80", "1200", "50"
    clean = re.sub(r"\s+\d+$", "", name).strip()
    # Remove parentheticals like "(Interlink 8/10)", "(Wilson 155)", "(Triangle)"
    clean = re.sub(r"\s*\(.*?\)", "", clean).strip()
    # Replace " & " with " and " for better geocoding
    clean = clean.replace(" & ", " and ")
    # Remove common non-geographic suffixes
    for suffix in [" Ph 2, 3, 4", " Ph II", " Ph 2", " Ph 3", " Ph 4"]:
        clean = clean.replace(suffix, "")
    return f"{clean}, {state}"

# ---------------------------------------------------------------------------
# GOOGLE PLACES SEARCH
# ---------------------------------------------------------------------------
def search_google_places(lat, lon, radius_m, google_types, api_key):
    results = []
    url = "https://maps.googleapis.com/maps/api/place/nearbysearch/json"
    seen_ids = set()

    for ptype in google_types:
        params = {
            "location": f"{lat},{lon}",
            "radius":   radius_m,
            "type":     ptype,
            "key":      api_key,
        }
        try:
            resp = requests.get(url, params=params, timeout=10)
            data = resp.json()
            for r in data.get("results", []):
                pid = r.get("place_id", "")
                if pid not in seen_ids:
                    seen_ids.add(pid)
                    results.append(r)
            time.sleep(0.15)
        except Exception as e:
            print(f"\n  ⚠ Google Places error ({ptype}): {e}")

    return results

def parse_google_result(r, origin, category_label, icon):
    loc = r.get("geometry", {}).get("location", {})
    dest_lat, dest_lon = loc.get("lat"), loc.get("lng")
    if dest_lat is None:
        return None
    dist, direction = distance_and_direction(origin, (dest_lat, dest_lon))
    return {
        "name":           r.get("name", "Unknown"),
        "category":       category_label,
        "icon":           icon,
        "address":        r.get("vicinity", ""),
        "latitude":       dest_lat,
        "longitude":      dest_lon,
        "distance_miles": dist,
        "direction":      direction,
        "distance_label": f"{direction} — {dist} mi",
        "rating":         r.get("rating", ""),
        "source":         "Google Places",
        "notes":          ", ".join(r.get("types", [])),
    }

# ---------------------------------------------------------------------------
# OSM OVERPASS (fallback when no Google key)
# ---------------------------------------------------------------------------
OVERPASS_URL = "https://overpass-api.de/api/interpreter"

OSM_TAG_MAP = {
    "supermarket":              [("shop", "supermarket")],
    "department_store":         [("shop", "department_store")],
    "shopping_mall":            [("shop", "mall")],
    "home_goods_store":         [("shop", "furniture")],
    "hardware_store":           [("shop", "doityourself")],
    "warehouse_store":          [("shop", "warehouse")],
    "lodging":                  [("tourism", "hotel"), ("tourism", "motel")],
    "storage":                  [("building", "warehouse")],
    "hospital":                 [("amenity", "hospital")],
    "doctor":                   [("amenity", "clinic")],
    "pharmacy":                 [("amenity", "pharmacy")],
    "school":                   [("amenity", "school")],
    "university":               [("amenity", "university")],
    "city_hall":                [("amenity", "townhall")],
    "local_government_office":  [("office", "government")],
    "restaurant":               [("amenity", "restaurant")],
    "meal_takeaway":            [("amenity", "fast_food")],
    "cafe":                     [("amenity", "cafe")],
    "gas_station":              [("amenity", "fuel")],
    "bank":                     [("amenity", "bank")],
    "park":                     [("leisure", "park")],
    "stadium":                  [("leisure", "stadium")],
}

def search_osm(lat, lon, radius_m, google_types):
    osm_tags = []
    for gt in google_types:
        osm_tags.extend(OSM_TAG_MAP.get(gt, []))
    if not osm_tags:
        return []

    around = f"(around:{radius_m},{lat},{lon})"
    lines = ["[out:json][timeout:25];("]
    for k, v in osm_tags:
        lines.append(f'  node["{k}"="{v}"]{around};')
        lines.append(f'  way["{k}"="{v}"]{around};')
    lines += [");", "out center;"]
    query = "\n".join(lines)

    try:
        resp = requests.post(OVERPASS_URL, data={"data": query}, timeout=30)
        return resp.json().get("elements", [])
    except Exception as e:
        print(f"\n  ⚠ Overpass error: {e}")
        return []

def parse_osm_result(el, origin, category_label, icon):
    tags = el.get("tags", {})
    if el.get("type") == "node":
        dest_lat, dest_lon = el.get("lat"), el.get("lon")
    else:
        c = el.get("center", {})
        dest_lat, dest_lon = c.get("lat"), c.get("lon")
    if dest_lat is None:
        return None

    name = tags.get("name") or tags.get("brand") or tags.get("operator") or "Unnamed"
    dist, direction = distance_and_direction(origin, (dest_lat, dest_lon))
    addr = " ".join(filter(None, [
        tags.get("addr:housenumber",""),
        tags.get("addr:street",""),
        tags.get("addr:city",""),
        tags.get("addr:state",""),
    ])).strip()

    return {
        "name":           name,
        "category":       category_label,
        "icon":           icon,
        "address":        addr,
        "latitude":       dest_lat,
        "longitude":      dest_lon,
        "distance_miles": dist,
        "direction":      direction,
        "distance_label": f"{direction} — {dist} mi",
        "rating":         "",
        "source":         "OpenStreetMap",
        "notes":          "",
    }

# ---------------------------------------------------------------------------
# EXPORT
# ---------------------------------------------------------------------------
def build_geojson(records, property_name, property_coords):
    features = [{
        "type": "Feature",
        "geometry": {"type": "Point", "coordinates": [property_coords[1], property_coords[0]]},
        "properties": {
            "name":           f"⭐ {property_name} (Subject)",
            "category":       "Subject Property",
            "distance_miles": 0,
            "direction":      "—",
            "distance_label": "Subject Property",
            "source":         "Vaulter Project Master",
            "marker-color":   "#FFD700",
        },
    }]

    for r in records:
        features.append({
            "type": "Feature",
            "geometry": {"type": "Point", "coordinates": [r["longitude"], r["latitude"]]},
            "properties": {
                "name":           f"{r['icon']} {r['name']}",
                "category":       r["category"],
                "address":        r["address"],
                "distance_miles": r["distance_miles"],
                "direction":      r["direction"],
                "distance_label": r["distance_label"],
                "rating":         r["rating"],
                "source":         r["source"],
                "notes":          r["notes"],
                "marker-color":   r.get("color", "#888888"),
            },
        })

    return {"type": "FeatureCollection", "features": features}

def export_csv(records, property_name, property_coords, filepath):
    fieldnames = ["name","category","address","latitude","longitude",
                  "distance_miles","direction","distance_label","rating","source","notes"]
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerow({
            "name":           f"⭐ {property_name} (Subject)",
            "category":       "Subject Property",
            "address":        "",
            "latitude":       property_coords[0],
            "longitude":      property_coords[1],
            "distance_miles": 0,
            "direction":      "—",
            "distance_label": "Subject Property",
            "rating":         "",
            "source":         "Vaulter Project Master",
            "notes":          "",
        })
        for r in sorted(records, key=lambda x: x["distance_miles"]):
            writer.writerow(r)

# ---------------------------------------------------------------------------
# MAIN RUNNER
# ---------------------------------------------------------------------------
def run(property_name: str, properties: list, categories: list, radius_miles: float = 3.0):
    radius_m  = int(radius_miles * 1609.34)
    api_key   = os.environ.get("GOOGLE_PLACES_API_KEY", "").strip()
    use_google = bool(api_key)

    # Find the property in the list
    match = next((p for p in properties if p["name"] == property_name), None)
    if not match:
        print(f"\n  ✗ '{property_name}' not found in Project Master.")
        sys.exit(1)

    state = match["state"]
    query = build_geocode_query(property_name, state)

    print(f"\n{'='*62}")
    print(f"  Vaulter Proximity Mapper")
    print(f"  Property : {property_name}  ({state})")
    print(f"  Radius   : {radius_miles} miles  ({radius_m:,} m)")
    print(f"  Source   : {'Google Places' if use_google else 'OpenStreetMap (Overpass)'}")
    print(f"{'='*62}\n")

    # Geocode
    print(f"  Geocoding: {query} ...", end=" ", flush=True)
    coords = geocode_address(query)
    if not coords:
        print("FAILED")
        print("  ✗ Could not geocode. Check your Google API key or try a more specific address.")
        sys.exit(1)
    print(f"✓  ({coords[0]:.5f}, {coords[1]:.5f})")

    # Search each category
    all_records = []
    color_map   = {c["label"]: c.get("color", "#888888") for c in categories}

    for cat in categories:
        label  = cat["label"]
        icon   = cat.get("icon", "📍")
        gtypes = cat.get("google_types", [])

        print(f"\n  Searching: {icon}  {label} ...", end=" ", flush=True)

        if use_google:
            raw     = search_google_places(coords[0], coords[1], radius_m, gtypes, api_key)
            parsed  = [parse_google_result(r, coords, label, icon) for r in raw]
        else:
            raw     = search_osm(coords[0], coords[1], radius_m, gtypes)
            parsed  = [parse_osm_result(r, coords, label, icon) for r in raw]

        parsed = [p for p in parsed if p and p["distance_miles"] <= radius_miles]
        parsed.sort(key=lambda x: x["distance_miles"])

        # Attach color for GeoJSON
        for p in parsed:
            p["color"] = cat.get("color", "#888888")

        print(f"{len(parsed)} found")
        all_records.extend(parsed)

    # De-duplicate by (name, rounded lat/lon)
    seen, deduped = set(), []
    for r in all_records:
        key = (r["name"].lower().strip(), round(r["latitude"],4), round(r["longitude"],4))
        if key not in seen:
            seen.add(key)
            deduped.append(r)

    print(f"\n  Total unique results: {len(deduped)}")

    # Print summary
    print(f"\n{'─'*62}")
    print(f"  PROXIMITY SUMMARY — {property_name}")
    print(f"{'─'*62}")

    by_cat = {}
    for r in deduped:
        by_cat.setdefault(r["category"], []).append(r)

    for cat in categories:
        label = cat["label"]
        icon  = cat.get("icon", "📍")
        rows  = by_cat.get(label, [])
        if not rows:
            continue
        print(f"\n  {icon}  {label.upper()}  ({len(rows)} places)")
        for r in rows[:10]:
            print(f"    {r['name'][:40]:<40}  {r['distance_label']}")

    # Export
    slug      = property_name.replace(" ", "_").replace("/", "-").replace("&", "and")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    geojson_path = os.path.join(OUTPUT_DIR, f"{slug}_{timestamp}.geojson")
    csv_path     = os.path.join(OUTPUT_DIR, f"{slug}_{timestamp}.csv")

    with open(geojson_path, "w", encoding="utf-8") as f:
        json.dump(build_geojson(deduped, property_name, coords), f, indent=2)
    export_csv(deduped, property_name, coords, csv_path)

    print(f"\n{'='*62}")
    print(f"  EXPORTS")
    print(f"{'='*62}")
    print(f"  GeoJSON → {geojson_path}")
    print(f"  CSV     → {csv_path}")
    print(f"\n  HOW TO LOAD INTO FELT:")
    print(f"  1. Open your Felt map at felt.com")
    print(f"  2. Click Upload (top toolbar)")
    print(f"  3. Drag the .geojson file onto the map")
    print(f"\n  Category colors:")
    for cat in categories:
        print(f"    {cat.get('icon','📍')}  {cat['label']:<30}  {cat.get('color','')}")
    print(f"{'='*62}\n")

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Vaulter Proximity Mapper — powered by Project Master + config.json",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            Examples:
              python proximity_mapper.py
              python proximity_mapper.py --property "Pacific & Pinson - Forney"
              python proximity_mapper.py --property "Pacific & Pinson - Forney" --radius 5
              python proximity_mapper.py --list-properties
        """),
    )
    parser.add_argument("--property", "-p", type=str, default=None)
    parser.add_argument("--radius",   "-r", type=float, default=3.0)
    parser.add_argument("--list-properties", "-l", action="store_true")
    args = parser.parse_args()

    print("\nLoading Project Master ...", end=" ", flush=True)
    properties = load_project_master()
    print(f"✓  {len(properties)} properties found")

    categories = load_categories()
    print(f"  Categories loaded: {len(categories)}")

    if args.list_properties:
        print("\nVaulter Portfolio — all properties:\n")
        by_state = {}
        for p in properties:
            by_state.setdefault(p["state"], []).append(p["name"])
        for state in sorted(by_state):
            print(f"  {state}")
            for name in by_state[state]:
                print(f"    • {name}")
        print()
        return

    if args.property is None:
        # Interactive
        print("\n" + "─"*42)
        print("Available properties:\n")
        names = [p["name"] for p in properties]
        for i, name in enumerate(names, 1):
            state = next(p["state"] for p in properties if p["name"] == name)
            print(f"  {i:>2}. {name}  ({state})")
        print()
        choice = input("Enter property name or number: ").strip()
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(names):
                args.property = names[idx]
            else:
                print("Invalid number.")
                sys.exit(1)
        else:
            # Fuzzy match
            matches = [p["name"] for p in properties if choice.lower() in p["name"].lower()]
            if len(matches) == 1:
                args.property = matches[0]
            elif len(matches) > 1:
                print(f"Multiple matches: {matches}")
                sys.exit(1)
            else:
                args.property = choice

        radius_input = input(f"Search radius in miles [default: {args.radius}]: ").strip()
        if radius_input:
            try:
                args.radius = float(radius_input)
            except ValueError:
                pass

    run(args.property, properties, categories, args.radius)

if __name__ == "__main__":
    main()
