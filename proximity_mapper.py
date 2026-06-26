"""
proximity_mapper.py — Vaulter Property Proximity Intelligence Tool
==================================================================
Given any property name from the Vaulter Project Master, this tool:
  1. Geocodes the property address
  2. Queries Google Places API (or OSM Overpass as fallback) for nearby
     key employers and businesses across 6 categories
  3. Computes distance (miles) and cardinal direction from the property
  4. Exports a GeoJSON + CSV ready to drag into Felt

Usage:
  python proximity_mapper.py
  python proximity_mapper.py --property "Pacific & Pinson - Forney"
  python proximity_mapper.py --property "Pacific & Pinson - Forney" --radius 5
  python proximity_mapper.py --list-properties

Environment variables (optional — OSM is used as fallback if absent):
  GOOGLE_PLACES_API_KEY   — Google Places API key

Dependencies:
  pip install geopy requests
"""

import os
import sys
import csv
import json
import math
import time
import argparse
import textwrap
from datetime import datetime
from typing import Optional

try:
    import requests
    from geopy.geocoders import Nominatim
    from geopy.distance import geodesic
except ImportError:
    print("Missing dependencies. Run: pip install -r requirements.txt")
    sys.exit(1)

# Load .env file if present (so GOOGLE_PLACES_API_KEY can be set there)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # python-dotenv not installed — env vars must be set manually


# ---------------------------------------------------------------------------
# VAULTER PROJECT MASTER — property registry
# Add/remove entries here as the portfolio changes.
# Format: "Display Name": "Geocodable address string"
# Google Geocoding API resolves each address to lat/lon at runtime.
# ---------------------------------------------------------------------------
PROPERTY_REGISTRY = {
    # Texas
    "Pacific & Pinson - Forney":      "Pacific Ave & Pinson Rd, Forney, TX 75126",
    "Long Branch (Wilson 155)":        "Wilson Rd, Forney, TX 75126",
    "Horseshoe Bay Lots":              "Horseshoe Bay, TX 78657",
    "Triad":                           "Forney, TX 75126",

    # Arizona
    "Rita Ranch":                      "Rita Ranch, Tucson, AZ 85747",
    "Cabazon":                         "Cabazon, AZ 85328",
    "Eloy 310 (Interlink 8/10)":       "Eloy, AZ 85131",
    "Kirby Hughes & Luckett":          "Maricopa, AZ 85138",
    "Picacho Crossing Ph II":          "Picacho, AZ 85241",
    "Magic Ranch 10":                  "Florence, AZ 85132",
    "Magic Ranch 80":                  "Florence, AZ 85132",
    "Mesquite Trails":                 "Queen Creek, AZ 85142",
    "Heartland 53":                    "Maricopa, AZ 85138",
    "Lucky Hunt":                      "Maricopa, AZ 85138",
    "Marabella":                       "Goodyear, AZ 85395",
    "Mountain View Ranch":             "Queen Creek, AZ 85142",
    "Hidden Canyon":                   "Peoria, AZ 85383",

    # California
    "Banning":                         "Banning, CA 92220",
    "Affresco East":                   "Hesperia, CA 92345",
    "Affresco West":                   "Hesperia, CA 92345",
    "Apple Valley & Ohna":             "Apple Valley, CA 92307",
    "Auburn & Verbena":                "Hesperia, CA 92345",
    "Fuchsia & Dos Palmas":            "Desert Hot Springs, CA 92240",
    "Hook & Cobalt/S&C":               "Hesperia, CA 92345",
    "Hopland & Cordova":               "Hopland, CA 95449",
    "Kemper Campbell":                 "Victorville, CA 92395",
    "South 20E":                       "Twentynine Palms, CA 92277",
    "Antelope & Ellis":                "Lancaster, CA 93536",
    "Griffin Ranch":                   "La Quinta, CA 92253",
    "Wilson & Florida":                "Hesperia, CA 92345",
    "Rosamond & 40th St.":             "Rosamond, CA 93560",
    "Calhoun 29":                      "Hesperia, CA 92345",

    # New Mexico
    "Mesa Del Sol":                    "Mesa del Sol, Albuquerque, NM 87105",
    "Los Senderos":                    "Albuquerque, NM 87121",

    # Colorado
    "Mead (WCR 34 & Hwy 25)":         "Mead, CO 80542",
}


# ---------------------------------------------------------------------------
# SEARCH CATEGORIES
# Each entry: (display_label, google_place_type, osm_tags, felt_color)
# ---------------------------------------------------------------------------
CATEGORIES = [
    {
        "label":        "Retail Anchor",
        "google_types": ["supermarket", "department_store", "shopping_mall",
                          "home_goods_store", "hardware_store", "warehouse_store"],
        "osm_tags":     [
            ('shop', 'supermarket'), ('shop', 'department_store'),
            ('shop', 'mall'), ('shop', 'warehouse'), ('shop', 'doityourself'),
        ],
        "color":        "#E74C3C",   # red
        "icon":         "🛒",
    },
    {
        "label":        "Hospitality",
        "google_types": ["lodging"],
        "osm_tags":     [('tourism', 'hotel'), ('tourism', 'motel')],
        "color":        "#9B59B6",   # purple
        "icon":         "🏨",
    },
    {
        "label":        "Industrial / Logistics",
        "google_types": ["storage", "moving_company", "logistics"],
        "osm_tags":     [
            ('landuse', 'industrial'), ('building', 'warehouse'),
            ('amenity', 'logistics'),
        ],
        "color":        "#F39C12",   # orange
        "icon":         "🏭",
    },
    {
        "label":        "Healthcare",
        "google_types": ["hospital", "doctor", "pharmacy", "health"],
        "osm_tags":     [
            ('amenity', 'hospital'), ('amenity', 'clinic'),
            ('amenity', 'pharmacy'), ('healthcare', '*'),
        ],
        "color":        "#2ECC71",   # green
        "icon":         "🏥",
    },
    {
        "label":        "School / Civic",
        "google_types": ["school", "university", "library", "city_hall",
                          "local_government_office", "stadium"],
        "osm_tags":     [
            ('amenity', 'school'), ('amenity', 'college'),
            ('amenity', 'university'), ('amenity', 'library'),
            ('leisure', 'stadium'),
        ],
        "color":        "#3498DB",   # blue
        "icon":         "🏫",
    },
    {
        "label":        "Restaurant / QSR",
        "google_types": ["restaurant", "fast_food", "cafe", "meal_takeaway"],
        "osm_tags":     [
            ('amenity', 'restaurant'), ('amenity', 'fast_food'),
            ('amenity', 'cafe'),
        ],
        "color":        "#1ABC9C",   # teal
        "icon":         "🍔",
    },
]


# ---------------------------------------------------------------------------
# DIRECTION HELPERS
# ---------------------------------------------------------------------------
def bearing_to_cardinal(bearing: float) -> str:
    dirs = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
            "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
    ix = round(bearing / 22.5) % 16
    return dirs[ix]


def compute_bearing(origin_lat, origin_lon, dest_lat, dest_lon) -> float:
    lat1 = math.radians(origin_lat)
    lat2 = math.radians(dest_lat)
    dlon = math.radians(dest_lon - origin_lon)
    x = math.sin(dlon) * math.cos(lat2)
    y = (math.cos(lat1) * math.sin(lat2)
         - math.sin(lat1) * math.cos(lat2) * math.cos(dlon))
    bearing = math.degrees(math.atan2(x, y))
    return (bearing + 360) % 360


def distance_and_direction(origin, dest):
    """Return (miles, cardinal_direction) between two (lat, lon) tuples."""
    dist_miles = geodesic(origin, dest).miles
    bearing = compute_bearing(origin[0], origin[1], dest[0], dest[1])
    cardinal = bearing_to_cardinal(bearing)
    return round(dist_miles, 2), cardinal


# ---------------------------------------------------------------------------
# GEOCODER
# ---------------------------------------------------------------------------
def geocode_property(entry) -> Optional[tuple]:
    """
    Return (lat, lon) for a property address string.
    Tries Google Geocoding API first, then Nominatim as fallback.
    """
    address = entry

    # Try Google Geocoding API first (if key available)
    api_key = os.environ.get("GOOGLE_PLACES_API_KEY", "").strip()
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
                print(f"  ⚠ Google Geocoding: {data.get('status')} — falling back to Nominatim")
        except Exception as e:
            print(f"  ⚠ Google Geocoding error: {e} — falling back to Nominatim")

    # Nominatim fallback (OpenStreetMap, free)
    try:
        from geopy.geocoders import Nominatim
        geolocator = Nominatim(user_agent="vaulter_proximity_mapper/1.0")
        location = geolocator.geocode(address, timeout=10)
        if location:
            return (location.latitude, location.longitude)
        else:
            print(f"  ⚠ Nominatim: no results for: {address}")
    except Exception as e:
        print(f"  ⚠ Nominatim error: {e}")

    return None


# ---------------------------------------------------------------------------
# GOOGLE PLACES API
# ---------------------------------------------------------------------------
def search_google_places(lat, lon, radius_m, place_types, api_key):
    """Query Google Places Nearby Search for one or more place types."""
    results = []
    url = "https://maps.googleapis.com/maps/api/place/nearbysearch/json"

    for ptype in place_types:
        params = {
            "location": f"{lat},{lon}",
            "radius": radius_m,
            "type": ptype,
            "key": api_key,
        }
        try:
            resp = requests.get(url, params=params, timeout=10)
            data = resp.json()
            if data.get("status") not in ("OK", "ZERO_RESULTS"):
                print(f"  ⚠ Google Places status: {data.get('status')} for type={ptype}")
            results.extend(data.get("results", []))
            # Respect rate limits
            time.sleep(0.2)
        except Exception as e:
            print(f"  ⚠ Google Places error ({ptype}): {e}")

    # De-duplicate by place_id
    seen = set()
    unique = []
    for r in results:
        pid = r.get("place_id", "")
        if pid not in seen:
            seen.add(pid)
            unique.append(r)
    return unique


def parse_google_result(r, origin, category_label, icon):
    """Convert a Google Places result dict into a normalized record."""
    loc = r.get("geometry", {}).get("location", {})
    dest_lat = loc.get("lat")
    dest_lon = loc.get("lng")
    if dest_lat is None or dest_lon is None:
        return None

    dist, direction = distance_and_direction(origin, (dest_lat, dest_lon))

    return {
        "name":            r.get("name", "Unknown"),
        "category":        category_label,
        "icon":            icon,
        "address":         r.get("vicinity", ""),
        "latitude":        dest_lat,
        "longitude":       dest_lon,
        "distance_miles":  dist,
        "direction":       direction,
        "distance_label":  f"{direction} — {dist} mi",
        "rating":          r.get("rating", ""),
        "source":          "Google Places",
        "notes":           ", ".join(r.get("types", [])),
    }


# ---------------------------------------------------------------------------
# OSM OVERPASS (fallback)
# ---------------------------------------------------------------------------
OVERPASS_URL = "https://overpass-api.de/api/interpreter"


def build_overpass_query(lat, lon, radius_m, osm_tags):
    """Build an Overpass QL query for multiple tag pairs."""
    around = f"(around:{radius_m},{lat},{lon})"
    node_lines = []
    way_lines = []
    for k, v in osm_tags:
        val = v if v != "*" else ""
        tag = f'["{k}"="{v}"]' if val else f'["{k}"]'
        node_lines.append(f"  node{tag}{around};")
        way_lines.append(f"  way{tag}{around};")

    lines = ["[out:json][timeout:25];", "("]
    lines.extend(node_lines)
    lines.extend(way_lines)
    lines.append(");")
    lines.append("out center;")
    return "\n".join(lines)


def search_osm(lat, lon, radius_m, osm_tags):
    """Query Overpass API and return list of element dicts."""
    query = build_overpass_query(lat, lon, radius_m, osm_tags)
    try:
        resp = requests.post(OVERPASS_URL, data={"data": query}, timeout=30)
        data = resp.json()
        return data.get("elements", [])
    except Exception as e:
        print(f"  ⚠ Overpass error: {e}")
        return []


def parse_osm_result(el, origin, category_label, icon):
    """Convert an OSM element into a normalized record."""
    tags = el.get("tags", {})
    # Nodes have lat/lon directly; ways have a center
    if el.get("type") == "node":
        dest_lat = el.get("lat")
        dest_lon = el.get("lon")
    else:
        center = el.get("center", {})
        dest_lat = center.get("lat")
        dest_lon = center.get("lon")

    if dest_lat is None or dest_lon is None:
        return None

    name = (tags.get("name")
            or tags.get("brand")
            or tags.get("operator")
            or "Unnamed")

    dist, direction = distance_and_direction(origin, (dest_lat, dest_lon))

    # Build a readable address from OSM address tags
    addr_parts = [
        tags.get("addr:housenumber", ""),
        tags.get("addr:street", ""),
        tags.get("addr:city", ""),
        tags.get("addr:state", ""),
    ]
    address = " ".join(p for p in addr_parts if p).strip()

    return {
        "name":            name,
        "category":        category_label,
        "icon":            icon,
        "address":         address,
        "latitude":        dest_lat,
        "longitude":       dest_lon,
        "distance_miles":  dist,
        "direction":       direction,
        "distance_label":  f"{direction} — {dist} mi",
        "rating":          "",
        "source":          "OpenStreetMap",
        "notes":           "; ".join(f"{k}={v}" for k, v in tags.items()
                                     if k not in ("name", "brand", "operator")
                                     and not k.startswith("addr:")),
    }


# ---------------------------------------------------------------------------
# EXPORT
# ---------------------------------------------------------------------------
def build_geojson(records, property_name, property_coords):
    """Build a GeoJSON FeatureCollection from normalized records."""
    features = []

    # Add the subject property as its own feature
    features.append({
        "type": "Feature",
        "geometry": {
            "type": "Point",
            "coordinates": [property_coords[1], property_coords[0]],
        },
        "properties": {
            "name":           f"⭐ {property_name} (Subject)",
            "category":       "Subject Property",
            "icon":           "⭐",
            "distance_miles": 0,
            "direction":      "—",
            "distance_label": "Subject Property",
            "source":         "Vaulter Project Master",
        },
    })

    # Category → Felt-friendly color
    color_map = {c["label"]: c["color"] for c in CATEGORIES}

    for r in records:
        features.append({
            "type": "Feature",
            "geometry": {
                "type": "Point",
                "coordinates": [r["longitude"], r["latitude"]],
            },
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
                "marker-color":   color_map.get(r["category"], "#888888"),
            },
        })

    return {
        "type": "FeatureCollection",
        "features": features,
    }


def export_csv(records, property_name, property_coords, filepath):
    fieldnames = [
        "name", "category", "address",
        "latitude", "longitude",
        "distance_miles", "direction", "distance_label",
        "rating", "source", "notes",
    ]
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        # Write subject property first (with coords so Felt can plot it)
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
def run(property_name: str, radius_miles: float = 3.0):
    radius_m = int(radius_miles * 1609.34)
    api_key = os.environ.get("GOOGLE_PLACES_API_KEY", "").strip()
    use_google = bool(api_key)

    print(f"\n{'='*60}")
    print(f"  Vaulter Proximity Mapper")
    print(f"  Property : {property_name}")
    print(f"  Radius   : {radius_miles} miles  ({radius_m:,} m)")
    print(f"  Source   : {'Google Places' if use_google else 'OpenStreetMap (Overpass)'}")
    print(f"{'='*60}\n")

    # Step 1 — Resolve address
    address = PROPERTY_REGISTRY.get(property_name)
    if not address:
        # Fuzzy match
        matches = [k for k in PROPERTY_REGISTRY if property_name.lower() in k.lower()]
        if len(matches) == 1:
            property_name = matches[0]
            address = PROPERTY_REGISTRY[property_name]
            print(f"  ℹ Matched to: {property_name}")
        elif len(matches) > 1:
            print(f"  Ambiguous name. Did you mean one of:\n" +
                  "\n".join(f"    • {m}" for m in matches))
            sys.exit(1)
        else:
            print(f"  ✗ Property '{property_name}' not found in registry.")
            print("    Run with --list-properties to see all options.")
            sys.exit(1)

    # Step 2 — Geocode (or use hardcoded coords directly)
    entry = PROPERTY_REGISTRY.get(property_name) or address
    print(f"  Geocoding: {entry} ...", end=" ", flush=True)
    coords = geocode_property(entry)
    if not coords:
        print("FAILED")
        print("  ✗ Could not resolve coordinates. Add hardcoded (lat, lon) to PROPERTY_REGISTRY.")
        sys.exit(1)
    print(f"✓  ({coords[0]:.5f}, {coords[1]:.5f})")

    # Step 3 — Search each category
    all_records = []
    for cat in CATEGORIES:
        print(f"\n  Searching: {cat['icon']}  {cat['label']} ...", end=" ", flush=True)

        raw_results = []
        if use_google:
            raw_results = search_google_places(
                coords[0], coords[1], radius_m,
                cat["google_types"], api_key
            )
            parsed = [parse_google_result(r, coords, cat["label"], cat["icon"])
                      for r in raw_results]
        else:
            raw_results = search_osm(
                coords[0], coords[1], radius_m, cat["osm_tags"]
            )
            parsed = [parse_osm_result(r, coords, cat["label"], cat["icon"])
                      for r in raw_results]

        parsed = [p for p in parsed if p is not None]

        # Filter to radius (OSM "around" isn't always exact)
        parsed = [p for p in parsed if p["distance_miles"] <= radius_miles]

        # Sort by distance
        parsed.sort(key=lambda x: x["distance_miles"])

        print(f"{len(parsed)} found")
        all_records.extend(parsed)

    # Step 4 — De-duplicate across categories by (name, lat, lon rounded)
    seen_keys = set()
    deduped = []
    for r in all_records:
        key = (
            r["name"].lower().strip(),
            round(r["latitude"], 4),
            round(r["longitude"], 4),
        )
        if key not in seen_keys:
            seen_keys.add(key)
            deduped.append(r)

    print(f"\n  Total unique results: {len(deduped)}")

    # Step 5 — Export
    slug = property_name.replace(" ", "_").replace("/", "-").replace("&", "and")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = os.path.join(os.path.dirname(__file__), "proximity_output")
    os.makedirs(out_dir, exist_ok=True)

    geojson_path = os.path.join(out_dir, f"{slug}_{timestamp}.geojson")
    csv_path = os.path.join(out_dir, f"{slug}_{timestamp}.csv")

    geojson_data = build_geojson(deduped, property_name, coords)
    with open(geojson_path, "w", encoding="utf-8") as f:
        json.dump(geojson_data, f, indent=2)

    export_csv(deduped, property_name, coords, csv_path)

    # Step 6 — Print summary table
    print(f"\n{'─'*60}")
    print(f"  PROXIMITY SUMMARY — {property_name}")
    print(f"{'─'*60}")

    by_cat = {}
    for r in deduped:
        by_cat.setdefault(r["category"], []).append(r)

    for cat in CATEGORIES:
        cat_records = by_cat.get(cat["label"], [])
        if not cat_records:
            continue
        print(f"\n  {cat['icon']}  {cat['label'].upper()}  ({len(cat_records)} places)")
        for r in cat_records[:10]:   # show top 10 per category
            name = r["name"][:38].ljust(38)
            print(f"    {name}  {r['distance_label']}")

    print(f"\n{'='*60}")
    print(f"  EXPORTS")
    print(f"{'='*60}")
    print(f"  GeoJSON → {geojson_path}")
    print(f"  CSV     → {csv_path}")
    print(f"\n  HOW TO LOAD INTO FELT:")
    print(f"  1. Open your Felt map at felt.com")
    print(f"  2. Click  Upload  (top toolbar)")
    print(f"  3. Drag the .geojson file onto the map")
    print(f"  4. Each category has its own colour:")
    for cat in CATEGORIES:
        print(f"       {cat['icon']}  {cat['label']:<22}  {cat['color']}")
    print(f"{'='*60}\n")

    return geojson_path, csv_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Vaulter Proximity Mapper — find key employers & businesses near any portfolio property",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""
            Examples:
              python proximity_mapper.py
              python proximity_mapper.py --property "Pacific & Pinson - Forney"
              python proximity_mapper.py --property "Pacific & Pinson - Forney" --radius 5
              python proximity_mapper.py --list-properties

            Set GOOGLE_PLACES_API_KEY env var to use Google Places.
            OSM/Overpass is used automatically as fallback (no key required).
        """),
    )
    parser.add_argument(
        "--property", "-p",
        type=str,
        default=None,
        help="Property name (must match or partially match a name in the registry)",
    )
    parser.add_argument(
        "--radius", "-r",
        type=float,
        default=3.0,
        help="Search radius in miles (default: 3.0)",
    )
    parser.add_argument(
        "--list-properties", "-l",
        action="store_true",
        help="Print all properties in the registry and exit",
    )
    args = parser.parse_args()

    if args.list_properties:
        print("\nVaulter Project Master — registered properties:\n")
        by_state = {}
        for name, addr in PROPERTY_REGISTRY.items():
            state = addr.split(",")[-1].strip()[:2]
            by_state.setdefault(state, []).append((name, addr))
        for state in sorted(by_state):
            print(f"  {state}")
            for name, addr in by_state[state]:
                print(f"    • {name}")
                print(f"        {addr}")
        print()
        return

    if args.property is None:
        # Interactive mode
        print("\nVaulter Proximity Mapper")
        print("─" * 40)
        print("Available properties:\n")
        names = list(PROPERTY_REGISTRY.keys())
        for i, name in enumerate(names, 1):
            print(f"  {i:>2}. {name}")
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
            args.property = choice

        radius_input = input(f"Search radius in miles [default: {args.radius}]: ").strip()
        if radius_input:
            try:
                args.radius = float(radius_input)
            except ValueError:
                print("Invalid radius — using default.")

    run(args.property, args.radius)


if __name__ == "__main__":
    main()
