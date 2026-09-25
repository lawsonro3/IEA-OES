#!/usr/bin/env python3
"""Add capacity and position columns to gazetteers/projects.csv, for the maps.

Fills only empty cells, so values corrected by hand survive a re-run.

capacity_kw / capacity_basis
  1. Pilot     - the largest capacity the Pilot extracted from the reports (its records
                 carry the page and a quoted passage)
  2. PRIMRE    - the largest Capacity MW among the project's PRIMRE records
lat / lon / location_basis
  1. test site - the project's test site in gazetteers/test_sites.csv
  2. PRIMRE    - the Coordinates of its first PRIMRE record that has them
  3. place / region / country - the Pilot's OpenStreetMap geocoding of the location the
                 report gave (tables/geo_results.json), graded by what was matched:
                 a locality = place; a region or water body = region; the country alone =
                 country (a position that says nothing about where in the country)
"""

import csv
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE / "scripts"))
from seed_gazetteers import parse_coords  # noqa: E402

PROJECTS = HERE / "gazetteers" / "projects.csv"
SITES = HERE / "gazetteers" / "test_sites.csv"
PILOT = HERE.parent / "Extraction Pilot" / "tables"
PRIMRE = HERE.parent / "PRIMRE databases" / "primre-projects.csv"

UNIT_KW = {"W": 0.001, "kW": 1, "MW": 1_000, "GW": 1_000_000}
GEO_TIER = {"locality_region_country": "place", "locality_country": "place",
            "region_country": "region", "water_body_country": "region", "country": "country"}
NEW = ["capacity_kw", "capacity_basis", "lat", "lon", "location_basis"]


def split(s):
    return [x.strip() for x in (s or "").split(" | ") if x.strip()]


def main():
    rows = list(csv.DictReader(open(PROJECTS, encoding="utf-8")))
    cols = list(rows[0].keys())
    for c in NEW:
        if c not in cols:
            cols.insert(cols.index("test_site_name_raw"), c) if c.startswith("capacity") \
                else cols.insert(cols.index("location") + 1, c)
            for r in rows:
                r[c] = ""
    # keep the new columns together, in order
    for c in NEW:
        cols.remove(c)
    cols[cols.index("location") + 1:cols.index("location") + 1] = NEW

    sites = {s["site_id"]: s for s in csv.DictReader(open(SITES, encoding="utf-8"))}
    pilot = {r["project_id"]: r for r in csv.DictReader(open(PILOT / "projects.csv", encoding="utf-8"))}
    geo = json.load(open(PILOT / "geo_results.json", encoding="utf-8"))
    primre = {r["Name"]: r for r in csv.DictReader(open(PRIMRE, encoding="utf-8-sig"))}

    filled = {c: 0 for c in NEW}
    for r in rows:
        pids = [i for i in split(r["pilot_project_id"]) if i in pilot]
        precs = [primre[n] for n in split(r["primre_name"]) if n in primre]

        if not r["capacity_kw"]:
            caps = []
            for i in pids:
                v, u = pilot[i]["capacity_max_value"], pilot[i]["capacity_max_unit"]
                try:
                    caps.append(float(v) * UNIT_KW[u])
                except (ValueError, KeyError):
                    pass
            basis = "Pilot (report)"
            if not caps:
                for p in precs:
                    try:
                        caps.append(float(p["Capacity MW"]) * 1000)
                    except ValueError:
                        pass
                basis = "PRIMRE"
            caps = [c for c in caps if c > 0]
            if caps:
                r["capacity_kw"] = f"{max(caps):g}"
                r["capacity_basis"] = basis
                filled["capacity_kw"] += 1

        if not r["lat"]:
            lat = lon = basis = ""
            site = sites.get(r["test_site_id"])
            if site and site["lat"]:
                lat, lon, basis = site["lat"], site["lon"], "test site"
            if not lat:
                for p in precs:
                    lat, lon = parse_coords(p["Coordinates"])
                    if lat:
                        basis = "PRIMRE"
                        break
            if not lat:
                for i in pids:
                    g = geo.get(i) or {}
                    if g.get("geo_lat") is not None:
                        lat, lon = f"{g['geo_lat']:.4f}", f"{g['geo_lon']:.4f}"
                        basis = GEO_TIER.get(g.get("geo_tier"), "place")
                        break
            if lat:
                r["lat"], r["lon"], r["location_basis"] = lat, lon, basis
                filled["lat"] += 1

    with open(PROJECTS, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    print(f"capacity filled for {filled['capacity_kw']} projects, position for {filled['lat']}")


if __name__ == "__main__":
    main()
