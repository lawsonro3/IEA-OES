#!/usr/bin/env python3
"""Seed the test-site and project gazetteers from the three existing sources.

Run once to create gazetteers/test_sites.csv and gazetteers/projects.csv. After that
the gazetteers are hand-edited files and this script should not be re-run over them
(it refuses to overwrite unless --force is given).

Sources (all read in place, never modified):
  - Annual Report Extraction/test_sites_gazetteer.csv  70 curated test sites (the base list)
  - Extraction Pilot/tables/projects.csv               474 LLM-extracted records, 2016/2020/2025
  - Extraction Pilot/tables/test_sites_tables.csv      site names from the reports' own tables
  - Extraction Pilot/validation/*.csv                  Pilot <-> PRIMRE matches and review verdicts
  - PRIMRE databases/primre-test-sites.csv             48 test sites (US DOE / PNNL)
  - PRIMRE databases/primre-projects.csv               326 projects

Test sites:
  The 70 curated sites are kept as they are. PRIMRE sites that match one of them
  (by name, alias or acronym, same country) are recorded against it; the rest are
  added as new rows marked for review. Pilot site names that match nothing are NOT
  added -- many are descriptive phrases that would make poor search aliases -- and go
  to gazetteers/test_site_candidates.csv instead.

Projects:
  Every Pilot project record becomes a row. Its PRIMRE match is taken from the Pilot's
  PRIMRE validation, honouring the verdicts already given in review_queue.csv. PRIMRE
  projects with no Pilot match are added as rows of their own. Nothing is merged: two
  Pilot rows that share a PRIMRE record stay two rows, flagged in `notes`.
"""

import argparse
import csv
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
ROOT = HERE.parent
ARE = ROOT / "Annual Report Extraction"
PILOT = ROOT / "Extraction Pilot"
PRIMRE = ROOT / "PRIMRE databases"
OUT = HERE / "gazetteers"

# Country names as the annual reports / Annual Report Extraction use them.
COUNTRY_FIX = {
    "south korea": "Republic of Korea",
    "korea": "Republic of Korea",
    "united states of america": "United States",
    "usa": "United States",
}

# Pilot technology codes and PRIMRE resources -> one vocabulary.
PILOT_TECH = {
    "wave": "wave", "tidal_stream": "tidal stream", "tidal_range": "tidal range",
    "river_current": "river current", "ocean_current": "ocean current", "OTEC": "OTEC",
    "salinity_gradient": "salinity gradient", "multi": "multiple", "hybrid": "hybrid",
    "other": "other", "unknown": "",
}
PRIMRE_TECH = {
    "Wave": "wave", "Tidal": "tidal", "Riverine": "river current", "Ocean Current": "ocean current",
    "Thermal Gradient": "OTEC", "Salinity Gradient": "salinity gradient",
}

# Review-queue verdicts that confirm a Pilot <-> PRIMRE pair is the same thing.
SAME_VERDICTS = {"match", "fix coordinate (same site)", "reconcile capacity"}

GENERIC = {"the", "test", "site", "sites", "centre", "center", "facility", "facilities",
           "national", "marine", "energy", "ocean", "renewable", "and", "of", "for", "at", "in"}


def read(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def write(path, rows, columns):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def country(name):
    name = " ".join((name or "").split())
    if name.lower() in COUNTRY_FIX:
        return COUNTRY_FIX[name.lower()]
    # "Republic Of Korea" -> "Republic of Korea"
    return " ".join(w if i == 0 or w.lower() not in ("of", "and", "the") else w.lower()
                    for i, w in enumerate(name.split()))


def norm(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()
    s = re.sub(r"['’]", "", s)          # Jennette's -> jennettes
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return " ".join(s.split())


def core(s):
    """Normalised name without generic words, for fuzzy comparison."""
    return " ".join(w for w in norm(s).split() if w not in GENERIC)


def acronym(name):
    m = re.search(r"\(([A-Za-z][A-Za-z0-9\-]{1,11})\)", name or "")
    return m.group(1) if m else ""


def aliases_for(name):
    """Search aliases for a new name: the name, the name without its bracketed
    acronym, and the acronym itself. Case-sensitive in extract.py, so keep case."""
    out = [name.strip()]
    stripped = re.sub(r"\s*\([^)]*\)", "", name).strip()
    if stripped and stripped != out[0]:
        out.append(stripped)
    if stripped.startswith("The "):             # reports often drop the article
        out.append(stripped[4:])
    a = acronym(name)
    if a and len(a) >= 3:
        out.append(a)
    return list(dict.fromkeys(x for x in out if x))


def slug(name, taken):
    base = acronym(name) or re.sub(r"\s*\([^)]*\)", "", name)
    base = re.sub(r"[^a-z0-9]+", "_", norm(base))[:40].strip("_") or "site"
    s, i = base, 2
    while s in taken:
        s, i = f"{base}_{i}", i + 1
    taken.add(s)
    return s


def parse_coords(text):
    """'29° 5' 55.86", -13° 41' 4.99"' or '29.1, -13.7' -> (lat, lon) strings, or ('', '')."""
    if not text or not text.strip():
        return "", ""
    parts = [p.strip() for p in text.split(",")]
    if len(parts) != 2:
        return "", ""
    vals = []
    for p in parts:
        m = re.match(r"^(-?)\s*(\d+(?:\.\d+)?)\s*°?\s*(?:(\d+(?:\.\d+)?)\s*'\s*)?(?:(\d+(?:\.\d+)?)\s*\"\s*)?$", p)
        if not m:
            return "", ""
        sign = -1 if m.group(1) else 1
        v = float(m.group(2)) + float(m.group(3) or 0) / 60 + float(m.group(4) or 0) / 3600
        vals.append(f"{sign * v:.4f}")
    return vals[0], vals[1]


# --------------------------------------------------------------------------------------
# Test sites
# --------------------------------------------------------------------------------------

class SiteIndex:
    """Matches free-text site names against the gazetteer rows."""

    def __init__(self, sites):
        self.sites = sites
        self.exact = defaultdict(list)   # normalised alias/name -> sites
        for s in sites:
            for a in [s["name"]] + s["aliases"].split("|"):
                a = a.strip()
                if a:
                    self.exact[norm(a)].append(s)
                    if core(a):
                        self.exact["core:" + core(a)].append(s)

    def match(self, name, ctry=None, threshold=0.88):
        ctry = country(ctry) if ctry else None
        ok = lambda s: not ctry or not s["country"] or s["country"] == ctry
        keys = [norm(name), "core:" + core(name)]
        a = acronym(name)
        if a:
            keys.append(norm(a))
        for k in keys:
            hits = [s for s in self.exact.get(k, []) if ok(s)]
            if hits:
                return hits[0]
        # Any alias of 3+ chars appearing as a whole word inside the name
        # ("DanWEC test site, Hanstholm" -> DanWEC).
        text = " " + norm(name) + " "
        best = None
        for k, ss in self.exact.items():
            if k.startswith("core:") or len(k) < 4:
                continue
            if f" {k} " in text:
                for s in ss:
                    if ok(s) and (best is None or len(k) > best[0]):
                        best = (len(k), s)
        if best:
            return best[1]
        c = core(name)
        if len(c) >= 5:
            scored = [(SequenceMatcher(None, c, core(s["name"])).ratio(), s)
                      for s in self.sites if ok(s) and core(s["name"])]
            if scored:
                r, s = max(scored, key=lambda x: x[0])
                if r >= threshold and same_places(c, core(s["name"])):
                    return s
        # Every distinctive word of a site's name appears in the text, same country
        # ("National Marine Comprehensive Test Site Zhoushan" -> "National Marine
        # Test Site (Zhoushan)"). Needs an explicit country to be safe.
        if ctry:
            words = set(norm(name).split())
            hits = []
            for s in self.sites:
                sc = set(core(re.sub(r"\([^)]*\)", "", s["name"])).split()) or set(core(s["name"]).split())
                if ok(s) and s["country"] and any(len(w) >= 5 for w in sc) and sc <= words:
                    hits.append((len(sc), s))
            if hits:
                return max(hits, key=lambda x: x[0])[1]
        return None


def same_places(a, b):
    """False when either name has a long word with no near-equal in the other: a
    different place name (Weihai vs Zhuhai), not a spelling variant (Brehat/Brahat)."""
    wa, wb = a.split(), b.split()
    for xs, ys in ((wa, wb), (wb, wa)):
        for x in xs:
            if len(x) >= 5 and not any(SequenceMatcher(None, x, y).ratio() >= 0.8 for y in ys):
                return False
    return True


# Known errors in the PRIMRE records, corrected on the way in.
PRIMRE_SITE_FIXES = {
    "Queens University Belfast - Strangford Lough": {"Country": "United Kingdom"},  # Northern Ireland
}


def seed_test_sites():
    base = read(ARE / "test_sites_gazetteer.csv")
    sites = []
    for r in base:
        sites.append({**r, "country": country(r["country"]), "source": "Annual Report Extraction",
                      "primre_name": "", "primre_status": "", "primre_resource": "",
                      "primre_website": "", "pilot_names": "", "review_status": "curated"})
    index = SiteIndex(sites)
    taken = {s["site_id"] for s in sites}

    # PRIMRE test sites: enrich a matching site or add a new one.
    added = enriched = 0
    for p in read(PRIMRE / "primre-test-sites.csv"):
        p = {**p, **PRIMRE_SITE_FIXES.get(p["Name"], {})}
        hit = index.match(p["Name"], p["Country"])
        info = {"primre_name": p["Name"], "primre_status": p["Status"],
                "primre_resource": p["Resource Type"], "primre_website": p["Website"]}
        if hit:
            if hit["primre_name"]:      # a second PRIMRE record for the same site
                for k, v in info.items():
                    hit[k] = hit[k] + " | " + v if v else hit[k]
            else:
                hit.update(info)
            hit["source"] += "; PRIMRE"
            enriched += 1
            continue
        lat, lon = parse_coords(p["Coordinates"])
        # Drop an acronym that already occurs inside another site's name or alias:
        # the US "AMEC" would otherwise match every mention of NAGASAKI-AMEC.
        existing = " ".join(" " + norm(x["name"] + " " + x["aliases"]) + " " for x in sites)
        aliases = [x for x in aliases_for(p["Name"])
                   if x != acronym(p["Name"]) or f" {norm(x)} " not in existing]
        row = {"site_id": slug(p["Name"], taken), "name": p["Name"], "country": country(p["Country"]),
               "location": p["Waterbody"], "lat": lat, "lon": lon,
               "coord_note": "from PRIMRE" if lat else "", "aliases": " | ".join(aliases),
               "source": "PRIMRE", **info, "pilot_names": "",
               "review_status": "added from PRIMRE; check aliases"}
        sites.append(row)
        index = SiteIndex(sites)
        added += 1

    # Pilot site names: record against a match, otherwise a candidate for review.
    pilot_names = Counter()
    where = defaultdict(set)
    for r in read(PILOT / "tables" / "projects.csv"):
        if r["record_type"] == "test_site":
            pilot_names[(r["canonical_name"], country(r["deployment_country"] or r["reporting_country"]))] += 1
            where[(r["canonical_name"], country(r["deployment_country"] or r["reporting_country"]))].add("Pilot projects (test_site)")
    for r in read(PILOT / "tables" / "test_sites_tables.csv"):
        key = (r["site_name"].strip(), country(r["country"]))
        pilot_names[key] += 1
        where[key].add("Pilot test-site tables")

    candidates = []
    for (name, ctry), n in sorted(pilot_names.items()):
        hit = index.match(name, ctry)
        if hit:
            names = [x for x in hit["pilot_names"].split(" | ") if x]
            if name not in names and name != hit["name"]:
                hit["pilot_names"] = " | ".join(names + [name])
            if "Pilot" not in hit["source"]:
                hit["source"] += "; Pilot"
        else:
            candidates.append({"candidate": name, "country": ctry, "n_records": n,
                               "found_in": "; ".join(sorted(where[(name, ctry)])),
                               "decision": "", "site_id": "", "notes": ""})

    cols = ["site_id", "name", "country", "location", "lat", "lon", "coord_note", "aliases",
            "source", "primre_name", "primre_status", "primre_resource", "primre_website",
            "pilot_names", "review_status"]
    write(OUT / "test_sites.csv", sites, cols)
    write(OUT / "test_site_candidates.csv", candidates,
          ["candidate", "country", "n_records", "found_in", "decision", "site_id", "notes"])
    print(f"test_sites.csv: {len(sites)} sites ({len(base)} curated, {enriched} PRIMRE matches "
          f"recorded, {added} added from PRIMRE)")
    print(f"test_site_candidates.csv: {len(candidates)} unmatched Pilot site names for review")
    return sites


# --------------------------------------------------------------------------------------
# Projects
# --------------------------------------------------------------------------------------

def seed_projects(sites):
    index = SiteIndex(sites)
    pilot = [r for r in read(PILOT / "tables" / "projects.csv") if r["record_type"] == "project"]
    primre = {p["Name"]: p for p in read(PRIMRE / "primre-projects.csv")}

    # Pilot -> PRIMRE links. The review queue overrides the automatic confidence.
    verdicts = {(r["iea_name"], r["primre_name"]): r["verdict"].strip()
                for r in read(PILOT / "validation" / "review_queue.csv") if r["dataset"] == "project"}
    link = {}
    for r in read(PILOT / "validation" / "primre_validation_projects.csv"):
        if not r["primre_name"]:
            continue
        pair = (r["iea_name"], r["primre_name"])
        v = verdicts.get(pair)
        if v is None:
            if r["confidence"] in ("high", "medium"):
                link[r["iea_name"]] = (r["primre_name"], f"auto ({r['confidence']} confidence)")
        elif v in SAME_VERDICTS:
            link[r["iea_name"]] = (r["primre_name"], f"reviewed: {v}")
        elif v.startswith("unsure"):
            link[r["iea_name"]] = (r["primre_name"], "UNSURE - still to review")

    rows, used_primre = [], Counter()
    for i, r in enumerate(pilot, 1):
        site = index.match(r["test_site_name"], r["deployment_country"]) if r["test_site_name"] else None
        pname, basis = link.get(r["canonical_name"], ("", ""))
        if pname and not basis.startswith("UNSURE"):
            used_primre[pname] += 1
        p = primre.get(pname, {})
        rows.append({
            "project_id": f"P{i:04d}",
            "name": r["canonical_name"],
            "aliases": r["aliases"].replace("; ", " | "),
            "developer": r["developer"],
            "device": "",
            "country": country(r["deployment_country"] or r["reporting_country"]),
            "reporting_country": country(r["reporting_country"]),
            "technology": PILOT_TECH.get(r["technology"], r["technology"]),
            "in_scope": "no" if r["technology"] == "other" else "yes",
            "test_site_id": site["site_id"] if site else "",
            "test_site_name_raw": r["test_site_name"],
            "location": r["location_raw_best"],
            "first_seen": r["first_seen"], "last_seen": r["last_seen"],
            "latest_status": r["latest_status"],
            "source": "Pilot" + ("; PRIMRE" if pname and not basis.startswith("UNSURE") else ""),
            "pilot_project_id": r["project_id"],
            "primre_name": pname, "primre_link_basis": basis,
            "primre_status": p.get("Project Status", ""), "primre_scale": p.get("Project Scale", ""),
            "primre_website": p.get("Website", ""),
            "review_status": "unreviewed",
            "notes": "",
        })

    # Flag Pilot rows that share one PRIMRE record: either duplicates to merge, or
    # distinct phases PRIMRE lumps together. Left for review, not merged here.
    for row in rows:
        if row["primre_name"] and used_primre[row["primre_name"]] > 1:
            row["notes"] = f"shares PRIMRE record with {used_primre[row['primre_name']] - 1} other row(s)"

    n = len(rows)
    for name, p in primre.items():
        if used_primre[name]:
            continue
        devices = [d.rsplit("/", 1)[-1] for d in p["Devices"].split(",") if d.strip()]
        site = index.match(p["Name"], p["Country"]) or (
            index.match(p["Waterbody"], p["Country"]) if p["Waterbody"] else None)
        n += 1
        rows.append({
            "project_id": f"P{n:04d}",
            "name": p["Name"],
            "aliases": "",
            "developer": p["Project Manager"],
            "device": " | ".join(devices),
            "country": country(p["Country"]),
            "reporting_country": "",
            "technology": PRIMRE_TECH.get(p["Resource"], p["Resource"]),
            "in_scope": "yes",
            "test_site_id": site["site_id"] if site else "",
            "test_site_name_raw": "",
            "location": p["Waterbody"],
            "first_seen": "", "last_seen": "", "latest_status": "",
            "source": "PRIMRE",
            "pilot_project_id": "",
            "primre_name": p["Name"], "primre_link_basis": "PRIMRE record",
            "primre_status": p["Project Status"], "primre_scale": p["Project Scale"],
            "primre_website": p["Website"],
            "review_status": "added from PRIMRE; not yet found in the reports",
            "notes": "",
        })

    cols = ["project_id", "name", "aliases", "developer", "device", "country", "reporting_country",
            "technology", "in_scope", "test_site_id", "test_site_name_raw", "location",
            "first_seen", "last_seen", "latest_status", "source", "pilot_project_id",
            "primre_name", "primre_link_basis", "primre_status", "primre_scale", "primre_website",
            "review_status", "notes"]
    write(OUT / "projects.csv", rows, cols)
    linked = sum(bool(r["primre_name"]) and r["source"].startswith("Pilot") for r in rows)
    print(f"projects.csv: {len(rows)} projects ({len(pilot)} from the Pilot, {linked} of them linked "
          f"to PRIMRE; {len(rows) - len(pilot)} PRIMRE-only)")
    print(f"  with a test site: {sum(bool(r['test_site_id']) for r in rows)}; "
          f"out of scope (technology 'other'): {sum(r['in_scope'] == 'no' for r in rows)}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--force", action="store_true", help="overwrite existing gazetteers")
    args = ap.parse_args()
    existing = [p for p in (OUT / "test_sites.csv", OUT / "projects.csv") if p.exists()]
    if existing and not args.force:
        sys.exit(f"Refusing to overwrite hand-edited gazetteers: {', '.join(p.name for p in existing)}. "
                 "Use --force to reseed from scratch.")
    OUT.mkdir(parents=True, exist_ok=True)
    sites = seed_test_sites()
    seed_projects(sites)


if __name__ == "__main__":
    main()
