#!/usr/bin/env python3
"""Propose search terms for each project in gazetteers/projects.csv.

Project names in the list are mostly descriptions ("Cape Sharp Tidal project at FORCE")
that never appear verbatim in a report. This proposes short, distinctive terms to search
for instead, and tests each one against the cleaned report text in output/text/
(run extract.py first).

Candidates come from the project's name, aliases, device names and PRIMRE names, plus
their bracketed acronyms and the part of a name before " at " or a comma. A candidate
is dropped when it is:
  - descriptive: starts lower-case, has lower-case words other than connectors
    ("of", "de", "in" ...), or is longer than 6 words
  - an ordinary word on its own ("Swell", "Nemo" is kept - not a dictionary word)
  - made only of generic words ("Tidal Energy Project", "WEC")
  - a test-site name or alias (those are matched as test sites), or a country name
  - a place name found more than 60 times ("Zhoushan"): it would match the place, not the project
and flagged, but kept, when it is:
  - shared: proposed for more than one project
  - a place: every word is a place name (Heraklion, Kyle Rhea) - often the project's
    real name, but it also matches any mention of the place
  - frequent: found in more than 60 places - probably not specific to the project

Matching is case-sensitive and whole-word, the same as test-site aliases in extract.py.

The kept terms are then pruned for readability: a term is dropped when a shorter, unflagged
term of the same project is part of it (it can only find the same mentions), and terms with
no hits are dropped when the project has a term that does hit. A project with no hitting
term keeps its two shortest candidates.

Writes:
  validation/search_terms_review.csv  one row per candidate term, with hits and flags
  gazetteers/projects.csv             fills the `search_terms` column where it is empty;
                                      terms already there (hand-edited) are left alone
"""

import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
PROJECTS = HERE / "gazetteers" / "projects.csv"
SITES = HERE / "gazetteers" / "test_sites.csv"
PLACES = HERE.parent / "Annual Report Extraction" / "places_gazetteer.csv"
TEXT_DIR = HERE / "output" / "text"
REVIEW = HERE / "validation" / "search_terms_review.csv"

GENERIC = set("""the of and at in on a an to for by with from project projects pilot demonstration demo
array arrays phase test testing tests deployment plant power station tidal wave waves energy ocean
marine turbine turbines device devices system systems prototype farm park site sites scale full sea
trial trials installation first new converter converters generator generation current currents
stream floating platform project's programme program development developments offshore onshore
nearshore commercial demonstrator unit units module modules hybrid technology technologies pto
wec wecs tec tecs owc otec mw kw gw i ii iii iv""".split())
FREQUENT = 60
# Lower-case words allowed inside a proper name ("Island in the Currents", "Marina di Pisa").
CONNECTORS = {"of", "de", "di", "du", "da", "do", "la", "le", "and", "the", "van", "von", "der",
              "den", "y", "et", "in", "on", "del", "des"}
# Ordinary words the system dictionary may lack.
EXTRA_ORDINARY = {"desalination", "renewables", "hydrokinetic", "photovoltaic", "aquaculture",
                  "decarbonisation", "decarbonization", "microgrid", "subsea"}


def load_words():
    try:
        with open("/usr/share/dict/words", encoding="utf-8") as f:
            return {w.strip() for w in f if w[:1].islower()}
    except OSError:
        return set()


COMMON = load_words()


def split(s):
    return [x.strip() for x in (s or "").split(" | ") if x.strip()]


def variants(text):
    """The text itself, its bracketed acronym, and shorter forms of it."""
    out = [text]
    for a in re.findall(r"\(([A-Za-z][\w.&-]{2,15})\)", text):
        out.append(a)
    base = re.sub(r"\s*\([^)]*\)", "", text).strip()
    out.append(base)
    first = re.split(r"\s+at\s+|,\s+|\s+-\s+|\s+/\s+|\s+–\s+", base)[0].strip()
    out.append(first)
    out.append(trim_generic(first))
    # leading run of name-like tokens: "SRNE 110 kW floating ..." -> "SRNE"
    lead = []
    for w in first.split():
        if re.fullmatch(r"[\d.,]+|[kMG]W", w) or not re.search(r"[A-Z]", w[:1] + w):
            break
        lead.append(w)
    out.append(trim_generic(" ".join(lead)))
    out = [re.sub(r"^(The|the)\s+", "", x).strip(" .,;:-") for x in out]
    return list(dict.fromkeys(x for x in out if x))


def trim_generic(text):
    """Drop trailing generic words and numbers: "Fair Head Phase 1" -> "Fair Head"."""
    ws = text.split()
    while ws and (ws[-1].lower().strip(".,") in GENERIC or re.fullmatch(r"[\d.,]+|[kMG]W|\d+[kMG]W", ws[-1])):
        ws.pop()
    return " ".join(ws)


def words(term):
    return re.findall(r"[\w'’.-]+", term)


def drop_reason(term, site_aliases, countries=frozenset()):
    ws = words(term)
    if len(term) < 3 or not ws:
        return "too short"
    if len(ws) > 6:
        return "descriptive (long)"
    if not re.search(r"[A-Z0-9]", term):
        return "descriptive (lower-case)"
    if any(w[:1].islower() and w not in CONNECTORS for w in ws[1:]):
        return "descriptive (lower-case words)"
    if term in site_aliases:
        return "test-site name"
    if term in countries:
        return "country name"
    core = [w for w in ws if w.lower().strip(".-") not in GENERIC]
    if not core:
        return "generic words only"
    if len(ws) == 1:
        w = ws[0]
        parts = [x for x in re.split(r"[-/]", w.lower()) if x]
        ordinary = all(x in COMMON or x in GENERIC or (x.endswith("s") and x[:-1] in COMMON)
                       for x in parts) or w.lower() in EXTRA_ORDINARY
        if ordinary and not re.search(r".[A-Z]|\d", w):   # keep CorPower, OE35, SR2000
            return "ordinary word"
    return ""


def term_regex(term):
    parts = [re.escape(p) for p in term.split()]
    return re.compile(r"(?<![\w-])" + r"\s+".join(parts) + r"(?![\w-])")


def main():
    if not TEXT_DIR.exists():
        sys.exit("output/text/ not found: run extract.py first")
    texts = {}
    for f in sorted(TEXT_DIR.glob("*.txt")):
        year = re.search(r"(\d{4})", f.name).group(1)
        body = re.sub(r"^\s*\d+  ", "", f.read_text(encoding="utf-8"), flags=re.M)
        texts[year] = re.sub(r"\s+", " ", body)

    sites = list(csv.DictReader(open(SITES, encoding="utf-8")))
    site_aliases = {a for s in sites for a in [s["name"]] + split(s["aliases"])}
    places = set()
    for s in sites:
        places.update(w for w in words(s["location"]) if w[:1].isupper())
    if PLACES.exists():
        for p in csv.DictReader(open(PLACES, encoding="utf-8")):
            for n in [p["name"]] + split(p["aliases"].replace("|", " | ")):
                places.update(words(n))

    rows = list(csv.DictReader(open(PROJECTS, encoding="utf-8")))
    cols = list(rows[0].keys())
    if "search_terms" not in cols:
        cols.insert(cols.index("aliases") + 1, "search_terms")
        for r in rows:
            r["search_terms"] = ""

    # candidates per project
    cands = []           # (project_id, term, came_from)
    for r in rows:
        seen = set()
        for field in ("name", "aliases", "device", "primre_name"):
            for text in split(r[field]) if field != "name" else [r["name"]]:
                for t in variants(text):
                    if t not in seen:
                        seen.add(t)
                        cands.append((r["project_id"], t, field))

    countries = {c for r in rows for c in split(r["country"]) + split(r["reporting_country"])}
    countries |= {s["country"] for s in sites} | {"UK", "USA", "US", "EU", "Europe", "Korea", "Scotland",
                                                  "Wales", "England", "Northern Ireland"}

    def is_place(t):
        return all(w in places or w in countries or w.lower() in GENERIC
                   or w.lower() in ("of", "de", "di", "la") for w in words(t))

    cache = {}
    def hits(term):
        if term not in cache:
            rx = term_regex(term)
            n, ys = 0, []
            for y, s in texts.items():
                if term.split()[0] in s:
                    k = len(rx.findall(s))
                    if k:
                        n += k
                        ys.append(y)
            cache[term] = (n, ys)
        return cache[term]

    def reason_for(t):
        r = drop_reason(t, site_aliases, countries)
        if not r and is_place(t) and hits(t)[0] > FREQUENT:
            r = f"busy place name ({hits(t)[0]} hits)"
        return r

    owners = defaultdict(set)
    for pid, t, _ in cands:
        if not reason_for(t):
            owners[t].add(pid)

    review, kept = [], defaultdict(list)
    for pid, t, field in cands:
        reason = reason_for(t)
        n, ys = hits(t) if not reason else (0, [])
        flags = []
        if not reason:
            if len(owners[t]) > 1:
                flags.append("shared with " + ", ".join(sorted(owners[t] - {pid})))
            if is_place(t):
                flags.append("place name")
            if n > FREQUENT:
                flags.append(f"frequent ({n})")
        row = {"project_id": pid, "term": t, "from": field,
               "status": "dropped: " + reason if reason else "kept",
               "hits": n, "years": f"{ys[0]}-{ys[-1]}" if ys else "",
               "n_years": len(ys), "flags": "; ".join(flags)}
        review.append(row)
        if not reason:
            kept[pid].append(row)

    # prune for readability
    keep = {}
    for pid, rs in kept.items():
        hitting = [r for r in rs if r["hits"] > 0]
        if hitting:
            pool = hitting
        else:
            pool = sorted(rs, key=lambda r: len(r["term"]))[:2]
        chosen = []
        for r in sorted(pool, key=lambda r: len(r["term"])):
            covered = any(not c["flags"] and term_regex(c["term"]).search(r["term"]) for c in chosen)
            if not covered:
                chosen.append(r)
        for r in rs:
            if r not in chosen:
                r["status"] = "kept, pruned"
        keep[pid] = [r["term"] for r in chosen]

    filled = 0
    for r in rows:
        if not r["search_terms"].strip() and keep.get(r["project_id"]):
            r["search_terms"] = " | ".join(keep[r["project_id"]])
            filled += 1

    with open(REVIEW, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(review[0].keys()))
        w.writeheader()
        w.writerows(review)
    with open(PROJECTS, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    print(f"{len(cands)} candidate terms, {sum(r['status'] == 'kept' for r in review)} kept; "
          f"search_terms filled for {filled} projects")


if __name__ == "__main__":
    main()
