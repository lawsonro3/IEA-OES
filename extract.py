#!/usr/bin/env python3
"""Extract test sites and projects from the IEA-OES annual report RTFs.

Deterministic, local parsing only (no LLM, no network). Writes one Excel workbook.

Pipeline, per report:
  1. RTF -> plain text (macOS `textutil`), then cleaned: page headers/footers removed,
     ligatures and quotes normalised.
  2. Each line is tagged with the country chapter and section heading it sits under.
  3. Test sites: every mention of a site in test_sites_gazetteer.csv is recorded.
     Names that look like test sites but are not in the gazetteer are listed as
     candidates for review, so the gazetteer can be grown.
  4. Projects: entries under project-type headings (TECHNOLOGY DEMONSTRATION,
     PROJECTS IN THE WATER, PLANNED DEPLOYMENTS, ...) are split into one row per
     entry, with capacity, technology, status keywords, years and test sites pulled
     out by pattern. Where a chapter has no such headings (mostly pre-2013 reports),
     paragraphs are kept only if they look like a deployment (a capacity or test site
     plus a deployment verb).
  5. Named projects: every search term in gazetteers/projects.csv is found in the text
     (case-sensitive, whole words, longest first). A term owned by one project counts
     for it; a term shared by several projects, or marked weak with "~" (organisations,
     places), counts only when the text around it names something else about one of them
     (another of its terms, or its test site) - otherwise the mention is left ambiguous.
     The resolved mentions give one row per project per report year.

Every row carries the source file and line number of the cleaned text, which is
written alongside the workbook (output/text/) so any row can be traced back.

Started as a copy of Annual Report Extraction/extract_annual_reports.py (kept there,
unchanged, as the baseline).

Usage:
    python3 extract.py
    python3 extract.py --input "<RTF folder>" --output out.xlsx
"""

import argparse
import bisect
import csv
import re
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path
from statistics import median

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
DEFAULT_INPUT = HERE.parent / "OES Publications" / "Annual Reports" / "Annual Report RTF"
DEFAULT_OUTPUT = HERE / "output" / "oes_test_sites_and_projects.xlsx"
DEFAULT_GAZETTEER = HERE / "gazetteers" / "test_sites.csv"
DEFAULT_PROJECTS = HERE / "gazetteers" / "projects.csv"
PROJECT_WINDOW = 600  # characters either side of a shared term searched for context

# --------------------------------------------------------------------------------------
# Vocabulary
# --------------------------------------------------------------------------------------

# Heading text (as it appears in the reports) -> canonical country name.
COUNTRIES = {
    "australia": "Australia",
    "belgium": "Belgium",
    "canada": "Canada",
    "china": "China",
    "denmark": "Denmark",
    "european commission": "European Commission",
    "france": "France",
    "germany": "Germany",
    "india": "India",
    "ireland": "Ireland",
    "italy": "Italy",
    "japan": "Japan",
    "mexico": "Mexico",
    "monaco": "Monaco",
    "netherlands": "Netherlands",
    "the netherlands": "Netherlands",
    "new zealand": "New Zealand",
    "nigeria": "Nigeria",
    "norway": "Norway",
    "portugal": "Portugal",
    "republic of korea": "Republic of Korea",
    "korea": "Republic of Korea",
    "singapore": "Singapore",
    "south africa": "South Africa",
    "spain": "Spain",
    "sweden": "Sweden",
    "united kingdom": "United Kingdom",
    "uk": "United Kingdom",
    "united states": "United States",
    "united states of america": "United States",
    "usa": "United States",
}

# A country heading only opens a chapter if one of these follows within a few lines.
# This separates real chapters from the country names in contents pages, tables and
# delegate lists.
CHAPTER_MARKER_RE = re.compile(
    r"^(authors?(\s*\(s\))?\s*:?|report prepared by\s*:?|this report (has been|was) prepared"
    r"|overview|introduction|ocean energy policy|national strategy|supporting policies"
    r"|introductory note)\b",
    re.I,
)
CHAPTER_LOOKAHEAD = 10
OUTSIDE = "(outside country chapters)"
ISOLATION = 15

# Lines that end the country-report part of a report.
BACK_MATTER_RE = re.compile(
    r"^(appendi(x|ces)\b|membership of the executive committee|exco members"
    r"|e x c o\b|executive committee members)",
    re.I,
)

# Section headings. Matched against the whole (normalised) line.
# Value = section type; None means "a non-project section" (ends a project section).
SECTION_HEADINGS = [
    (r"technology demonstration", "demonstration"),
    (r"research,? development and (technology )?demonstration", "demonstration"),
    (r"(arrays and )?(demonstration )?projects in the water", "deployment"),
    (r"operational (ocean energy )?(projects|deployments)", "deployment"),
    (r"(ongoing |completed |large |demonstration |ocean energy )?projects", "deployment"),
    (r"planned (deployments|projects)", "planned"),
    (r"consented( projects)?", "planned"),
    (r"(existing )?(open )?sea test (sites|facilities)", "test sites"),
    (r"(national )?(open )?(sea )?test (sites|centres|centers|facilities)( and technology demonstration)?", "test sites"),
    (r"test centres? (&|and) demonstration zones", "test sites"),
    (r"(key |relevant )?r ?& ?d projects", "R&D"),
    (r"key r&d institutions and relevant r&d projects", "R&D"),
    # Non-project sections: these end the current project section.
    (r"overview|introduction|introductory note|summary|background", None),
    (r"ocean energy policy|policies|regulatory framework|marine spatial planning( policy)?", None),
    (r"national strategy( and targets)?|supporting policies( for ocean energy)?", None),
    (r"market incentives|public funding programm?e?s|public funding", None),
    (r"research (&|and) development|r ?& ?d", None),
    (r"relevant national events|other relevant national activities", None),
    (r"specific initiatives for international cooperation|international cooperation", None),
    (r"further information|references|bibliography|contacts?", None),
    (r"authors?(\s*\(s\))?|report prepared by", None),
]
PROJECT_KINDS = {"demonstration", "deployment", "planned", "R&D"}
SECTION_RES = [(re.compile(r"^(" + p + r")\s*:?$", re.I), t) for p, t in SECTION_HEADINGS]

# Page furniture that textutil leaves glued onto body text.
FOOTER_RES = [
    re.compile(r"IEA\s*\|\s*OES\s*-\s*ANNUAL REPORT 20\d\d", re.I),
    re.compile(r"ANNUAL REPORT 20\d\d\s*[—–-]\s*\d{1,3}", re.I),
    re.compile(r"\d{1,3}\s*[—–-]\s*OCEAN ENERGY SYSTEMS", re.I),
    re.compile(r"\d{1,3}#?\s*annual report 20\d\d", re.I),
    re.compile(r"Annual Report 20\d\d\s?\d{1,3}", re.I),
    re.compile(r"\b\d{1,3}\s?Ocean Energy Systems(?=\S|$)", re.I),
]

CHAR_FIXES = {
    "ﬀ": "ff", "ﬁ": "fi", "ﬂ": "fl", "ﬃ": "ffi", "ﬄ": "ffl",
    "’": "'", "‘": "'", "“": '"', "”": '"', " ": " ",
    " ": "\n", " ": "\n", "­": "", "\f": "\n", "\v": "\n",
}

TECHNOLOGIES = [
    ("wave", r"\bwave\b|\bWEC\b|oscillating water column|\bOWC\b|point absorber"),
    ("tidal", r"\btidal\b"),
    ("river / hydrokinetic", r"\briver(ine)?\b|hydrokinetic|in-stream"),
    ("ocean current", r"ocean current|\bcurrent turbine"),
    ("OTEC / thermal", r"\bOTEC\b|ocean thermal"),
    ("salinity gradient", r"salinity gradient|osmotic|blue energy|reverse electrodialysis"),
    ("other (wind / solar / hybrid)", r"offshore wind|floating wind|floating solar|\bFPV\b|photovoltaic"),
]
TECH_RES = [(name, re.compile(p, re.I)) for name, p in TECHNOLOGIES]

STATUS_WORDS = {
    "planned": r"\b(will|plan(s|ned)?|expected|anticipated?|scheduled|aim(s|ing)?|intend(s|ed)?|propos(ed|al)|to be (deployed|installed))\b",
    "consented": r"\b(consent(ed)?|licen[cs]ed|permit(ted)?|lease)\b",
    "under construction": r"\b(under construction|being built|fabricat(ed|ion)|manufactur(ed|ing)|assembl(ed|y))\b",
    "deployed / installed": r"\b(deploy(ed|ment)|install(ed|ation)|launched|submerged|moored)\b",
    "grid connected": r"\bgrid[- ]connect(ed|ion)\b|exported? (electricity|power) to the grid",
    "operating / testing": r"\b(operat(ing|ional|ion)|generat(ed|ing)|produc(ed|ing)|test(ed|ing)|trials?)\b",
    "completed / removed": r"\b(decommission(ed|ing)|removed|retrieved|recovered|completed|ended|concluded)\b",
}
STATUS_RES = [(k, re.compile(p, re.I)) for k, p in STATUS_WORDS.items()]

DEPLOY_VERB_RE = re.compile(
    r"\b(deploy(ed|ment)?|install(ed|ation)?|commission(ed|ing)?|grid[- ]connect(ed|ion)|"
    r"launched|tested|testing|operat(ing|ional)|decommission(ed)?|retrieved|redeployed|"
    r"demonstrat(ion|ed|or))\b",
    re.I,
)
DEVICE_WORD_RE = re.compile(
    r"\b(device|prototype|turbine|converter|WEC|array|buoy|plant|project|demonstrator|"
    r"farm|platform|generator|barrage|OWC|kite)s?\b",
    re.I,
)
CAPACITY_RE = re.compile(r"(?<![\w.])(\d{1,3}(?:[ ,]\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?)\s?(kW|MW|GW)\b")
YEAR_RE = re.compile(r"\b(19[89]\d|20[0-4]\d)\b")

# Words that start sentences or are otherwise not names, for the named-entity column.
NOT_NAMES = set(
    """a an the this these that those in on at of for from to by with and or but as it its
    their there they we our he she his her during since after before while when where which
    who what how all also some several each other another further however in addition
    january february march april may june july august september october november december
    spring summer autumn fall winter figure fig table source photo image courtesy credit
    annual report ocean energy systems oes iea government ministry department""".split()
)
NAME_RE = re.compile(
    r"\b(?:[A-Z][\w®'·.-]*|[a-z]+[A-Z][\w®'·.-]*)"
    r"(?:\s+(?:(?:of|de|du|da|do|la|le|for|and|&)\s+)?(?:[A-Z][\w®'·.-]*|\d[\w.-]*))*"
)

# ---- Location vocabulary
# Words that make a capitalised phrase a place: "Kvichak River", "Bay of Fundy".
GEO_WORDS = (
    "River|Bay|Strait|Straits|Sound|Firth|Channel|Passage|Island|Islands|Isle|Isles|Harbour|Harbor|"
    "Port|Coast|Lough|Loch|Estuary|Sea|Ocean|Canal|Fjord|Fiord|Inlet|Lake|Gulf|Head|Point|Peninsula|"
    "Reef|Cove|Waterway|Barrier|Race|Narrows|Archipelago|Lagoon|Shoal|Banks?|Cape"
)
GEO_WORD_RE = re.compile(r"\b(" + GEO_WORDS + r")\b")

# Sub-national regions, used to recognise "Igiugig, Alaska" style places.
REGIONS = """Alabama Alaska California Connecticut Delaware Florida Georgia Hawaii Louisiana Maine
Maryland Massachusetts Michigan Minnesota New_Hampshire New_Jersey New_York North_Carolina Oregon
Rhode_Island South_Carolina Texas Virginia Washington Puerto_Rico
Nova_Scotia New_Brunswick Newfoundland Labrador Newfoundland_and_Labrador British_Columbia Manitoba
Quebec Ontario Nunavut Prince_Edward_Island
Queensland New_South_Wales Victoria Tasmania Western_Australia South_Australia Northern_Territory
Scotland England Wales Northern_Ireland Orkney Shetland Cornwall Devon Pembrokeshire Anglesey
Galicia Asturias Cantabria Basque_Country Canary_Islands Andalucía Andalucia Catalonia
Brittany Normandy Jeju Jindo Zhejiang Shandong Guangdong Fujian Hainan Zhoushan
Nagasaki Okinawa Iwate Kyushu Hokkaido Tamil_Nadu Gujarat Kerala
Azores Madeira Sardinia Sicily Tuscany Liguria Calabria Campania Lombardy
Baja_California Quintana_Roo Yucatán Yucatan Cozumel""".split()
REGIONS = {r.replace("_", " ") for r in REGIONS}

PLACE_TOKEN = r"[A-Z][\w'’.-]*"
PLACE_PHRASE = PLACE_TOKEN + r"(?:\s+(?:(?:of|de|du|da|do|del|la|le|the|and|on|upon)\s+)?" + PLACE_TOKEN + r"){0,4}"
LOCATION_RE = re.compile(
    r"\b(in|at|off|near|into|on|along|from|within|outside|around|across)\s+"
    r"(?:the\s+)?(?:(?:coast|shores?|waters|port|harbour|harbor|mouth|island|region|area)\s+(?:of|off)\s+(?:the\s+)?)?"
    r"(" + PLACE_PHRASE + r"(?:,\s+" + PLACE_PHRASE + r"){0,2})"
)
# Phrases that look like places after "in"/"at" but are organisations, programmes, etc.
NOT_PLACE_RE = re.compile(
    r"\b(Ltd|Limited|Inc|GmbH|University|Universidad|Institute|Ministry|Department|Agency|Council|"
    r"Programme|Program|Project|Fund|Commission|Association|Group|Company|Corporation|Energy|Power|"
    r"Technolog\w*|Conference|Workshop|Committee|Report|Plan|Strategy|Act|Call|Horizon|Interreg|"
    r"Laboratory|Lab|Tank|Basin|Flume|Directive|Phase|Stage|Annex|Task|Network|Initiative|Centre|Center|"
    r"Platform|Site|Zone|Facility)\b"
)
COORD_RE = re.compile(
    r"\d{1,3}(?:\.\d+)?\s?°\s?(?:\d{1,2}(?:\.\d+)?\s?['′]\s?)?(?:\d{1,2}(?:\.\d+)?\s?[\"″]\s?)?[NSEW]\b"
    r"(?:[,;\s]+\d{1,3}(?:\.\d+)?\s?°\s?(?:\d{1,2}(?:\.\d+)?\s?['′]\s?)?(?:\d{1,2}(?:\.\d+)?\s?[\"″]\s?)?[NSEW]\b)?"
)
DEPTH_RE = re.compile(
    r"(?:depths?\s+of\s+(?:about\s+|approximately\s+|around\s+|up\s+to\s+)?(\d+(?:[.,]\d+)?(?:\s?[-–]\s?\d+)?)\s?m\b"
    r"|(\d+(?:[.,]\d+)?(?:\s?[-–]\s?\d+)?)\s?m(?:etres?|eters?)?\s+(?:of\s+)?(?:water\s+)?depths?\b"
    r"|(\d+(?:[.,]\d+)?(?:\s?[-–]\s?\d+)?)\s?m(?:etres?|eters?)?\s+deep\b)",
    re.I,
)
OFFSHORE_RE = re.compile(
    r"(\d+(?:[.,]\d+)?(?:\s?[-–]\s?\d+)?)\s?(km|kilometres?|kilometers?|nm|nautical miles?|miles?)\s+"
    r"(?:off\s*shore|from\s+(?:the\s+)?(?:shore|coast(?:line)?)|off\s+(?:the\s+)?coast)",
    re.I,
)

# Pattern for test-site-like names that may be missing from the gazetteer.
CANDIDATE_RE = re.compile(
    r"((?:(?:[A-Z][\w'.&-]*|of|and|de|do|da|for|du|la|the)\s+){0,6}"
    r"(?:[A-Z][\w'.&-]*)\s+"
    r"(?:[Tt]est (?:[Ss]ite|[Cc]ent(?:re|er)|[Ff]acility|[Aa]rea|[Zz]one|[Bb]erth|[Bb]ed)s?"
    r"|[Mm]arine [Ee]nergy [Cc]ent(?:re|er)|[Dd]emonstration [Zz]one|[Tt]esting [Cc]ent(?:re|er)"
    # Laboratories and tanks count as test sites too.
    r"|[Ll]aborator(?:y|ies)|Lab|(?:[Tt]owing|[Ww]ave|[Tt]est|[Ff]lume) [Tt]anks?"
    r"|(?:[Oo]cean|[Ww]ave|[Cc]oastal|[Tt]est) [Bb]asins?|[Ww]ave [Ff]lumes?))"
    r"(\s?\([A-Z][\w-]{1,12}\))?"
)
CANDIDATE_STRIP_LEAD = re.compile(
    r"^((the|at|in|of|and|to|a|an|from|for|with|on|by|test|site|name|location|"
    r"technology|demonstration|existing|open|sea|systems|" + "|".join(
        re.escape(c) for c in sorted(COUNTRIES, key=len, reverse=True)) + r")\s+)+",
    re.I,
)


# --------------------------------------------------------------------------------------
# Text preparation
# --------------------------------------------------------------------------------------

def rtf_to_text(path):
    if not shutil.which("textutil"):
        sys.exit("textutil not found: this script uses macOS's built-in RTF converter.")
    out = subprocess.run(
        ["textutil", "-convert", "txt", "-stdout", str(path)],
        check=True, capture_output=True,
    )
    return out.stdout.decode("utf-8", errors="replace")


UPPER_COUNTRY_ALT = "|".join(re.escape(c.upper()) for c in sorted(COUNTRIES, key=len, reverse=True))
# "67DENMARKOVERVIEW", "ITALYSUPPORTING POLICIES": a page number and/or a country heading
# glued onto the next heading.
GLUED_PAGE_RE = re.compile(r"^\d{1,3}(?=[A-Z]{3})")
GLUED_REPORT_RE = re.compile(r"^(20\d\d\s*)?ANNUAL REPORT(?=[A-Z]{3})")
GLUED_COUNTRY_RE = re.compile(
    r"^(" + UPPER_COUNTRY_ALT + r")\s*(?=(OVERVIEW|INTRODUCTION|SUPPORTING POLICIES( FOR OCEAN ENERGY)?"
    r"|OCEAN ENERGY POLICY|NATIONAL STRATEGY|AUTHORS?|REPORT PREPARED BY)\s*:?\s*$)",
    re.I,
)


def clean_text(raw):
    for bad, good in CHAR_FIXES.items():
        raw = raw.replace(bad, good)
    for rx in FOOTER_RES:
        raw = rx.sub("\n", raw)
    lines = []
    for ln in raw.split("\n"):
        ln = re.sub(r"[ \t]+", " ", ln).strip()
        ln = GLUED_REPORT_RE.sub("", GLUED_PAGE_RE.sub("", ln))
        ln = re.sub(r"^\d{1,3} 0\d COUNTRY REPORTS(?=[A-Z]{3})", "", ln)  # 2013 running header
        m = GLUED_COUNTRY_RE.match(ln)
        if m:
            lines.append(m.group(1).strip())
            ln = ln[m.end():]
        if ln.strip():
            lines.append(ln)
    return lines


def normalise_heading(line):
    s = line.strip()
    s = re.sub(r"^(\d+(\.\d+)*\.?|[IVX]+\.)\s*", "", s)  # chapter/section numbers
    s = re.sub(r"\s+", " ", s)
    return s.strip(" :")


def section_type(line):
    """Return (is_heading, type) for a line."""
    s = normalise_heading(line)
    if len(s) > 80 or not s:
        return False, None
    for rx, typ in SECTION_RES:
        if rx.match(s):
            return True, typ
    return False, None


# --------------------------------------------------------------------------------------
# Structure: chapters and sections
# --------------------------------------------------------------------------------------

def find_chapters(lines):
    """Return list of (start, end, country, author_text)."""
    starts = []
    for i, ln in enumerate(lines):
        key = normalise_heading(ln).lower()
        skip = 1
        if key not in COUNTRIES and i + 1 < len(lines):
            # Country names wrapped over two lines ("EUROPEAN" / "COMMISSION").
            key = (key + " " + normalise_heading(lines[i + 1]).lower()).strip()
            skip = 2
        if key not in COUNTRIES:
            continue
        ahead = lines[i + skip:i + skip + CHAPTER_LOOKAHEAD]
        if any("TEST SITE NAME" in a.upper() for a in ahead[:3]):
            continue
        # The next line being another country name means a list, not a chapter.
        if ahead and normalise_heading(ahead[0]).lower() in COUNTRIES:
            continue
        marks = [CHAPTER_MARKER_RE.match(normalise_heading(a)) for a in ahead]
        if any(marks):
            has_author = any(m and re.match(r"author|report prepared|this report", m.group(0), re.I)
                             for m in marks)
            starts.append((i, COUNTRIES[key], has_author))

    # Country names in summary tables ("China / Plan for ...", "Denmark / National
    # Strategy ...") can pass the marker test. Real chapter headings stand alone: drop a
    # candidate with another country's candidate within ISOLATION lines of it.
    starts = [
        (i, c, a) for i, c, a in starts
        if not any(c2 != c and abs(i2 - i) < ISOLATION for i2, c2, _ in starts)
    ]

    # Where a country still has several candidates (running headers, repeated
    # summaries), prefer one followed by an author line, then the longest span.
    def rank(n):
        span = (starts[n + 1][0] if n + 1 < len(starts) else len(lines)) - starts[n][0]
        return (starts[n][2], span)

    best = {}
    for n, (i, c, _) in enumerate(starts):
        if c not in best or rank(n) > rank(best[c]):
            best[c] = n
    dedup = sorted(starts[n][:2] for n in best.values())

    chapters = []
    for n, (i, c) in enumerate(dedup):
        end = dedup[n + 1][0] if n + 1 < len(dedup) else len(lines)
        for j in range(i + 1, end):
            if BACK_MATTER_RE.match(normalise_heading(lines[j])):
                end = j
                break
        author = ""
        for j in range(i + 1, min(i + 1 + CHAPTER_LOOKAHEAD, end)):
            s = normalise_heading(lines[j])
            if re.match(r"^(authors?|report prepared by)", s, re.I):
                rest = re.sub(r"^(authors?(\s*\(s\))?|report prepared by)\s*:?\s*", "", s, flags=re.I)
                nxt = [normalise_heading(x) for x in lines[j + 1:j + 4]]
                nxt = [x for x in nxt if not CHAPTER_MARKER_RE.match(x)]
                author = "; ".join([rest] + nxt[:2] if rest else nxt[:3]).strip("; ")
                break
        chapters.append((i, end, c, author))
    return chapters


def tag_lines(lines, chapters):
    """Per line: country (or ''), section heading text, section type."""
    country = [""] * len(lines)
    for s, e, c, _ in chapters:
        for k in range(s, e):
            country[k] = c
    heading = [""] * len(lines)
    stype = [None] * len(lines)
    cur_h, cur_t, cur_c = "", None, None
    for k, ln in enumerate(lines):
        if country[k] != cur_c:  # a new chapter resets the section
            cur_h, cur_t, cur_c = "", None, country[k]
        is_h, t = section_type(ln)
        if "TEST SITE NAME" in ln.upper():
            is_h, t = True, "test sites"
            ln = "Test site table"
        if is_h:
            # A test-site subheading inside a project section ("PROJECTS IN THE WATER" /
            # "Open Sea Test Sites") is a subsection: its entries are still projects.
            if t == "test sites" and cur_t in PROJECT_KINDS:
                t = cur_t
            cur_h, cur_t = normalise_heading(ln), t
        heading[k], stype[k] = cur_h, cur_t
    return country, heading, stype


# --------------------------------------------------------------------------------------
# Test sites
# --------------------------------------------------------------------------------------

def load_gazetteer(path):
    sites = []
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            aliases = [a.strip() for a in row["aliases"].split("|") if a.strip()]
            sites.append({**row, "alias_list": aliases})
    return sites


def alias_regex(alias):
    parts = [re.escape(p) for p in alias.replace("'", "\u0000").split()]
    body = r"\s+".join(parts).replace("\u0000", "'")
    return re.compile(r"(?<![\w-])" + body + r"(?![\w-])")


def build_alias_index(sites):
    """alias -> (compiled regex, [site, ...]). An alias may belong to several sites."""
    index = {}
    for s in sites:
        for a in s["alias_list"]:
            if a not in index:
                index[a] = (alias_regex(a), [])
            index[a][1].append(s)
    # Longest aliases first, so "KRISO-WETS" wins over "WETS".
    return sorted(index.items(), key=lambda kv: -len(kv[0]))


class Joined:
    """The report as one string (lines joined by spaces), with a map back to lines."""

    def __init__(self, lines):
        self.starts, parts, pos = [], [], 0
        for ln in lines:
            self.starts.append(pos)
            parts.append(ln)
            pos += len(ln) + 1
        self.text = " ".join(parts)

    def line_of(self, offset):
        return bisect.bisect_right(self.starts, offset) - 1

    def context(self, a, b, width=160):
        s, e = max(0, a - width), min(len(self.text), b + width)
        # Snap to sentence edges where possible.
        left = self.text.rfind(". ", s, a)
        s = left + 2 if left != -1 else s
        right = self.text.find(". ", b, e)
        e = right + 1 if right != -1 else e
        return self.text[s:e].strip()


def find_site_mentions(joined, alias_index, country_of_line):
    taken = []  # (start, end) spans already claimed

    def overlaps(a, b):
        return any(a < te and b > ts for ts, te in taken)

    found = []
    for alias, (rx, sites) in alias_index:
        for m in rx.finditer(joined.text):
            a, b = m.span()
            if overlaps(a, b):
                continue
            line = joined.line_of(a)
            chapter_country = country_of_line[line]
            site = sites[0]
            if len(sites) > 1:  # shared alias: prefer the site in the chapter's country
                site = next((s for s in sites if s["country"] == chapter_country), sites[0])
            taken.append((a, b))
            found.append((a, b, line, site, m.group(0)))
    found.sort(key=lambda t: t[0])
    return found


def find_site_candidates(joined, taken_spans):
    out = []
    for m in CANDIDATE_RE.finditer(joined.text):
        a, b = m.span()
        if any(a < te and b > ts for ts, te in taken_spans):
            continue
        name = CANDIDATE_STRIP_LEAD.sub("", m.group(0)).strip(" ,.;:")
        words = name.split()
        if len(words) < 2 or not re.match(r"[A-Z]", name):
            continue
        if name.isupper() and "TEST SITE" in name:  # table header debris
            continue
        out.append((a, b, joined.line_of(a), name))
    return out


# --------------------------------------------------------------------------------------
# Named projects (gazetteers/projects.csv)
# --------------------------------------------------------------------------------------

def load_projects(path):
    projects = []
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            raw = [t.strip() for t in row["search_terms"].split("|") if t.strip()]
            # "~Term" marks a weak term (an organisation, a place, a generic name): it only
            # counts when the passage around it also names something else about the project.
            terms = [t.lstrip("~").strip() for t in raw]
            weak = {t.lstrip("~").strip() for t in raw if t.startswith("~")}
            projects.append({**row, "term_list": terms, "weak_terms": weak})
    return projects


def build_project_index(projects, sites):
    """One regex for every search term (longest first), the projects owning each term, and
    for each project the cues that point to it when a term is shared: its own terms and its
    test site's aliases."""
    owners = defaultdict(list)
    for p in projects:
        for t in p["term_list"]:
            owners[t].append(p)
    site_by_id = {s["site_id"]: s for s in sites}
    for p in projects:
        cues = set(p["term_list"])
        site = site_by_id.get(p["test_site_id"])
        if site:
            cues.update(site["alias_list"])
        p["cue_res"] = {c: alias_regex(c) for c in cues}
    terms = sorted(owners, key=len, reverse=True)
    body = "|".join(r"\s+".join(re.escape(w) for w in t.split()) for t in terms)
    return re.compile(r"(?<![\w-])(?:" + body + r")(?![\w-])"), owners


def resolve_project(term, owners, window, chapter_country):
    """Pick the project a shared term refers to, from what surrounds it. Returns
    (project or None, how it was resolved, [candidate ids])."""
    if len(owners) == 1 and term not in owners[0]["weak_terms"]:
        return owners[0], "unique term", [owners[0]["project_id"]]
    scored = []
    for p in owners:
        cues = sorted(c for c, rx in p["cue_res"].items() if c != term and rx.search(window))
        score = len(cues)
        if cues and chapter_country and chapter_country in (p["country"], p["reporting_country"]):
            score += 0.5
            cues.append(f"chapter {chapter_country}")
        scored.append((score, p, cues))
    scored.sort(key=lambda x: -x[0])
    ids = [p["project_id"] for _, p, _ in scored]
    best = scored[0][0]
    second = scored[1][0] if len(scored) > 1 else 0
    if best >= 1 and best > second:
        return scored[0][1], "context: " + ", ".join(scored[0][2]), ids
    return None, ("ambiguous" if len(owners) > 1 else "weak term, no context"), ids


def find_project_mentions(joined, project_index, country_of_line, window=PROJECT_WINDOW):
    rx, owners = project_index
    found = []
    for m in rx.finditer(joined.text):
        a, b = m.span()
        term = " ".join(m.group(0).split())
        line = joined.line_of(a)
        ctx = joined.text[max(0, a - window):b + window]
        project, how, ids = resolve_project(term, owners[term], ctx, country_of_line[line])
        found.append((a, b, line, term, project, how, ids))
    return found


# --------------------------------------------------------------------------------------
# Projects
# --------------------------------------------------------------------------------------

def is_title_line(line, prev, nxt, typical=60):
    s = line.strip()
    words = s.split()
    if not s or len(s) > min(70, 0.7 * typical) or len(words) > 10:
        return False
    if s[-1] in ".,;:-" or not re.match(r"[A-Z0-9\"(]", s):
        return False
    if re.match(r"^(annual report|ocean energy systems|\d+)$", s, re.I):
        return False
    # Titles are names: mostly capitalised words ("Verdant Power", "MeyGen Phase 1A").
    caps = sum(1 for w in words if re.match(r"[A-Z0-9(\"]", w))
    if caps / len(words) < 0.5:
        return False
    if re.match(r"^(©|\(c\)|courtesy|photo|image|source|figure|fig\.|table)\b", s, re.I):
        return False
    if nxt is None or len(nxt.strip()) < 25:  # a title is followed by body text
        return False
    return prev is None or prev.rstrip()[-1:] in ".!?:)" or len(prev.strip()) < 40


def split_entries(block):
    """block = [(line_no, text)] -> list of entries {title, lines:[(no,text)]}."""
    if not block:
        return []
    lengths = [len(t) for _, t in block if len(t) > 20]
    typical = median(lengths) if lengths else 60
    entries, cur = [], None
    for k, (no, t) in enumerate(block):
        prev = block[k - 1][1] if k else None
        nxt = block[k + 1][1] if k + 1 < len(block) else None
        stripped = t.strip()
        if stripped in {"•", "", "-", "–", "▪", "●"}:
            continue
        bullet = bool(re.match(r"^\s*[•▪●►]\s*", t))
        if is_title_line(t, prev, nxt, typical):
            cur = {"title": stripped, "lines": []}
            entries.append(cur)
            continue
        new_para = (
            cur is None
            or bullet
            or (prev is not None and prev.rstrip()[-1:] in ".!?" and len(prev) < 0.75 * typical
                and re.match(r"\s*[A-Z0-9\"]", t) and cur["lines"])
        )
        if new_para and not (cur is not None and cur["title"] and not cur["lines"]):
            cur = {"title": "", "lines": []}
            entries.append(cur)
        cur["lines"].append((no, re.sub(r"^\s*[•▪●►]\s*", "", t)))
    return [e for e in entries if e["lines"]]


def join_lines(texts):
    out = ""
    for t in texts:
        t = t.strip()
        if out.endswith("-") and t[:1].islower():
            out = out[:-1] + t
        else:
            out = (out + " " + t).strip()
    return out


def named_entities(text, limit=12):
    seen, out = set(), []
    for m in NAME_RE.finditer(text):
        n = m.group(0).strip(" .-")
        words = n.split()
        while words and words[0].lower() in NOT_NAMES:
            words = words[1:]
        n = " ".join(words)
        if len(n) < 2 or n.lower() in NOT_NAMES or n.lower() in COUNTRIES or YEAR_RE.fullmatch(n):
            continue
        if n.lower() not in seen:
            seen.add(n.lower())
            out.append(n)
        if len(out) >= limit:
            break
    return out


def _load_words():
    try:
        with open("/usr/share/dict/words", encoding="utf-8") as f:
            return {w.strip() for w in f if w[:1].islower()}
    except OSError:
        return set()


COMMON_WORDS = _load_words()  # lower-case English words, where the OS provides a list


def plausible_name(s):
    """Filter for name_guess: drop numbers, units, and single ordinary words ("Many")."""
    s = (s or "").strip()
    if len(re.sub(r"[^A-Za-z]", "", s)) < 3 or CAPACITY_RE.fullmatch(s) or re.fullmatch(r"[kMG]W", s):
        return False
    words = s.split()
    if len(words) == 1:
        w = words[0]
        low = w.lower()
        if low in NOT_NAMES or low in COUNTRIES:
            return False
        ordinary = low in COMMON_WORDS or (low.endswith("s") and low[:-1] in COMMON_WORDS)
        if ordinary and not re.search(r"[A-Z].*[A-Z]|\d", w):  # keep OTEC, CorPower, PENGUIN2
            return False
    return True


def describe(text):
    caps = ["{} {}".format(m.group(1).replace(" ", ""), m.group(2)) for m in CAPACITY_RE.finditer(text)]
    techs = [name for name, rx in TECH_RES if rx.search(text)]
    status = [k for k, rx in STATUS_RES if rx.search(text)]
    years = sorted(set(YEAR_RE.findall(text)))
    return caps, techs, status, years


def signal_score(text, has_site):
    return (
        bool(CAPACITY_RE.search(text))
        + bool(DEPLOY_VERB_RE.search(text))
        + bool(DEVICE_WORD_RE.search(text))
        + bool(has_site)
    )


# Place names known in advance: regions above, plus every place in the gazetteer's
# location column. Filled in by main() once the gazetteer is loaded.
KNOWN_PLACES = set(REGIONS)
COUNTRY_MENTION_RE = re.compile(
    r"\b(" + "|".join(sorted({k.title() for k in COUNTRIES if k not in {"korea", "uk", "usa"}}
                             | {"Korea", "UK", "USA", "U.S.", "US", "The Netherlands", "Holland",
                                "Scotland", "England", "Wales", "Northern Ireland"}, key=len, reverse=True)
                    ).replace(".", r"\.") + r")\b"
)
COUNTRY_MENTION_MAP = {k.title(): v for k, v in COUNTRIES.items()}
COUNTRY_MENTION_MAP.update({
    "Korea": "Republic of Korea", "UK": "United Kingdom", "USA": "United States", "U.S.": "United States",
    "US": "United States", "The Netherlands": "Netherlands", "Holland": "Netherlands",
    "Scotland": "United Kingdom", "England": "United Kingdom", "Wales": "United Kingdom",
    "Northern Ireland": "United Kingdom", "Republic Of Korea": "Republic of Korea",
    "United States Of America": "United States",
})


def register_places(sites):
    for s in sites:
        for part in re.split(r"[,/;]| in | of the ", s["location"]):
            part = part.strip()
            if part and part[0].isupper():
                KNOWN_PLACES.add(part)


def is_place(phrase, prep):
    words = phrase.split()
    parts = [p.strip() for p in phrase.split(",")]
    if all(p.lower() in COUNTRIES or p in COUNTRY_MENTION_MAP for p in parts):
        return False  # a country or a list of countries: see countries_mentioned
    if len(words) == 1 and GEO_WORD_RE.fullmatch(phrase):
        return False  # "Channel", "Marina" on their own
    if not words or words[0].lower() in NOT_NAMES or YEAR_RE.search(phrase):
        return False
    if NOT_PLACE_RE.search(phrase) or phrase.lower() in COUNTRIES:
        return False
    if GEO_WORD_RE.search(phrase):
        return True
    if any(p in KNOWN_PLACES or p.lower() in COUNTRIES for p in parts):
        return True
    return prep in {"off", "near", "along"} and len(words) <= 4


def extract_locations(text):
    """Location details stated in a piece of text. Every value is verbatim."""
    places = []
    for m in LOCATION_RE.finditer(text):
        prep, phrase = m.group(1).lower(), m.group(2)
        phrase = re.split(r"\.\s", phrase)[0].strip(" ,.")  # stop at a sentence break
        if is_place(phrase, prep) and phrase not in places:
            places.append(phrase)
    # Drop places that are a substring of a longer one.
    places = [p for p in places if not any(p != q and p in q for q in places)]
    water = [p for p in places if GEO_WORD_RE.search(p)]
    countries = list(dict.fromkeys(COUNTRY_MENTION_MAP.get(m.group(1), m.group(1))
                                   for m in COUNTRY_MENTION_RE.finditer(text)))
    coords = [m.group(0) for m in COORD_RE.finditer(text)]
    depths = [next(g for g in m.groups() if g) + " m" for m in DEPTH_RE.finditer(text)]
    offshore = ["{} {}".format(m.group(1), m.group(2)) for m in OFFSHORE_RE.finditer(text)]
    return {
        "locations_mentioned": "; ".join(places),
        "water_bodies": "; ".join(water),
        "countries_mentioned": "; ".join(countries),
        "coordinates": "; ".join(dict.fromkeys(coords)),
        "water_depth": "; ".join(dict.fromkeys(depths)),
        "distance_offshore": "; ".join(dict.fromkeys(offshore)),
    }


# --------------------------------------------------------------------------------------
# Per-report extraction
# --------------------------------------------------------------------------------------

def year_of(path):
    m = re.search(r"(20\d\d)", path.stem)
    return int(m.group(1)) if m else None


def process_report(path, alias_index, text_dir, project_index=None):
    year = year_of(path)
    lines = clean_text(rtf_to_text(path))
    (text_dir / (path.stem + ".txt")).write_text(
        "\n".join(f"{i + 1:>6}  {ln}" for i, ln in enumerate(lines)), encoding="utf-8")

    chapters = find_chapters(lines)
    country, heading, stype = tag_lines(lines, chapters)
    joined = Joined(lines)

    # ---- Test sites
    mentions = find_site_mentions(joined, alias_index, country)
    site_rows = []
    for a, b, ln, site, matched in mentions:
        site_rows.append({
            "report_year": year, "reporting_country": country[ln] or OUTSIDE,
            "section": heading[ln], "site_id": site["site_id"], "site_name": site["name"],
            "site_country": site["country"], "site_location": site["location"],
            "matched_text": matched, "context": joined.context(a, b), "source_file": path.name + ".txt", "line": ln + 1,
        })
    candidates = find_site_candidates(joined, [(a, b) for a, b, *_ in mentions])

    # ---- Named projects
    named_rows = []
    if project_index:
        for a, b, ln, term, project, how, ids in find_project_mentions(joined, project_index, country):
            named_rows.append({
                "report_year": year, "reporting_country": country[ln] or OUTSIDE,
                "section": heading[ln], "section_type": stype[ln] or "",
                "project_id": project["project_id"] if project else "",
                "project_name": project["name"] if project else "",
                "category": project["category"] if project else "",
                "matched_term": term, "resolution": how,
                "candidates": "; ".join(ids) if len(ids) > 1 else "",
                "context": joined.context(a, b), "source_file": path.name + ".txt", "line": ln + 1,
            })
    cand_rows = [{
        "report_year": year, "candidate": name, "reporting_country": country[ln],
        "context": joined.context(a, b), "line": ln + 1,
    } for a, b, ln, name in candidates]

    # Per-line lookup of which sites are mentioned, for the project rows.
    sites_by_line = defaultdict(set)
    site_by_id = {}
    for a, b, ln, site, _ in mentions:
        site_by_id[site["site_id"]] = site
        for k in range(ln, joined.line_of(b - 1) + 1):
            sites_by_line[k].add(site["site_id"])

    # ---- Projects
    project_rows = []

    def emit(entries, method, section, section_kind, chapter_country, min_score):
        for e in entries:
            text = join_lines([t for _, t in e["lines"]])
            if len(text) < 40:
                continue
            nos = [n for n, _ in e["lines"]]
            sites = sorted(set().union(*(sites_by_line[n] for n in nos)))
            if e["title"] and e["lines"]:
                sites = sorted(set(sites) | sites_by_line[nos[0] - 1])
            score = signal_score(text, sites)
            if score < min_score:
                continue
            if method == "keyword scan" and not (DEPLOY_VERB_RE.search(text)
                                                 and (CAPACITY_RE.search(text) or sites)):
                continue
            caps, techs, status, years = describe(e["title"] + " " + text)
            ents = named_entities(text)
            loc = extract_locations(e["title"] + ". " + text)
            site_objs = [site_by_id[i] for i in sites]
            project_rows.append({
                "report_year": year,
                "reporting_country": chapter_country or OUTSIDE,
                "section": section,
                "section_type": section_kind,
                "entry_title": e["title"],
                "name_guess": next((n for n in [e["title"]] + ents if plausible_name(n)), ""),
                "technology": "; ".join(techs),
                "capacities": "; ".join(dict.fromkeys(caps)),
                "status_keywords": "; ".join(status),
                "years_mentioned": "; ".join(years),
                "test_sites_mentioned": "; ".join(s["name"] for s in site_objs),
                "test_site_locations": "; ".join(
                    "{}, {}".format(s["location"], s["country"]) if s["location"] else s["country"]
                    for s in site_objs),
                **loc,
                "named_entities": "; ".join(ents),
                "signal_score": score,
                "text": text,
                "method": method,
                "source_file": path.name + ".txt",
                "line_start": nos[0] + 1 - (1 if e["title"] else 0),
                "line_end": nos[-1] + 1,
            })

    def blocks(start, end, predicate):
        """Contiguous runs of lines in [start,end) where predicate(k), minus heading lines."""
        run = []
        for k in range(start, end):
            if predicate(k):
                if section_type(lines[k])[0] or "TEST SITE NAME" in lines[k].upper():
                    if run:
                        yield run
                    run = []
                    continue
                run.append(k)
            elif run:
                yield run
                run = []
        if run:
            yield run

    project_kinds = PROJECT_KINDS
    for s, e, c, _ in chapters:
        in_proj = [k for k in range(s, e) if stype[k] in project_kinds]
        if in_proj:
            for run in blocks(s, e, lambda k: stype[k] in project_kinds):
                block = [(k, lines[k]) for k in run]
                emit(split_entries(block), "project section", heading[run[0]],
                     stype[run[0]], c, min_score=1)
        else:
            for run in blocks(s, e, lambda k: True):
                emit(split_entries([(k, lines[k]) for k in run]), "keyword scan",
                     heading[run[0]], stype[run[0]] or "", c, min_score=2)

    if not chapters:  # early reports without recognisable country chapters
        for run in blocks(0, len(lines), lambda k: True):
            emit(split_entries([(k, lines[k]) for k in run]), "keyword scan",
                 heading[run[0]], stype[run[0]] or "", "", min_score=2)

    chapter_rows = [{
        "report_year": year, "reporting_country": c, "author_line": author,
        "line_start": s + 1, "line_end": e, "n_lines": e - s,
        "project_sections_found": "yes" if any(stype[k] in project_kinds for k in range(s, e)) else "no",
    } for s, e, c, author in chapters]

    return site_rows, cand_rows, project_rows, chapter_rows, named_rows, len(lines)


# --------------------------------------------------------------------------------------
# Workbook
# --------------------------------------------------------------------------------------

PROJECT_YEAR_COLUMNS = [
    "project_id", "name", "category", "in_scope", "country", "test_site_id", "report_year",
    "n_mentions", "terms_matched", "reporting_countries", "sections", "lines",
    "in_pilot", "pilot_years", "example_context", "source_file",
]


def summarise_project_years(named_rows, projects):
    """One row per project per report year, from the resolved mentions; plus how the result
    compares with the Pilot's verified report years (2016, 2020, 2025)."""
    by_id = {p["project_id"]: p for p in projects}
    groups = defaultdict(list)
    for r in named_rows:
        if r["project_id"]:
            groups[(r["project_id"], str(r["report_year"]))].append(r)
    rows = []
    for (pid, year), rs in sorted(groups.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        p = by_id[pid]
        pilot = [y.strip() for y in p.get("report_years", "").split(";") if y.strip()] \
            if p.get("report_years_basis", "").startswith("Pilot") else []
        in_pilot = ("yes" if year in pilot else "no") if year in PILOT_YEARS else ""
        best = max(rs, key=lambda r: (r["resolution"] == "unique term", r["section_type"] != ""))
        rows.append({
            "project_id": pid, "name": p["name"], "category": p["category"], "in_scope": p["in_scope"],
            "country": p["country"], "test_site_id": p["test_site_id"], "report_year": year,
            "n_mentions": len(rs),
            "terms_matched": "; ".join(dict.fromkeys(r["matched_term"] for r in rs)),
            "reporting_countries": "; ".join(dict.fromkeys(r["reporting_country"] for r in rs)),
            "sections": "; ".join(dict.fromkeys(r["section"] for r in rs if r["section"])),
            "lines": "; ".join(str(r["line"]) for r in rs[:20]),
            "in_pilot": in_pilot, "pilot_years": "; ".join(pilot),
            "example_context": best["context"], "source_file": rs[0]["source_file"],
        })
    # Pilot recall: of the project-years the Pilot verified, how many were found here.
    found = {(r["project_id"], r["report_year"]) for r in rows}
    pilot_pairs = {(p["project_id"], y.strip()) for p in projects
                   if p.get("report_years_basis", "").startswith("Pilot")
                   for y in p["report_years"].split(";") if y.strip()}
    hit = len(pilot_pairs & found)
    extra = sum(1 for r in rows if r["in_pilot"] == "no")
    check = {
        "Pilot project-years found (recall)": f"{hit} of {len(pilot_pairs)}"
                                              f" ({100 * hit / max(1, len(pilot_pairs)):.0f}%)",
        "Project-years in 2016/2020/2025 not in the Pilot": f"{extra} (Pilot misses or false matches: review)",
    }
    return rows, check


PILOT_YEARS = {"2016", "2020", "2025"}


def summarise_sites(site_rows, sites):
    by_id = defaultdict(list)
    for r in site_rows:
        by_id[r["site_id"]].append(r)
    out = []
    for s in sites:
        rows = by_id.get(s["site_id"], [])
        years = sorted({r["report_year"] for r in rows})
        out.append({
            "site_id": s["site_id"], "site_name": s["name"], "country": s["country"],
            "location": s["location"], "n_mentions": len(rows), "n_reports": len(years),
            "first_year": years[0] if years else "", "last_year": years[-1] if years else "",
            "years_mentioned": ", ".join(map(str, years)),
            "reporting_countries": "; ".join(sorted({r["reporting_country"] for r in rows})),
            "aliases": " | ".join(s["alias_list"]),
        })
    out.sort(key=lambda r: (-r["n_mentions"], r["site_name"]))
    return out


def summarise_candidates(cand_rows):
    groups = defaultdict(list)
    for r in cand_rows:
        groups[r["candidate"].lower()].append(r)
    out = []
    for rows in groups.values():
        names = defaultdict(int)
        for r in rows:
            names[r["candidate"]] += 1
        years = sorted({r["report_year"] for r in rows})
        first = rows[0]
        out.append({
            "candidate": max(names, key=names.get), "n_mentions": len(rows),
            "years_mentioned": ", ".join(map(str, years)),
            "reporting_countries": "; ".join(sorted({r["reporting_country"] for r in rows if r["reporting_country"]})),
            "example_context": first["context"],
            "example_source": f"{first['report_year']} line {first['line']}",
        })
    out.sort(key=lambda r: (-r["n_mentions"], r["candidate"]))
    return out


def summarise_country_years(site_rows, project_rows, chapter_rows, sites):
    """One row per (report year, country): test sites mentioned and project rows found."""
    site_country = {s["site_id"]: s["country"] for s in sites}
    site_name = {s["site_id"]: s["name"] for s in sites}
    cell = defaultdict(lambda: {"sites": set(), "mentions": 0, "projects": 0, "strong": 0, "with_cap": 0,
                                "chapter": False})
    for r in site_rows:
        c = cell[(r["report_year"], site_country[r["site_id"]])]
        c["sites"].add(r["site_id"])
        c["mentions"] += 1
    for r in project_rows:
        if r["reporting_country"] == OUTSIDE:
            continue
        c = cell[(r["report_year"], r["reporting_country"])]
        c["projects"] += 1
        c["strong"] += r["signal_score"] >= 2
        c["with_cap"] += bool(r["capacities"])
    for r in chapter_rows:
        cell[(r["report_year"], r["reporting_country"])]["chapter"] = True
    return [{
        "report_year": year,
        "country": country,
        "country_chapter_found": "yes" if c["chapter"] else "no",
        "n_test_sites": len(c["sites"]),
        "n_test_site_mentions": c["mentions"],
        "n_projects": c["projects"],
        "n_projects_signal_2plus": c["strong"],
        "n_projects_with_capacity": c["with_cap"],
        "n_projects_without_capacity": c["projects"] - c["with_cap"],
        "test_sites": "; ".join(sorted(site_name[i] for i in c["sites"])),
    } for (year, country), c in sorted(cell.items())]


def write_sheet(wb, title, rows, columns, widths=None, wrap=()):
    ws = wb.create_sheet(title)
    ws.append(columns)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for r in rows:
        vals = []
        for col in columns:
            v = r.get(col, "")
            if isinstance(v, str):
                v = ILLEGAL_CHARACTERS_RE.sub("", v)[:32000]
            vals.append(v)
        ws.append(vals)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = widths or {}
    for i, col in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 16)
        if col in wrap:
            for cell in ws.iter_cols(min_col=i, max_col=i, min_row=2):
                for c in cell:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def write_readme(wb, stats):
    ws = wb.active
    ws.title = "README"
    text = [
        ("IEA-OES Annual Reports: test sites and projects", True),
        ("Generated by extract.py from the RTF versions of the annual reports. "
         "Local, deterministic parsing only: every value is a pattern match on the report text, "
         "not an interpretation.", False),
        ("", False),
        ("Sheets", True),
        ("Project tabs: one row per project entry. Inside project-type sections (TECHNOLOGY "
         "DEMONSTRATION, PROJECTS IN THE WATER, PLANNED DEPLOYMENTS, OPERATIONAL PROJECTS, "
         "R&D PROJECTS...) each titled entry, bullet or paragraph is a row (method = 'project "
         "section'). Chapters without such headings, mostly before 2013, fall back to paragraphs "
         "that mention a deployment verb together with a capacity or a test site (method = "
         "'keyword scan'). The same project appears once per report year; rows are not merged "
         "across years.", False),
        ("Projects with capacity / Projects without capacity: the project entries found in country "
         "chapters, split by whether the entry text states a capacity (a number with kW, MW, GW, "
         "kWh, MWh etc.; see the capacities column). Both tabs have the same columns. A capacity "
         "in the text may belong to a programme or target rather than the device, and a project "
         "without one may still have a capacity stated in another year or entry.", False),
        ("Projects outside chapters: the same columns as the two Projects tabs, for entries found outside any "
         "detected country chapter: executive summaries, overview sections, and the early reports "
         "(mostly 2002-2007) where no country chapters are detected. These rows have no "
         "reporting country; use the location columns.", False),
        ("Counts by country and year: one row per report year and country. n_test_sites = "
         "distinct gazetteer test sites located in that country (by the gazetteer's country) that "
         "the report mentions anywhere; n_test_site_mentions = total mentions of them. n_projects = "
         "rows on the Projects sheet from that country's chapter; n_projects_signal_2plus = those "
         "with signal_score >= 2. Projects outside country chapters are not counted here. "
         "Project counts are extracted entries, not unique projects: a project described in several "
         "entries, or in several years, is counted each time. A 0 in n_projects with "
         "country_chapter_found = no means no chapter was detected, not that there were no projects.", False),
        ("Project years: one row per named project (gazetteers/projects.csv) per report year in which "
         "it is mentioned, from the Project mentions sheet. n_mentions, the terms matched, the chapters "
         "and sections they sit in, and line numbers in output/text/. in_pilot compares with the "
         "Pilot's verified years for 2016, 2020 and 2025: 'yes' = the Pilot has it too, 'no' = found "
         "here only (a Pilot miss or a false match); blank in other years. A mention can be the "
         "developer or device rather than the deployment itself: read example_context.", False),
        ("Project mentions: every match of a project search term, with how it was resolved: "
         "'unique term'; 'context: ...' (the cues that picked the project for a shared or weak term); "
         "'ambiguous' (a shared term the context could not settle - candidates lists the projects); "
         "'weak term, no context' (an organisation or place name with nothing else about the project "
         "nearby). Only resolved mentions count in Project years.", False),
        ("Test site mentions: every match of a site in test_sites.csv, with context.", False),
        ("Test sites: one row per gazetteer site, with the years in which it is mentioned.", False),
        ("Test site candidates: names that look like test sites (e.g. '... Test Site', "
         "'... Marine Energy Centre') but are not in the gazetteer. Review these and add real "
         "sites to the gazetteer, then re-run.", False),
        ("Chapters: the country chapters detected in each report, for checking coverage.", False),
        ("", False),
        ("Reading the columns", True),
        ("capacities, technology, status_keywords, years_mentioned: every value found anywhere in "
         "the entry text. An entry mentioning '500 kW' and 'planned' is not necessarily a planned "
         "500 kW device; read the text column.", False),
        ("name_guess: the entry's title line where the report gives one; otherwise the first "
         "capitalised name in the text, which is only a guess.", False),
        ("signal_score (0-4): one point each for a capacity, a deployment verb, a device word and "
         "a test site. Filter to >= 2 for the rows most likely to describe a physical project.", False),
        ("line / line_start: line number in the cleaned text file in the text/ folder next to the "
         "workbook, for tracing any row back to the report.", False),
        ("report_year is the year the annual report covers (from the file name), on every row. "
         "years_mentioned is different: the years that appear in the entry's text.", False),
        ("reporting_country is the country whose chapter the text sits in, not necessarily the "
         "country where the project is: 2016 Belgium, for example, reports work at EMEC in "
         "Scotland. For where the project is, use the location columns.", False),
        ("Location columns (Projects sheet), all verbatim from the entry text: "
         "locations_mentioned = place phrases following in/at/off/near/into etc. that contain a "
         "geographic word (River, Bay, Strait, Island...) or a known region or place; "
         "water_bodies = the subset with a geographic word; countries_mentioned = country names "
         "in the text; test_sites_mentioned / test_site_locations = gazetteer sites named in the "
         "entry and their location from the gazetteer; coordinates, water_depth and "
         "distance_offshore = values stated in the text. Nothing is geocoded or inferred.", False),
        ("Test site locations come from test_sites_gazetteer.csv (location and country columns), "
         "which were taken from the reports' own test-site tables. Edit that file to correct them.", False),
        ("", False),
        ("Run summary", True),
    ] + [(f"{k}: {v}", False) for k, v in stats.items()]
    for i, (t, bold) in enumerate(text, 1):
        c = ws.cell(row=i, column=1, value=t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if bold:
            c.font = Font(bold=True, size=12)
    ws.column_dimensions["A"].width = 120


# --------------------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="folder of annual report .rtf files")
    ap.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Excel workbook to write")
    ap.add_argument("--gazetteer", type=Path, default=DEFAULT_GAZETTEER, help="test site list (CSV)")
    ap.add_argument("--projects", type=Path, default=DEFAULT_PROJECTS, help="project list (CSV)")
    args = ap.parse_args()

    files = sorted(args.input.glob("*.rtf"))
    if not files:
        sys.exit(f"No .rtf files in {args.input}")
    sites = load_gazetteer(args.gazetteer)
    alias_index = build_alias_index(sites)
    register_places(sites)
    projects = load_projects(args.projects) if args.projects.exists() else []
    project_index = build_project_index(projects, sites) if projects else None
    text_dir = args.output.parent / "text"
    text_dir.mkdir(parents=True, exist_ok=True)

    all_sites, all_cands, all_projects, all_chapters, all_named = [], [], [], [], []
    for f in files:
        s, c, p, ch, nm, n = process_report(f, alias_index, text_dir, project_index)
        all_named += nm
        print(f"{f.name}: {n} lines, {len(ch)} chapters, {len(p)} project rows, {len(s)} site mentions")
        all_sites += s
        all_cands += c
        all_projects += p
        all_chapters += ch

    project_years, pilot_check = summarise_project_years(all_named, projects)
    site_summary = summarise_sites(all_sites, sites)
    cand_summary = summarise_candidates(all_cands)
    outside = [r for r in all_projects if r["reporting_country"] == OUTSIDE]
    in_chapters = [r for r in all_projects if r["reporting_country"] != OUTSIDE]
    stats = {
        "Reports": f"{len(files)} ({year_of(files[0])}-{year_of(files[-1])})",
        "Country chapters detected": len(all_chapters),
        "Project rows": f"{len(all_projects)} ({len(in_chapters)} in country chapters, "
                        f"{len(outside)} outside them)",
        "Project rows in chapters with a capacity": sum(bool(r["capacities"]) for r in in_chapters),
        "Project rows with signal_score >= 2": sum(r["signal_score"] >= 2 for r in all_projects),
        "Test site mentions": len(all_sites),
        "Gazetteer sites mentioned at least once": f"{sum(r['n_mentions'] > 0 for r in site_summary)} of {len(sites)}",
        "Unmatched test-site candidates": len(cand_summary),
        "Named project mentions": f"{len(all_named)} ({sum(bool(r['project_id']) for r in all_named)} resolved "
                                  f"to one project, {sum(not r['project_id'] for r in all_named)} ambiguous)",
        "Project-years found": f"{len(project_years)} rows, {len({r['project_id'] for r in project_years})} "
                               f"of {len(projects)} projects",
        **pilot_check,
    }

    wb = Workbook()
    write_readme(wb, stats)
    project_columns = [
        "report_year", "reporting_country", "section", "section_type", "entry_title", "name_guess",
        "technology", "capacities", "status_keywords", "years_mentioned",
        "locations_mentioned", "water_bodies", "countries_mentioned", "test_sites_mentioned",
        "test_site_locations", "coordinates", "water_depth", "distance_offshore",
        "signal_score", "text", "named_entities", "method", "source_file", "line_start", "line_end",
    ]
    project_widths = {"reporting_country": 18, "section": 24, "entry_title": 28, "name_guess": 28,
               "technology": 18, "status_keywords": 24, "locations_mentioned": 36, "water_bodies": 28,
               "countries_mentioned": 22, "test_sites_mentioned": 30, "test_site_locations": 34,
               "text": 90, "named_entities": 40, "source_file": 30}
    write_sheet(wb, "Projects with capacity", [r for r in in_chapters if r["capacities"]],
                project_columns, widths=project_widths, wrap=("text",))
    write_sheet(wb, "Projects without capacity", [r for r in in_chapters if not r["capacities"]],
                project_columns, widths=project_widths, wrap=("text",))
    write_sheet(wb, "Projects outside chapters", outside, project_columns, widths=project_widths,
                wrap=("text",))
    write_sheet(wb, "Counts by country and year",
                summarise_country_years(all_sites, all_projects, all_chapters, sites), [
        "report_year", "country", "country_chapter_found", "n_test_sites", "n_test_site_mentions",
        "n_projects", "n_projects_signal_2plus", "n_projects_with_capacity",
        "n_projects_without_capacity", "test_sites",
    ], widths={"country": 22, "country_chapter_found": 12, "test_sites": 90})
    write_sheet(wb, "Project years", project_years, PROJECT_YEAR_COLUMNS,
                widths={"name": 45, "reporting_countries": 24, "sections": 40, "terms_matched": 30,
                        "lines": 30, "example_context": 90, "pilot_years": 16}, wrap=("example_context",))
    write_sheet(wb, "Project mentions", all_named, [
        "report_year", "reporting_country", "section", "section_type", "project_id", "project_name",
        "category", "matched_term", "resolution", "candidates", "context", "source_file", "line",
    ], widths={"reporting_country": 18, "section": 24, "project_name": 40, "matched_term": 22,
               "resolution": 30, "candidates": 24, "context": 90, "source_file": 30}, wrap=("context",))
    write_sheet(wb, "Test sites", site_summary, [
        "site_id", "site_name", "country", "location", "n_mentions", "n_reports", "first_year",
        "last_year", "years_mentioned", "reporting_countries", "aliases",
    ], widths={"site_name": 45, "location": 30, "years_mentioned": 50, "reporting_countries": 40, "aliases": 60})
    write_sheet(wb, "Test site mentions", all_sites, [
        "report_year", "reporting_country", "section", "site_id", "site_name", "site_country",
        "site_location", "matched_text", "context", "source_file", "line",
    ], widths={"reporting_country": 18, "site_name": 40, "site_location": 30, "section": 24, "matched_text": 26, "context": 90, "source_file": 30},
        wrap=("context",))
    write_sheet(wb, "Test site candidates", cand_summary, [
        "candidate", "n_mentions", "years_mentioned", "reporting_countries", "example_context", "example_source",
    ], widths={"candidate": 45, "years_mentioned": 40, "reporting_countries": 30, "example_context": 90},
        wrap=("example_context",))
    write_sheet(wb, "Chapters", all_chapters, [
        "report_year", "reporting_country", "author_line", "line_start", "line_end", "n_lines", "project_sections_found",
    ], widths={"reporting_country": 22, "author_line": 70})

    args.output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)
    with open(args.output.parent / "project_years.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=PROJECT_YEAR_COLUMNS)
        w.writeheader()
        w.writerows(project_years)
    print()
    for k, v in stats.items():
        print(f"{k}: {v}")
    print(f"\nWrote {args.output}")


if __name__ == "__main__":
    main()
