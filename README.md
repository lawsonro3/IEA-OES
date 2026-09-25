# IEA-OES Test Sites and Projects

A reviewed list of ocean energy **test sites** and **pilot / demonstration projects** reported in
the IEA-OES annual reports (2002–2025), built by deterministic parsing and validated against the
original PDFs and the PRIMRE databases.

This folder builds on two earlier pieces of work, which are kept unchanged as references:

| Folder | What it contributes |
|---|---|
| `../Annual Report Extraction` | The parser (copied here as `extract.py`) and the 70-site test-site gazetteer. Its workbook is the baseline to compare against. |
| `../Extraction Pilot` | LLM-extracted projects for 2016/2020/2025, the PRIMRE matching script and completed review verdicts. |
| `../PRIMRE databases` | Independent project and test-site records (US DOE / PNNL). |
| `../OES Publications/Annual Reports` | The source reports: RTFs (parsed) and PDFs (for validation). |

## Layout

```
extract.py                  parse the RTF reports -> output/ workbook (+ output/text/)
gazetteers/
  test_sites.csv            THE test-site list. Hand-edited; extract.py searches the reports for it
  test_site_candidates.csv  Pilot site names that matched no gazetteer site — review queue
  projects.csv              THE project list (seed). Hand-edited
scripts/
  seed_gazetteers.py        one-off: built the gazetteers from the sources above
  propose_search_terms.py   proposes project search terms; fills empty cells only
validation/                 PDF and PRIMRE checks, review CSVs (to come)
output/                     generated — safe to delete, not version-controlled
```

## Running

```bash
python3 extract.py
```

Needs macOS (`textutil`) and `openpyxl`. Takes about 40 seconds for all 24 reports.

## The gazetteers are the source of truth

`gazetteers/*.csv` were seeded once by `scripts/seed_gazetteers.py` and are **edited by hand from
now on**. The seed script refuses to overwrite them without `--force`; don't use `--force` after
review has started, or the edits are lost.

- **test_sites.csv** keeps the Annual Report Extraction columns (`site_id … aliases`) so
  `extract.py` reads it unchanged, plus `site_type` and provenance: `source`, the matching
  `primre_*` record, the Pilot's names for the site (`pilot_names`), and `review_status`.
  - **Laboratories and tanks count as test sites.** `site_type` is one of:
    `open-water` (sea, estuary, river or natural lab in the sea), `laboratory` (tank, basin,
    flume, blade or grid test facility), `shore-based` (plant on the shore: OTEC, breakwater,
    dike), `mixed` (a centre running several of these, e.g. PMEC, HINMREC).
  - `aliases` are matched **case-sensitively** as whole words. Keep them distinctive: a short
    acronym shared with another site produces false matches (the US "AMEC" vs NAGASAKI-AMEC).
- **projects.csv** is one row per project: Pilot records (duplicates merged) and PRIMRE
  projects. `notes` records every merge and link change, and `primre_link_basis` says how each
  PRIMRE link was made; a project may carry several PRIMRE records (its phases).
  `in_scope = no` marks non-ocean-energy items (floating solar, seawater heat pumps …).
  `category` is `annual reports` (named in a report) or `PRIMRE` (PRIMRE only).
  - **Which report each project came from:** `report_years`, `report_pages` (`2016:p45; 2025:p130`)
    and `obs_ids` (the Pilot's observations, each with a quoted passage), with
    `report_years_basis` saying where they came from: *Pilot observations* (verified, with page
    and quote) or *name search* (a term found in that year's text, unverified). `first_seen` /
    `last_seen` are the range only; a project can be missing from a year in between.
  - The same links for test sites are in `test_sites.csv`: `pilot_record_ids`,
    `pilot_report_years`, `pilot_report_pages`, `pilot_obs_ids`.
  - All 594 Pilot observations and all 474 Pilot records are held in one of the two lists.
  - `search_terms` are what `extract.py` searches for: short names, case-sensitive, whole
    words, like test-site aliases. A leading `~` marks a weak term (an organisation, a place,
    a short acronym, a device name PRIMRE uses for several deployments): it only counts when
    the passage also names something else about the project. `scripts/propose_search_terms.py`
    proposes them and fills only empty cells, so hand edits survive a re-run.
    `validation/search_terms_review.csv` lists every candidate with its hits and flags:
    *shared* (a developer or device name used by several projects: needs context to
    resolve), *place name*, *frequent*.

## Status

- [x] Folder set up; `extract.py` runs on the new test-site list and reproduces the baseline
- [x] Labs/tanks in scope: `site_type` added; 35 Pilot candidates decided (26 added, 7 merged,
      2 dropped — see `test_site_candidates.csv`); CCOB added; `extract.py` candidate search
      now also flags lab/tank/basin/flume names
- [x] Review the 8 PRIMRE additions: 5 kept (AMEC, Sequim Bay, ZJU Zhairuoshan, QML Strangford,
      Cal Poly Pier), 2 merged (Zhuhai, HOST Park), Thames moved to
      `validation/primre_only_test_sites.csv`; Clallam Bay added from the reports
- [ ] Review the parser's candidate sheet (418 names, about half lab-type, many are institutions)
- [x] Projects, steps 1–3: unsure PRIMRE links and shared records resolved; missed Pilot–PRIMRE
      links added and Pilot duplicates merged (`validation/project_review_step2.csv`);
      `category` column: `annual reports` (named in a report) or `PRIMRE` (PRIMRE only)
- [x] Projects, step 4: `search_terms` column proposed by `scripts/propose_search_terms.py`
      and checked against the reports (`validation/search_terms_review.csv`); 17 more
      test-site links added from the Pilot's site names
- [x] `extract.py`: named projects found by their search terms across all 24 reports ->
      `Project years` / `Project mentions` sheets and `output/project_years.csv`.
      Recall against the Pilot's verified project-years 78%; precision in a 40-row sample of
      other years 32 right, 5 partly (developer named, not the deployment), 3 wrong
      (`validation/project_years_sample_check.csv`)
- [ ] Validation against the PDFs (page matching, 2016 annex, audit sample)
- [ ] Validation against PRIMRE
