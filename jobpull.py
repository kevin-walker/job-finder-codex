#!/usr/bin/env python3
"""
jobpull.py — Fresh jobs straight from company career pages (Greenhouse / Lever /
              Ashby public APIs). No scraping, no API keys, no recruiter spam.

THREE MODES
-----------
1. Discovery (one-time-ish): scan the seed registry concurrently, keep only
   companies that currently have an India/Bangalore posting, cache them.
       python jobpull.py --discover

2. Scan (your daily driver): pull jobs from your active company list, filter by
   keyword + recency, dedupe, export to Excel.
       python jobpull.py

3. Verify: ping every token in the registry, report which are live.
       python jobpull.py --verify

SETUP
-----
    pip install requests python-dateutil openpyxl

Then run discovery once to build india_companies.json, then just `python jobpull.py`
whenever you want a fresh pull. Re-run --discover every week or two to catch new
companies (and prune ones that stopped hiring in India).

The seed company list lives in registry.py — add tokens there to widen coverage.
"""

import sys
import json
import os
import time
import datetime as dt
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import truststore
    truststore.inject_into_ssl()  # use macOS Keychain certs
except ImportError:
    pass

import certifi
import requests
from dateutil import parser as dateparser
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ───────────────────────────────────────────────────────────────────────────
# CONFIG
# ───────────────────────────────────────────────────────────────────────────

# Target role keywords — at least one must appear in the job title.
# Empty [] = keep every title.
KEYWORDS = [
    "devops",
    "technical support",
    "support engineer",
    "cloud support",
    "cloud engineer",
    "technical solutions",
    "solutions engineer",
    "technical consultant",
    "sre",
    "site reliability",
    "platform engineer",
    "infrastructure engineer",
    "infra support",
    "infrastructure support",
    "reliability engineer",
]

# Titles containing any of these are excluded (management / executive only).
# Senior DevOps / Senior Support / Lead roles are intentionally kept.
EXCLUDE_SENIORITY = [
    "manager", "director",
    "head of", "head,",
    "vp ", "vice president", "architect",
]

# India detection — a posting counts as "India" if its location contains any of these.
INDIA_TERMS = [
    "india", "bangalore", "bengaluru",
    "kochi", "cochin", "trivandrum", "thiruvananthapuram",
    "hyderabad", "pune", "gurgaon", "gurugram",
    "noida", "mumbai", "delhi", "chennai",
    "remote - india", "remote, india", "remote (india)", "remote india",
    "india remote",
]

# Remote APAC and Remote US are also included.
APAC_REMOTE_TERMS = [
    "remote - apac", "remote, apac", "remote (apac)", "remote apac",
    "apac remote", "asia pacific", "asia-pacific",
    "remote - asia", "remote, asia",
]

US_REMOTE_TERMS = [
    "remote - us", "remote - usa", "remote, us", "remote, usa",
    "remote (us)", "remote (usa)", "united states - remote",
    "remote - united states", "us remote", "usa remote",
]

# Recency window for the scan, in hours. 336 = 14 days.
HOURS_BACK = 336

# Cap results in the Excel output (after newest-first sort).
MAX_RESULTS = 400

# Concurrency — increased for faster full-registry scans.
MAX_WORKERS = 15
REQUEST_TIMEOUT = 20

# Files
CACHE_FILE = "india_companies.json"     # written by --discover, read by scan
OUTPUT_PREFIX = "jobs"

HEADERS = {"User-Agent": "jobpull/2.0 (personal job search; contact: you@example.com)"}
_SSL = certifi.where()  # fixes SSL cert errors on macOS/Homebrew Python

# ───────────────────────────────────────────────────────────────────────────
# Load seed registry
# ───────────────────────────────────────────────────────────────────────────
try:
    from registry import SEED
except ImportError:
    print("ERROR: registry.py not found in the same folder. It holds the company list.")
    sys.exit(1)


_IST = dt.timezone(dt.timedelta(hours=5, minutes=30))


def _now_utc():
    return dt.datetime.now(dt.timezone.utc)


# ───────────────────────────────────────────────────────────────────────────
# ATS fetchers — each returns list of normalized dicts:
#   {role, company, location, posted_dt|None, source, link, age_known}
# ───────────────────────────────────────────────────────────────────────────

import re as _re

def _parse_dt(s):
    if not s:
        return None
    try:
        return dateparser.parse(s).astimezone(dt.timezone.utc)
    except Exception:
        return None


def _parse_workday_date(text):
    """Convert Workday 'Posted X Days Ago' text to UTC datetime."""
    if not text:
        return None
    t = text.lower()
    now = _now_utc()
    if "today" in t:
        return now.replace(hour=0, minute=0, second=0, microsecond=0)
    if "yesterday" in t:
        return (now - dt.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    m = _re.search(r"(\d+)\+?\s*days?\s*ago", t)
    if m:
        return (now - dt.timedelta(days=int(m.group(1)))).replace(hour=0, minute=0, second=0, microsecond=0)
    if "month" in t:
        return (now - dt.timedelta(days=31)).replace(hour=0, minute=0, second=0, microsecond=0)
    return None


def fetch_greenhouse(name, token):
    url = f"https://boards-api.greenhouse.io/v1/boards/{token}/jobs?content=false"
    out = []
    r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, verify=_SSL)
    r.raise_for_status()
    for j in r.json().get("jobs", []):
        posted = _parse_dt(j.get("updated_at") or j.get("first_published"))
        out.append({
            "role": (j.get("title") or "").strip(),
            "company": name,
            "location": (j.get("location") or {}).get("name", "") or "",
            "posted_dt": posted,
            "source": "Greenhouse",
            "link": j.get("absolute_url", ""),
            "age_known": posted is not None,
        })
    return out


def fetch_lever(name, token):
    url = f"https://api.lever.co/v0/postings/{token}?mode=json"
    out = []
    r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, verify=_SSL)
    r.raise_for_status()
    for j in r.json():
        created = j.get("createdAt")
        posted = None
        if created:
            try:
                posted = dt.datetime.fromtimestamp(created / 1000, tz=dt.timezone.utc)
            except Exception:
                posted = None
        cats = j.get("categories", {}) or {}
        out.append({
            "role": (j.get("text") or "").strip(),
            "company": name,
            "location": cats.get("location", "") or "",
            "posted_dt": posted,
            "source": "Lever",
            "link": j.get("hostedUrl", "") or j.get("applyUrl", ""),
            "age_known": posted is not None,
        })
    return out


def fetch_ashby(name, token):
    url = f"https://api.ashbyhq.com/posting-api/job-board/{token}"
    out = []
    r = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, verify=_SSL,
                     params={"includeCompensation": "true"})
    r.raise_for_status()
    for j in r.json().get("jobs", []):
        posted = _parse_dt(j.get("publishedAt"))
        out.append({
            "role": (j.get("title") or "").strip(),
            "company": name,
            "location": j.get("location", "") or "",
            "posted_dt": posted,
            "source": "Ashby",
            "link": j.get("jobUrl", "") or j.get("applyUrl", ""),
            "age_known": posted is not None,
        })
    return out


_WORKDAY_INDIA_PATH_TERMS = [
    "india", "bangalore", "bengaluru", "hyderabad", "mumbai",
    "chennai", "pune", "delhi", "noida", "gurgaon", "gurugram",
    "kolkata", "kochi", "cochin", "trivandrum", "thiruvananthapuram",
]
_WORKDAY_APAC_PATH_TERMS = [
    "singapore", "australia", "japan", "south-korea", "indonesia",
    "malaysia", "philippines", "thailand", "vietnam", "hong-kong",
    "taiwan", "new-zealand", "bangladesh", "sri-lanka",
]


def _workday_path_location(ext_path):
    """Extract a readable location string from a Workday externalPath.

    '/job/India---Bangalore/Engineer_R123'  →  'India - Bangalore'
    '/job/USA---Sunnyvale-CA/...'            →  'USA - Sunnyvale, CA'
    """
    try:
        seg = ext_path.split("/")[2]      # e.g. 'India---Bangalore'
        seg = seg.replace("---", " - ").replace("--", ", ").replace("-", " ")
        return seg.strip()
    except (IndexError, AttributeError):
        return ""


def _workday_path_is_india_apac(ext_path):
    p = ext_path.lower()
    return any(t in p for t in _WORKDAY_INDIA_PATH_TERMS) or \
           any(t in p for t in _WORKDAY_APAC_PATH_TERMS)


def fetch_workday(name, token):
    """Workday public job board. token = 'tenant/wdX/career_path' e.g. 'crowdstrike/wd5/crowdstrikecareers'.

    Fetches all jobs then keeps only India/APAC ones by parsing the externalPath
    (e.g. /job/India---Bangalore/...).  Avoids the session-cookie dependency that
    the appliedFacets filter requires.
    """
    parts = token.split("/")
    if len(parts) != 3:
        raise ValueError(f"Workday token must be 'tenant/wdX/career_path', got: {token}")
    tenant, wdv, career_path = parts
    api_url = f"https://{tenant}.{wdv}.myworkdayjobs.com/wday/cxs/{tenant}/{career_path}/jobs"

    out, offset, limit = [], 0, 20
    while True:
        r = requests.post(api_url,
                          json={"limit": limit, "offset": offset, "searchText": ""},
                          headers=HEADERS, timeout=REQUEST_TIMEOUT, verify=_SSL)
        r.raise_for_status()
        data = r.json()
        postings = data.get("jobPostings") or []
        total = data.get("total") or 0
        for j in postings:
            ext_path = j.get("externalPath") or ""
            if not _workday_path_is_india_apac(ext_path):
                continue
            posted_dt = _parse_workday_date(j.get("postedOn") or "")
            loc_text = (j.get("locationsText") or "").strip()
            # locationsText is often "2 Locations" — derive real name from path instead
            if not loc_text or loc_text[0].isdigit():
                loc_text = _workday_path_location(ext_path)
            link = f"https://{tenant}.{wdv}.myworkdayjobs.com/en-US/{career_path}{ext_path}"
            out.append({
                "role": (j.get("title") or "").strip(),
                "company": name,
                "location": loc_text,
                "posted_dt": posted_dt,
                "source": "Workday",
                "link": link,
                "age_known": posted_dt is not None,
            })
        if offset + limit >= total or not postings:
            break
        offset += limit
    return out


FETCHERS = {
    "greenhouse": fetch_greenhouse,
    "lever": fetch_lever,
    "ashby": fetch_ashby,
    "workday": fetch_workday,
}


def safe_fetch(entry):
    """Wrap a fetch so one bad company never kills the run. Retries on 429/502/503."""
    name, ats, token = entry
    fetcher = FETCHERS.get(ats)
    if not fetcher:
        return entry, [], f"unknown ATS '{ats}'"
    for attempt in range(3):
        try:
            return entry, fetcher(name, token), None
        except requests.exceptions.HTTPError as e:
            code = e.response.status_code if e.response is not None else 0
            if code in (429, 502, 503) and attempt < 2:
                time.sleep(1.5 * (attempt + 1))
                continue
            return entry, [], str(e)
        except Exception as e:
            return entry, [], str(e)
    return entry, [], "max retries exceeded"


def fetch_all_concurrent(companies, label="Fetching"):
    """Fetch every company in parallel. Returns (all_jobs, results_summary)."""
    all_jobs, summary = [], []
    total = len(companies)
    done = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(safe_fetch, c): c for c in companies}
        for fut in as_completed(futures):
            entry, jobs, err = fut.result()
            done += 1
            name = entry[0]
            if err:
                summary.append((name, entry[1], 0, err))
            else:
                summary.append((name, entry[1], len(jobs), None))
                all_jobs.extend(jobs)
            # lightweight progress
            print(f"\r  {label}: {done}/{total} companies...", end="", flush=True)
    print()
    return all_jobs, summary


# ───────────────────────────────────────────────────────────────────────────
# Filters
# ───────────────────────────────────────────────────────────────────────────

def _contains_any(text, terms):
    t = (text or "").lower()
    return any(term in t for term in terms)


def matches_keywords(role):
    if not KEYWORDS:
        return True
    role_lower = (role or "").lower()
    if not any(k in role_lower for k in KEYWORDS):
        return False
    return not any(s in role_lower for s in EXCLUDE_SENIORITY)


def is_india(location):
    return _contains_any(location, INDIA_TERMS)


def location_ok_for_scan(location):
    if is_india(location):
        return True
    if _contains_any(location, APAC_REMOTE_TERMS):
        return True
    return False


def within_window(job, cutoff):
    if not job["age_known"] or job["posted_dt"] is None:
        return True  # keep but it'll be flagged "age n/a"
    return job["posted_dt"] >= cutoff


def dedupe(jobs):
    seen, out = set(), []
    for j in jobs:
        key = (j["company"].lower(), j["role"].lower(), (j["location"] or "").lower())
        if key not in seen:
            seen.add(key)
            out.append(j)
    return out


# ───────────────────────────────────────────────────────────────────────────
# Excel
# ───────────────────────────────────────────────────────────────────────────

def write_excel(jobs, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Fresh Jobs"
    headers = ["Role", "Company", "Location", "Posted", "Age", "Source", "Link"]
    ws.append(headers)

    head_fill = PatternFill("solid", fgColor="0D1117")
    head_font = Font(color="2DD4BF", bold=True, size=11)
    thin = Side(style="thin", color="30363D")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.fill, c.font, c.border = head_fill, head_font, border
        c.alignment = Alignment(horizontal="left", vertical="center")

    now = _now_utc()
    for j in jobs:
        if j["posted_dt"]:
            ist_dt = j["posted_dt"].astimezone(_IST)
            posted_str = ist_dt.strftime("%d %b %Y  %I:%M %p IST")
            hrs = (now - j["posted_dt"]).total_seconds() / 3600
            if hrs < 1:
                age_str = f"{int(hrs * 60)}m ago"
            elif hrs < 24:
                age_str = f"{hrs:.0f}h ago"
            elif hrs < 48:
                age_str = "1d ago"
            else:
                age_str = f"{hrs / 24:.0f}d ago"
        else:
            posted_str, age_str = "unknown", "age n/a"
        ws.append([j["role"], j["company"], j["location"],
                   posted_str, age_str, j["source"], j["link"]])

    for row in range(2, ws.max_row + 1):
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=(col in (1, 3)))
        link_cell = ws.cell(row=row, column=7)
        if link_cell.value:
            link_cell.hyperlink = link_cell.value
            link_cell.value = "Open"
            link_cell.font = Font(color="0969DA", underline="single")

    for i, w in enumerate([42, 20, 26, 20, 11, 11, 9], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


# ───────────────────────────────────────────────────────────────────────────
# Modes
# ───────────────────────────────────────────────────────────────────────────

def mode_verify():
    print(f"Verifying {len(SEED)} tokens (concurrent)...\n")
    _, summary = fetch_all_concurrent(SEED, label="Verifying")
    summary.sort(key=lambda s: (s[3] is None, -s[2]))  # live first, by count
    live = [s for s in summary if s[3] is None and s[2] > 0]
    dead = [s for s in summary if s[3] is not None or s[2] == 0]
    print("\nLIVE:")
    for name, ats, n, _ in sorted(live, key=lambda x: -x[2]):
        print(f"  OK {name:24s} {ats:11s} {n}")
    if dead:
        print("\nNEEDS ATTENTION (wrong token / changed ATS / no open roles):")
        for name, ats, n, err in dead:
            reason = err if err else "0 postings"
            print(f"  XX {name:24s} {ats:11s} {reason[:50]}")
    print(f"\n  {len(live)} live, {len(dead)} need attention.")


def mode_discover():
    print(f"Discovery: scanning {len(SEED)} companies for India/Bangalore roles...\n")
    all_jobs, summary = fetch_all_concurrent(SEED, label="Discovering")

    # Which companies have at least one India posting?
    india_companies = {}
    for j in all_jobs:
        if is_india(j["location"]):
            key = j["company"]
            india_companies.setdefault(key, {"count": 0})
            india_companies[key]["count"] += 1

    # Map company name -> its registry entry, keep those with India roles
    name_to_entry = {name: (name, ats, tok) for name, ats, tok in SEED}
    kept = []
    for name, info in india_companies.items():
        if name in name_to_entry:
            entry = name_to_entry[name]
            kept.append({"name": entry[0], "ats": entry[1], "token": entry[2],
                         "india_roles_seen": info["count"]})

    kept.sort(key=lambda x: -x["india_roles_seen"])

    with open(CACHE_FILE, "w") as f:
        json.dump({"generated": _now_utc().isoformat(), "companies": kept}, f, indent=2)

    failed = [s for s in summary if s[3] is not None]
    print(f"\n  {len(kept)} companies are hiring in India right now.")
    print(f"  ({len(failed)} companies failed to fetch — likely wrong tokens.)")
    print(f"  Saved -> {CACHE_FILE}\n")
    print("  India-hiring companies (by # India roles):")
    for c in kept[:30]:
        print(f"    {c['india_roles_seen']:3d}  {c['name']} ({c['ats']})")
    if len(kept) > 30:
        print(f"    ... and {len(kept) - 30} more")
    print(f"\n  Now run:  python jobpull.py   (uses this cached list)")


def load_active_companies():
    """Always scan every company in the registry."""
    ats_counts = {}
    for _, ats, _ in SEED:
        ats_counts[ats] = ats_counts.get(ats, 0) + 1
    summary = ", ".join(f"{v} {k}" for k, v in sorted(ats_counts.items()))
    print(f"  Scanning all {len(SEED)} companies ({summary}).")
    return list(SEED)


def mode_scan():
    print(f"jobpull scan — last {HOURS_BACK}h, keep {MAX_RESULTS}\n")
    companies = load_active_companies()
    all_jobs, summary = fetch_all_concurrent(companies, label="Scanning")
    print(f"\n  Collected {len(all_jobs)} raw postings. Filtering...")

    cutoff = _now_utc() - dt.timedelta(hours=HOURS_BACK)
    filtered = [
        j for j in all_jobs
        if matches_keywords(j["role"])
        and location_ok_for_scan(j["location"])
        and within_window(j, cutoff)
    ]
    filtered = dedupe(filtered)
    filtered.sort(
        key=lambda j: j["posted_dt"] or dt.datetime.min.replace(tzinfo=dt.timezone.utc),
        reverse=True,
    )
    filtered = filtered[:MAX_RESULTS]

    print(f"  {len(filtered)} jobs after keyword + location + recency.\n")
    if not filtered:
        print("  No matches. Widen HOURS_BACK or KEYWORDS, or run --discover to refresh.")
        return

    out_path = f"{OUTPUT_PREFIX}_{dt.datetime.now(_IST).strftime('%Y-%m-%d_%H%M')}.xlsx"
    write_excel(filtered, out_path)
    print(f"  Wrote {len(filtered)} jobs -> {out_path}\n")
    for j in filtered[:12]:
        age = ""
        if j["posted_dt"]:
            hrs = (_now_utc() - j["posted_dt"]).total_seconds() / 3600
            age = f"  ({hrs:.1f}h)"
        print(f"    [{j['company']}] {j['role']} — {j['location']}{age}")


# ───────────────────────────────────────────────────────────────────────────
# Registry management helpers
# ───────────────────────────────────────────────────────────────────────────

def _rewrite_registry(entries):
    """Write a cleaned SEED list back to registry.py."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "registry.py")
    by_ats = {}
    for e in entries:
        by_ats.setdefault(e[1], []).append(e)

    lines = [
        '"""',
        "registry.py — Company ATS seed list for jobpull.",
        "",
        'Format: ("Display Name", "ats", "token")',
        '  ats in {"greenhouse", "lever", "ashby", "workday"}',
        "",
        'Workday token format: "tenant/wdX/career_path"',
        '  e.g. "crowdstrike/wd5/crowdstrikecareers"',
        "",
        "Run `python jobpull.py --clean-registry` to prune dead tokens.",
        '"""',
        "",
        "SEED = [",
    ]
    for ats_name in ["greenhouse", "lever", "ashby", "workday"]:
        group = sorted(by_ats.get(ats_name, []), key=lambda x: x[0].lower())
        if group:
            lines.append(f"    # {ats_name.upper()}")
            for name, ats, token in group:
                n = name.replace("\\", "\\\\").replace('"', '\\"')
                t = token.replace("\\", "\\\\").replace('"', '\\"')
                lines.append(f'    ("{n}", "{ats}", "{t}"),')
            lines.append("")
    lines.append("]")
    lines.append("")

    with open(path, "w") as f:
        f.write("\n".join(lines))


def mode_clean_registry(auto_yes=False):
    """Verify all tokens; remove entries that return hard HTTP errors (404/401/403/5xx).

    Pass auto_yes=True (or run with --yes flag) to skip the confirmation prompt.
    """
    # Use more workers for the verify pass to finish faster
    _workers = min(30, MAX_WORKERS * 2)
    print(f"Clean-registry: verifying {len(SEED)} tokens ({_workers} workers)...\n")

    all_jobs, summary = [], []
    total = len(SEED)
    done = 0
    from concurrent.futures import ThreadPoolExecutor, as_completed as _as_completed
    with ThreadPoolExecutor(max_workers=_workers) as pool:
        futures = {pool.submit(safe_fetch, c): c for c in SEED}
        for fut in _as_completed(futures):
            entry, jobs, err = fut.result()
            done += 1
            summary.append((entry[0], entry[1], len(jobs), err))
            print(f"\r  Verifying: {done}/{total}...", end="", flush=True)
    print()

    # Hard failures only — 400/401/403/404/410/422 or connection refused
    HARD_CODES = {"400", "401", "403", "404", "410", "422"}
    dead_keys, dead_list = set(), []
    for name, ats, n, err in summary:
        if err:
            code = "".join(c for c in (err or "") if c.isdigit())[:3]
            if code in HARD_CODES or "not found" in (err or "").lower():
                dead_keys.add((name, ats))
                dead_list.append((name, ats, err))

    live = [s for s in summary if s[3] is None]
    print(f"\n  Live: {len(live)}   Hard-dead: {len(dead_keys)}   "
          f"Zero-postings (kept): {len(summary)-len(live)-len(dead_keys)}")

    if dead_list:
        print(f"\n  Removing {len(dead_keys)} entries:")
        for name, ats, err in sorted(dead_list, key=lambda x: x[0])[:40]:
            print(f"    XX {name:30s}  {ats:11s}  {err[:45]}")
        if len(dead_list) > 40:
            print(f"    ... and {len(dead_list)-40} more")

    clean = [(n, a, t) for n, a, t in SEED if (n, a) not in dead_keys]
    print(f"\n  Before: {len(SEED)}  →  After: {len(clean)}  (removing {len(SEED)-len(clean)})")

    if not dead_keys:
        print("  Nothing to remove — registry already clean.")
        return

    if auto_yes:
        answer = "y"
    else:
        answer = input("\n  Write cleaned registry.py? [y/N] ").strip().lower()

    if answer == "y":
        _rewrite_registry(clean)
        print("  registry.py updated.")
    else:
        print("  Aborted — no changes made.")


def mode_add_company(spec):
    """Test a single company token and append it to registry.py if live.

    spec format:  "DisplayName|ats|token"
    Example:      "Notion|ashby|notion"
    """
    parts = [p.strip() for p in spec.split("|")]
    if len(parts) != 3:
        print(f"ERROR: expected 'DisplayName|ats|token', got: {spec!r}")
        return

    name, ats, token = parts
    ats = ats.lower()
    if ats not in FETCHERS:
        print(f"ERROR: unknown ATS '{ats}'. Valid: {', '.join(FETCHERS)}")
        return

    print(f"Testing  {name}  ({ats} / {token})...")
    entry = (name, ats, token)
    _, jobs, err = safe_fetch(entry)

    if err:
        print(f"  FAILED: {err}")
        print("  Not added to registry.")
        return

    print(f"  OK — {len(jobs)} total postings fetched.")

    # Check if already in SEED
    existing = [(n, a) for n, a, t in SEED if a == ats and t == token]
    if existing:
        print(f"  Already in registry as: {existing[0][0]}")
        return

    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "registry.py")
    with open(path, "a") as f:
        n_esc = name.replace('"', '\\"')
        t_esc = token.replace('"', '\\"')
        # Insert just before the closing bracket of SEED
    with open(path, "r") as f:
        content = f.read()

    insert_line = f'    ("{n_esc}", "{ats}", "{t_esc}"),\n'
    if content.rstrip().endswith("]"):
        new_content = content.rstrip()[:-1].rstrip() + "\n" + insert_line + "]\n"
        with open(path, "w") as f:
            f.write(new_content)
        print(f"  Added to registry.py")
    else:
        print(f"  Could not auto-insert (unexpected registry format). Add manually:")
        print(f'      {insert_line.strip()}')


def mode_smoke_test():
    """Quick end-to-end test against one confirmed company per ATS."""
    smoke = [
        ("Cloudflare",   "greenhouse", "cloudflare"),
        ("Sentry",       "ashby",      "sentry"),
        ("Atlassian",    "lever",      "atlassian"),
        ("CrowdStrike",  "workday",    "crowdstrike/wd5/crowdstrikecareers"),
    ]
    print("Smoke test — one company per ATS\n")
    all_ok = True
    for entry in smoke:
        name, ats, token = entry
        _, jobs, err = safe_fetch(entry)
        if err:
            print(f"  FAIL  [{ats:11s}] {name}: {err}")
            all_ok = False
        else:
            india_jobs = [j for j in jobs if location_ok_for_scan(j["location"])]
            kw_jobs = [j for j in india_jobs if matches_keywords(j["role"])]
            print(f"  OK    [{ats:11s}] {name}: {len(jobs)} total, {len(india_jobs)} India/APAC, {len(kw_jobs)} keyword-match")
            for j in kw_jobs[:3]:
                print(f"          {j['role'][:55]}  —  {j['location']}")

    print(f"\n{'All systems go.' if all_ok else 'Some ATSes failed — check tokens.'}")


# ───────────────────────────────────────────────────────────────────────────
# Web discovery — find new company ATS tokens via DuckDuckGo
# ───────────────────────────────────────────────────────────────────────────

_ATS_URL_PATTERNS = {
    "greenhouse": _re.compile(r"boards\.greenhouse\.io/([a-zA-Z0-9][a-zA-Z0-9_-]{1,60})"),
    "lever":      _re.compile(r"jobs\.lever\.co/([a-zA-Z0-9][a-zA-Z0-9_-]{1,60})"),
    "ashby":      _re.compile(r"jobs\.ashbyhq\.com/([a-zA-Z0-9][a-zA-Z0-9_-]{1,60})"),
}

_DISCOVER_QUERIES = [
    ("greenhouse", 'site:boards.greenhouse.io "devops engineer" "india"'),
    ("greenhouse", 'site:boards.greenhouse.io "support engineer" "bangalore"'),
    ("greenhouse", 'site:boards.greenhouse.io "site reliability" "india"'),
    ("greenhouse", 'site:boards.greenhouse.io "platform engineer" "india"'),
    ("greenhouse", 'site:boards.greenhouse.io "cloud engineer" "india"'),
    ("lever",      'site:jobs.lever.co "devops" "india"'),
    ("lever",      'site:jobs.lever.co "support engineer" "bangalore"'),
    ("lever",      'site:jobs.lever.co "site reliability" "india"'),
    ("lever",      'site:jobs.lever.co "platform engineer" "india"'),
    ("ashby",      'site:jobs.ashbyhq.com "devops" "india"'),
    ("ashby",      'site:jobs.ashbyhq.com "support engineer" "india"'),
    ("ashby",      'site:jobs.ashbyhq.com "site reliability" "india"'),
]

_DDG_NOISE_TOKENS = {
    "embed", "job_board", "api", "jobs", "all", "careers", "career",
    "search", "apply", "content", "index", "js", "css",
}


def _ddg_search_ats_tokens(ats, query):
    """Query DuckDuckGo HTML and extract ATS board tokens from result URLs."""
    pattern = _ATS_URL_PATTERNS.get(ats)
    if not pattern:
        return set()
    try:
        r = requests.get(
            "https://html.duckduckgo.com/html/",
            params={"q": query, "kl": "in-en"},
            headers={
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) "
                              "Chrome/120.0.0.0 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml",
                "Accept-Language": "en-US,en;q=0.9",
            },
            timeout=20,
            verify=_SSL,
        )
        tokens = set(
            m.group(1).lower()
            for m in pattern.finditer(r.text)
            if m.group(1).lower() not in _DDG_NOISE_TOKENS
        )
        return tokens
    except Exception:
        return set()


def mode_discover_web():
    """Search DuckDuckGo for companies posting India DevOps/SRE/Support jobs on public
    ATS boards (Greenhouse / Lever / Ashby).  Found tokens are tested against the real
    APIs; only companies with actual India/APAC jobs are written into registry.py.

    Usage:  python jobpull.py --discover-web
    """
    print("Web discovery: searching for India DevOps/SRE/Support roles on public ATS boards...\n")

    existing_tokens = {(a, t.lower()) for _, a, t in SEED}
    candidates = []
    seen = set()

    for ats, query in _DISCOVER_QUERIES:
        short_q = query[:60]
        print(f"  [{ats:11s}] {short_q}...", end=" ", flush=True)
        tokens = _ddg_search_ats_tokens(ats, query)
        new_t = [t for t in tokens if (ats, t) not in existing_tokens and t not in seen]
        for t in new_t:
            seen.add(t)
            candidates.append((f"[web] {t}", ats, t))
        print(f"{len(tokens)} found, {len(new_t)} new")
        time.sleep(1.8)   # polite pause between DDG requests

    if not candidates:
        print("\n  No new candidates found. Try again later (DDG rate-limits heavy use).")
        return

    print(f"\n  Testing {len(candidates)} candidates against ATS APIs...\n")

    to_add = []
    done = 0
    total = len(candidates)
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(safe_fetch, c): c for c in candidates}
        for fut in as_completed(futures):
            entry, jobs, err = fut.result()
            done += 1
            print(f"\r  Scanning: {done}/{total}", end="", flush=True)
            if err or not jobs:
                continue
            token = entry[2]
            india = [j for j in jobs if location_ok_for_scan(j["location"])]
            if not india:
                continue
            kw_count = sum(1 for j in india if matches_keywords(j["role"]))
            display = token.replace("-", " ").replace("_", " ").title()
            to_add.append((display, entry[1], token, len(india), kw_count))
    print()

    if not to_add:
        print("  Candidates found but none have India/APAC jobs right now.")
        return

    to_add.sort(key=lambda x: (-x[4], -x[3]))
    print(f"\n  {len(to_add)} new companies with India jobs:\n")
    for display, ats, token, india_n, kw_n in to_add:
        star = "**" if kw_n > 0 else "  "
        print(f"  {star} {display:32s} ({ats:11s}) "
              f"— {india_n} India/APAC jobs, {kw_n} keyword-match")

    import importlib, registry as _reg_mod
    importlib.reload(_reg_mod)
    new_seed = list(_reg_mod.SEED) + [(d, a, t) for d, a, t, _, _ in to_add]
    _rewrite_registry(new_seed)
    print(f"\n  Added {len(to_add)} companies. Registry total: {len(new_seed)}.")
    print("  Run '--clean-registry --yes' periodically to prune dead tokens.")


if __name__ == "__main__":
    if "--verify" in sys.argv:
        mode_verify()
    elif "--discover" in sys.argv:
        mode_discover()
    elif "--clean-registry" in sys.argv:
        mode_clean_registry(auto_yes="--yes" in sys.argv)
    elif "--add-company" in sys.argv:
        idx = sys.argv.index("--add-company")
        if idx + 1 >= len(sys.argv):
            print("Usage: python jobpull.py --add-company 'DisplayName|ats|token'")
        else:
            mode_add_company(sys.argv[idx + 1])
    elif "--smoke-test" in sys.argv:
        mode_smoke_test()
    elif "--discover-web" in sys.argv:
        mode_discover_web()
    else:
        mode_scan()
