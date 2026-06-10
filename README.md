# Job Finder

Pulls fresh job listings directly from company career pages via public ATS APIs (Greenhouse / Lever / Ashby / Workday).  
No scraping, no API keys, no LinkedIn, no Naukri. Results go straight to Excel with IST timestamps.

## Setup

```bash
pip3 install requests python-dateutil openpyxl truststore --break-system-packages
```

## Usage

### Daily scan — your main command
```bash
python3 jobpull.py
```
Scans all 700+ companies in the registry, filters by role and location, writes `jobs_YYYY-MM-DD_HHMM.xlsx` (time in IST).

### Web discovery — find new companies automatically
```bash
python3 jobpull.py --discover-web
```
Searches DuckDuckGo for companies with India DevOps/SRE/Support jobs on Greenhouse, Lever, and Ashby boards.  
Tests each discovered token against the real API and adds companies with India/APAC jobs to `registry.py`.  
Run this periodically to grow the registry organically.

### Clean up dead tokens
```bash
python3 jobpull.py --clean-registry --yes
```
Verifies all tokens concurrently (30 workers). Removes entries that return hard HTTP errors (404/401/403/422).  
Run this after `--discover-web` or after manually adding new companies.

### Add a single company manually
```bash
python3 jobpull.py --add-company "CompanyName|ats|token"
# e.g.:
python3 jobpull.py --add-company "Notion|ashby|notion"
python3 jobpull.py --add-company "HashiCorp|lever|hashicorp"
python3 jobpull.py --add-company "CrowdStrike|workday|crowdstrike/wd5/crowdstrikecareers"
```
Tests the token first — only appends to `registry.py` if it's live.

### Smoke test — confirm all 4 ATSes are working
```bash
python3 jobpull.py --smoke-test
```
Quick sanity check (~30 s). Tests one confirmed company per ATS.

### Verify all tokens
```bash
python3 jobpull.py --verify
```
Pings every token and reports live vs dead. Useful after a bulk registry update.

### Discovery (legacy cache mode)
```bash
python3 jobpull.py --discover
```
Scans all companies, caches ones that have India roles to `india_companies.json`. Older workflow — `--discover-web` is preferred.

---

## What it finds

**Target roles** (edit `KEYWORDS` in `jobpull.py`):
- DevOps Engineer / Senior DevOps Engineer
- Site Reliability Engineer (SRE)
- Platform Engineer / Infrastructure Engineer
- Cloud Support Engineer / Cloud Engineer
- Technical Support Engineer / Support Engineer
- Technical Solutions Engineer / Technical Consultant

**Manager / Director / VP / Architect** roles are excluded automatically.

**Locations included:**
- India — Bangalore, Bengaluru, Hyderabad, Pune, Mumbai, Delhi, Chennai, Kochi, Gurgaon, Noida, Trivandrum
- Remote – India / Remote – APAC
- US Remote is **excluded**

---

## Excel output

Each run writes a timestamped file: `jobs_2026-06-10_1357.xlsx`

| Column | Example |
|---|---|
| Role | Senior DevOps Engineer |
| Company | Harness |
| Location | Bangalore, India |
| Posted | 10 Jun 2026  12:45 PM IST |
| Age | 3h ago |
| Source | Greenhouse |
| Link | Open (clickable hyperlink) |

---

## Config

All tuning knobs are at the top of `jobpull.py`:

| Setting | Default | What it does |
|---|---|---|
| `KEYWORDS` | target roles list | Add/remove role title keywords |
| `EXCLUDE_SENIORITY` | manager/director/vp... | Titles filtered OUT |
| `HOURS_BACK` | 336 (14 days) | How far back to look |
| `MAX_RESULTS` | 400 | Max rows in Excel output |
| `MAX_WORKERS` | 15 | Concurrent fetches per run |

---

## Registry — 701 companies

| ATS | Count | Notable companies |
|---|---|---|
| Greenhouse | 380 | Cloudflare, Datadog, PhonePe, Razorpay, Harness, Stripe, HackerRank, Netskope, BrowserStack, JFrog, Rubrik, Confluent, Wiz, Snyk, CyberArk, Lacework, UiPath, Talkdesk, DigitalOcean, Palo Alto Networks, Datto, RingCentral, Zscaler, Zoho, Gupshup... |
| Ashby | 142 | OpenAI, Cohere, ElevenLabs, Cognition, Sentry, PostHog, Supabase, Vanta, Kong, Drata, ClickUp, Neon, WorkOS, Zapier, Notion, Linear, Vercel, Ramp, Harvey, Perplexity, Statsig, SigNoz, Coralogix, Rootly... |
| Lever | 110 | Atlassian, Spotify, Netflix, Canva, Miro, GitHub, HashiCorp, Freshworks, Monday.com, Gong, Outreach, Sysdig, Auth0, Seismic, Netlify, Retool, UserTesting, Timescale, TigerGraph... |
| Workday | 69 | CrowdStrike, Dell, Trend Micro, Zendesk, ServiceNow, IBM, Cisco, SAP, Infosys, Wipro, HCL, TCS, Cognizant, Tech Mahindra, LTIMindtree, Salesforce, Adobe, Nutanix, NetApp, Splunk, Capgemini, Accenture... |

> **Note:** Workday enterprise tokens are best-guess — run `--clean-registry --yes` to prune any that return 404.

### How to find a token

- **Greenhouse:** token is in `boards.greenhouse.io/TOKEN`
- **Lever:** token is in `jobs.lever.co/TOKEN`
- **Ashby:** token is in `jobs.ashbyhq.com/TOKEN`
- **Workday:** token is `tenant/wdX/career_path` from `tenant.wdX.myworkdayjobs.com/en-US/career_path`

### Workday token format

```
crowdstrike/wd5/crowdstrikecareers
↑ tenant    ↑ server  ↑ career path
```

---

## Recommended workflow

```bash
# First time setup
python3 jobpull.py --smoke-test          # confirm all 4 ATSes work
python3 jobpull.py --clean-registry --yes  # prune dead tokens from the 701

# Daily use
python3 jobpull.py                        # fresh Excel in ~2 min

# Weekly — grow the registry
python3 jobpull.py --discover-web         # find new companies from web search
python3 jobpull.py --clean-registry --yes # prune any new dead tokens
```
