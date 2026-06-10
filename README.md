# Job Finder

Pulls fresh job listings directly from company career pages via public ATS APIs — **Greenhouse / Lever / Ashby / Workday**.  
No scraping, no API keys, no LinkedIn, no Naukri. Results go straight to Excel (IST timestamps) inside `job_postings_results/`.

## Setup

```bash
pip3 install requests python-dateutil openpyxl truststore --break-system-packages
```

## Usage

### Daily scan — your main command
```bash
python3 jobpull.py
```
Scans all 700+ companies, filters by role and location, writes  
`job_postings_results/jobs_YYYY-MM-DD_HHMM.xlsx`  
The `job_postings_results/` folder is created automatically on first run.

---

### Web discovery — find new companies automatically
```bash
python3 jobpull.py --discover-web
```
Searches DuckDuckGo for companies posting India DevOps/SRE/Support jobs on all four ATS platforms:

| ATS | Search pattern |
|---|---|
| Greenhouse | `site:boards.greenhouse.io "devops engineer" "india"` (5 queries) |
| Lever | `site:jobs.lever.co "support engineer" "bangalore"` (4 queries) |
| Ashby | `site:jobs.ashbyhq.com "site reliability" "india"` (3 queries) |
| Workday | `site:myworkdayjobs.com "platform engineer" "india"` (4 queries) |

Discovered tokens are tested against the real APIs. Only companies with confirmed India/APAC jobs get added to `registry.py`.  
Workday URLs (`tenant.wdX.myworkdayjobs.com/en-US/career_path/...`) are parsed automatically into the correct composite token format.

Run this periodically to grow the registry organically.

---

### Clean up dead tokens
```bash
python3 jobpull.py --clean-registry --yes
```
Verifies all 700+ tokens concurrently (30 workers). Removes entries returning hard HTTP errors (404/401/403/422).  
Run after `--discover-web` or after manually adding companies.

---

### Add a single company manually
```bash
python3 jobpull.py --add-company "CompanyName|ats|token"

# Examples:
python3 jobpull.py --add-company "Notion|ashby|notion"
python3 jobpull.py --add-company "HashiCorp|lever|hashicorp"
python3 jobpull.py --add-company "ServiceNow|workday|servicenow/wd5/ServiceNowCareers"
```
Tests the token first — only appends to `registry.py` if it's live.

---

### Smoke test — confirm all 4 ATSes work
```bash
python3 jobpull.py --smoke-test
```
Quick sanity check (~30 s). Tests one confirmed company per ATS (Cloudflare, Sentry, Atlassian, CrowdStrike).

---

### Verify all tokens
```bash
python3 jobpull.py --verify
```
Pings every token and reports live vs dead counts. Useful after a bulk registry update.

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

Files are saved to `job_postings_results/` (created automatically).  
Example filename: `job_postings_results/jobs_2026-06-10_1357.xlsx`

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
| `EXCLUDE_SENIORITY` | manager/director/vp/architect... | Titles filtered OUT |
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

> **Workday note:** Enterprise tokens are best-guess — run `--clean-registry --yes` to prune any that return 404. The `--discover-web` command also finds new Workday companies automatically by parsing `tenant.wdX.myworkdayjobs.com/en-US/career_path` URLs from search results.

### How to find a token manually

| ATS | URL pattern | Token |
|---|---|---|
| Greenhouse | `boards.greenhouse.io/TOKEN` | `TOKEN` |
| Lever | `jobs.lever.co/TOKEN` | `TOKEN` |
| Ashby | `jobs.ashbyhq.com/TOKEN` | `TOKEN` |
| Workday | `TENANT.wdX.myworkdayjobs.com/en-US/CAREER_PATH` | `TENANT/wdX/CAREER_PATH` |

---

## Recommended workflow

```bash
# First-time setup
python3 jobpull.py --smoke-test            # confirm all 4 ATSes work
python3 jobpull.py --clean-registry --yes  # prune dead tokens from the 701

# Daily use
python3 jobpull.py                         # fresh Excel → job_postings_results/

# Weekly — grow the registry
python3 jobpull.py --discover-web          # find new companies via DuckDuckGo
python3 jobpull.py --clean-registry --yes  # prune any newly dead tokens
```
