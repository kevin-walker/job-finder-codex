# Job Finder

Pulls fresh job listings directly from company career pages via public ATS APIs — **Greenhouse / Lever / Ashby / Workday**.  
No scraping, no API keys, no LinkedIn, no Naukri. Results go straight to Excel (IST timestamps) inside `job_postings_results/`.

---

## Setup

### macOS

```bash
# Python dependencies
pip3 install requests python-dateutil openpyxl truststore requests[socks] --break-system-packages
```

### Linux (Ubuntu / Debian)

```bash
# Python dependencies
pip3 install requests python-dateutil openpyxl requests[socks]

# truststore is macOS-only — skip it on Linux
```

### Windows

```powershell
# Open Command Prompt or PowerShell
pip install requests python-dateutil openpyxl requests[socks]

# Use python instead of python3
python jobpull.py
```

> On Windows, replace `python3` with `python` and `pip3` with `pip` in all commands below.

---

## Usage

### Daily scan — your main command
```bash
python3 jobpull.py
```
Scans all 594 companies, filters by role and location, writes  
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

> **DuckDuckGo rate limits:** If all queries return 0, DDG has blocked your IP temporarily. Wait 30 minutes and try again, or use Tor for automatic IP rotation (see below).

---

### Web discovery with Tor — rotate IP between queries

DuckDuckGo rate-limits repeated searches from the same IP. Running `--discover-web` through Tor gives each query a different exit-node IP, effectively bypassing the limit.

The script auto-detects whether Tor is running. If it is, IP rotation happens automatically — no extra flags needed.

#### macOS

```bash
# Install Tor
brew install tor

# Start Tor (runs in background, takes ~10 s to bootstrap)
tor --ControlPort 9051 --CookieAuthentication 0 &

# Run discovery — Tor is detected automatically
python3 jobpull.py --discover-web

# Stop Tor when done
pkill tor
```

#### Linux (Ubuntu / Debian)

```bash
# Install Tor
sudo apt install tor

# Edit torrc to enable the control port
echo -e "ControlPort 9051\nCookieAuthentication 0" | sudo tee -a /etc/tor/torrc

# Start Tor service
sudo systemctl start tor

# Run discovery
python3 jobpull.py --discover-web

# Stop Tor when done
sudo systemctl stop tor
```

#### Windows

```powershell
# Download Tor Expert Bundle from https://www.torproject.org/download/tor/
# Extract to e.g. C:\tor\

# Create a torrc file at C:\tor\torrc with these two lines:
#   ControlPort 9051
#   CookieAuthentication 0

# Start Tor (in a separate terminal window, keep it open)
C:\tor\tor.exe -f C:\tor\torrc

# In your main terminal, run discovery
python jobpull.py --discover-web
```

**How it works:**
- Each query gets a brand-new Tor circuit (new exit-node IP)
- After each query, `SIGNAL NEWNYM` is sent to the Tor control port
- A fresh session with new cookies is created after each IP rotation
- DDG sees 16 different IPs across 16 queries instead of the same IP 16 times
- Each rotation takes ~5 s for Tor to build the new circuit, so the full run takes ~2 min

**Without Tor:** the script falls back to a persistent browser-like session with a 5-second pause between queries. This works for occasional use but may hit limits if run multiple times in the same hour.

---

### Clean up dead tokens
```bash
python3 jobpull.py --clean-registry --yes
```
Verifies all tokens concurrently (30 workers). Removes entries returning hard HTTP errors (404/401/403/422).  
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

## Registry — 594 companies

| ATS | Count | Notable companies |
|---|---|---|
| Greenhouse | 215 | Cloudflare, Datadog, PhonePe, Razorpay, Harness, Stripe, HackerRank, Netskope, BrowserStack, JFrog, Rubrik, Confluent, Wiz, Snyk, CyberArk, Lacework, UiPath, Talkdesk, DigitalOcean, Palo Alto Networks, Datto, RingCentral, Zscaler, Zoho, Gupshup... |
| Ashby | 108 | OpenAI, Cohere, ElevenLabs, Cognition, Sentry, PostHog, Supabase, Vanta, Kong, Drata, ClickUp, Neon, WorkOS, Zapier, Notion, Linear, Vercel, Ramp, Harvey, Perplexity, Statsig, SigNoz, Coralogix, Rootly... |
| Workday | 242 | CrowdStrike, Accenture, HPE, GE Vernova, NVIDIA, F5, LSEG, Broadridge, Worldpay, Thomson Reuters, Red Hat, State Street, Hitachi, Morgan Stanley, JLL, Synechron, TSYS, Boeing, Mastercard, Unisys, Zendesk, Visa, Equifax... |
| Lever | 29 | Atlassian, Spotify, Netflix, Canva, Miro, GitHub, HashiCorp, Freshworks, Monday.com, Gong, Outreach, Sysdig, Auth0, Seismic, Netlify, Retool, UserTesting, Timescale, TigerGraph... |

### How to find a token manually

| ATS | URL pattern | Token |
|---|---|---|
| Greenhouse | `boards.greenhouse.io/TOKEN` | `TOKEN` |
| Lever | `jobs.lever.co/TOKEN` | `TOKEN` |
| Ashby | `jobs.ashbyhq.com/TOKEN` | `TOKEN` |
| Workday | `TENANT.wdX.myworkdayjobs.com/en-US/CAREER_PATH` | `TENANT/wdX/CAREER_PATH` |

> **Workday tip:** Run `python3 jobpull.py --find-workday-path tenant/wdX` to brute-force the career_path for a known tenant. Or just run `--discover-web` which extracts the exact path from real job URLs.

---

## Recommended workflow

```bash
# First-time setup
python3 jobpull.py --smoke-test            # confirm all 4 ATSes work
python3 jobpull.py --clean-registry --yes  # prune dead tokens

# Daily use
python3 jobpull.py                         # fresh Excel → job_postings_results/

# Weekly — grow the registry
tor --ControlPort 9051 --CookieAuthentication 0 &   # optional but recommended
python3 jobpull.py --discover-web          # find new companies via DuckDuckGo
python3 jobpull.py --clean-registry --yes  # prune any newly dead tokens
pkill tor                                  # stop Tor when done
```
