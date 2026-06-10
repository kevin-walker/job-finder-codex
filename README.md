# Job Finder

Pulls fresh job listings directly from company career pages (Greenhouse / Lever / Ashby public APIs).
No scraping, no API keys, no LinkedIn, no Naukri.

## Setup

```bash
pip3 install requests python-dateutil openpyxl truststore --break-system-packages
```

## Usage

### Daily scan (your main command)
```bash
python3 jobpull.py
```
Fetches all companies in the registry, filters by role and location, writes `jobs_YYYY-MM-DD.xlsx`.

### First-time discovery (run once, then weekly)
```bash
python3 jobpull.py --discover
```
Scans all 244+ companies, keeps only those actively hiring in India/APAC/US-remote, saves `india_companies.json`.  
After this, the daily scan runs faster against the cached shortlist.

### Verify live tokens
```bash
python3 jobpull.py --verify
```
Pings every ATS token and reports which are alive vs dead. Run after adding new companies to registry.py.

## What it finds

**Target roles** (configured in `jobpull.py → KEYWORDS`):
- Technical Support Engineer
- Support Engineer / Cloud Support
- Technical Solutions Engineer
- Technical Consultant
- DevOps Engineer
- Site Reliability Engineer (SRE)
- Platform / Infrastructure / Cloud Engineer

**Senior/lead/manager roles are excluded** automatically.

**Locations included:**
- India (Bangalore, Bengaluru, Kochi, Trivandrum, Hyderabad, Pune, Mumbai, Delhi, Chennai, Gurgaon, Noida)
- Remote – India
- Remote – APAC
- Remote – US

## Tuning

All config is at the top of `jobpull.py`:

| Setting | Default | What it does |
|---|---|---|
| `KEYWORDS` | target roles list | Add/remove role keywords |
| `EXCLUDE_SENIORITY` | senior/lead/staff/manager... | Titles that are filtered OUT |
| `HOURS_BACK` | 168 (7 days) | How far back to look |
| `MAX_RESULTS` | 200 | Max rows in the Excel output |
| `MAX_WORKERS` | 8 | Concurrent fetches |

## Adding more companies

Edit `registry.py` — add a line at the bottom:
```python
("Company Name", "greenhouse", "their-gh-token"),
("Another Co",   "lever",      "their-lever-slug"),
("Startup",      "ashby",      "their-ashby-slug"),
```
Then run `--verify` to confirm the token is live.

**How to find a token:**
- Greenhouse: look at `boards.greenhouse.io/TOKEN` or `job-boards.greenhouse.io/TOKEN`
- Lever: look at `jobs.lever.co/TOKEN`
- Ashby: look at `jobs.ashbyhq.com/TOKEN`

## Company coverage (244 companies)

| ATS | Count | Notable companies |
|---|---|---|
| Greenhouse | ~150 | Stripe, Cloudflare, Datadog, PhonePe, Razorpay, HackerRank, Harness, Netskope, InMobi, BrowserStack, Braze, Intercom, JFrog, Rubrik, KnowBe4, Amplitude, Klaviyo... |
| Lever | ~26 | Spotify, Binance, Hugging Face, Everbridge, Coursera, Outreach, HighLevel... |
| Ashby | ~68 | OpenAI, Cohere, ElevenLabs, Cognition, Deliveroo, Vanta, Sentry, ClickUp, Kong, Illumio, Drata, PostHog, Temporal, Zapier, Neon, WorkOS, Scribe... |
