# Daily News Headlines

Automated financial news aggregator that scrapes headlines from major English and Chinese sources, deduplicates similar titles, and delivers a formatted daily digest by email.

## Sources

Bloomberg, Reuters, Financial Times, CNBC, Seeking Alpha, 華爾街見聞, 香港經濟日報, 信報財經

## Features

- Per-source headline limits and logo branding
- Title deduplication with fuzzy matching
- CSV export with date-stamped filename
- HTML email digest via **Gmail SMTP** (GitHub Actions / any machine) or **local Outlook** (Windows)

## Requirements

- Python 3.10+
- For GitHub Actions / Gmail: a Gmail account with an [App Password](https://myaccount.google.com/apppasswords)
- For local Outlook sending (optional): Windows + Outlook + `pip install pywin32`

## Installation

```bash
pip install -r requirements.txt
```

Optional — for Bloomberg pages that require JavaScript rendering:

```bash
pip install playwright
playwright install chromium
```

Optional — local Outlook on Windows:

```bash
pip install pywin32
```

## Usage

### Local with Gmail SMTP

```bash
set EMAIL_USER=you@gmail.com
set EMAIL_PASS=your-16-char-app-password
set EMAIL_TO=you@company.com
python Newsheadline.py
```

### Local with Outlook (Windows)

Leave `EMAIL_USER` / `EMAIL_PASS` unset. The script uses Outlook and sends to `EMAIL_TO` (default: `hchan@penjing-am.com`).

```bash
python Newsheadline.py
```

Output CSV is saved as `Daily News Headlines_DD Mon YYYY.csv` in the project folder.

## GitHub Actions (daily email)

A workflow runs every day at **08:00 Hong Kong time** (`0 0 * * *` UTC) and can also be started manually from the Actions tab.

### 1. Create a Gmail App Password

1. Enable 2-Step Verification on your Google account
2. Open [App Passwords](https://myaccount.google.com/apppasswords)
3. Create one for “Mail” and copy the 16-character password

### 2. Add repository secrets

In the repo: **Settings → Secrets and variables → Actions → New repository secret**

| Secret | Value |
|--------|--------|
| `EMAIL_USER` | Your Gmail address |
| `EMAIL_PASS` | Gmail App Password (not your normal password) |
| `EMAIL_TO` | Corporate inbox, e.g. `hchan@penjing-am.com` |
| `EMAIL_FROM` | Optional; defaults to `EMAIL_USER` |

### 3. Run it

- **Automatic:** every day on the schedule
- **Manual:** Actions → **Daily news scrape** → **Run workflow**

If the run fails, open the job log. A CSV artifact is also uploaded when available.

## Note

This tool is for personal research workflows. Respect each news site's terms of service and robots.txt when scraping.

## License

MIT
