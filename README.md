# Daily News Headlines

Automated financial news aggregator that scrapes headlines from major English and Chinese sources, deduplicates similar titles, and delivers a formatted daily digest via Outlook on Windows.

## Sources

Bloomberg, Reuters, Financial Times, CNBC, Seeking Alpha, 華爾街見聞, 香港經濟日報, 信報財經

## Features

- Per-source headline limits and logo branding
- Title deduplication with fuzzy matching
- CSV export with date-stamped filename
- HTML email digest sent through local Outlook

## Requirements

- Python 3.10+
- Windows 10/11 (uses `pywin32` for Outlook integration)
- Microsoft Outlook installed and configured

## Installation

```bash
pip install -r requirements.txt
```

Optional — for Bloomberg pages that require JavaScript rendering:

```bash
pip install playwright
playwright install chromium
```

## Usage

```bash
python Newsheadline.py
```

Output CSV is saved as `Daily News Headlines_DD Mon YYYY.csv` in the project folder.

## Note

This tool is for personal research workflows. Respect each news site's terms of service and robots.txt when scraping.

## License

MIT
