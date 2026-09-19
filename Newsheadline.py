import base64
import html
import json
import mimetypes
import os
import re
import smtplib
import ssl
from datetime import date
from difflib import SequenceMatcher
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import parse_qs, unquote, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

try:
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False

# --- Configuration ---
MAX_TOTAL = 30
DEDUP_SIMILAR_TITLES = True
DEDUP_THRESHOLD = 0.75
EMAIL_SUBJECT_PREFIX = "Daily News Headlines"
TRY_PLAYWRIGHT_FOR_BLOOMBERG = True

# Email: set EMAIL_USER + EMAIL_PASS (Gmail App Password) to send via SMTP.
# Falls back to local Outlook on Windows when SMTP secrets are not set.
EMAIL_TO = os.environ.get("EMAIL_TO", "hchan@penjing-am.com")
EMAIL_USER = os.environ.get("EMAIL_USER", "").strip()
EMAIL_PASS = os.environ.get("EMAIL_PASS", "").strip()
EMAIL_FROM = os.environ.get("EMAIL_FROM", EMAIL_USER).strip()
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))

SOURCE_LIMITS = {
    "Bloomberg": 4,
    "Reuters": 5,
    "Financial Times": 5,
    "Seeking Alpha": 3,
    "CNBC": 5,
    "華爾街見聞": 10,
    "香港經濟日報": 5,
    "信報財經": 5,
}

LOGOS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logos")

# User-provided logos — never auto-generate placeholders for these.
CUSTOM_LOGO_FILES = {"bloomberg.png", "reuters.png", "hket.png"}

SOURCE_LOGO_SPECS = {
    "Bloomberg": ("bloomberg.png", "Bloomberg", "#2800FF", "#FFFFFF"),
    "Reuters": ("reuters.png", "Reuters", "#FF8000", "#FFFFFF"),
    "Financial Times": ("financial_times.png", "Financial Times", "#FFF1E5", "#333333"),
    "CNBC": ("cnbc.png", "CNBC", "#005594", "#FFFFFF"),
    "Seeking Alpha": ("seeking_alpha.png", "Seeking Alpha", "#FF7200", "#FFFFFF"),
    "華爾街見聞": ("wallstreetcn.png", "华尔街见闻", "#C41E3A", "#FFFFFF"),
    "香港經濟日報": ("hket.png", "香港經濟日報", "#003366", "#FFFFFF"),
    "信報財經": ("hkej.png", "信報財經", "#8B0000", "#FFFFFF"),
}

BLOOMBERG_PAGE_URLS = (
    "https://www.bloomberg.com/markets",
    "https://www.bloomberg.com/",
)

BLOOMBERG_SKIP_TITLE_PATTERNS = (
    r"^Latest ",
    r"Podcast$",
    r"^Odd Lots:",
    r"MLIV$",
)
BLOOMBERG_SKIP_URL_PATTERNS = (
    r"/news/videos/",
    r"/news/audio/",
    r"/news/newsletters/",
)

today = date.today().strftime("%d %b %Y")
headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
results = []
seen_titles = []
filePath = os.path.dirname(os.path.abspath(__file__))
_LOCAL_IMAGE_CACHE = {}


def ensure_logos():
    os.makedirs(LOGOS_DIR, exist_ok=True)
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        return

    for _source, (filename, label, bg, fg) in SOURCE_LOGO_SPECS.items():
        if filename in CUSTOM_LOGO_FILES:
            continue
        path = os.path.join(LOGOS_DIR, filename)
        if os.path.isfile(path):
            continue
        img = Image.new("RGB", (400, 200), bg)
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("arial.ttf", 28)
        except OSError:
            font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), label, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        draw.text(
            ((400 - text_width) / 2, (200 - text_height) / 2),
            label,
            fill=fg,
            font=font,
        )
        img.save(path)


ensure_logos()

DEFAULT_IMAGES = {
    source: os.path.join(LOGOS_DIR, spec[0]) for source, spec in SOURCE_LOGO_SPECS.items()
}


def is_remote_image_url(image_url):
    return image_url.startswith("http://") or image_url.startswith("https://")


def local_image_to_data_uri(image_path):
    if image_path in _LOCAL_IMAGE_CACHE:
        return _LOCAL_IMAGE_CACHE[image_path]
    if not image_path or not os.path.isfile(image_path):
        return ""
    mime_type = mimetypes.guess_type(image_path)[0] or "image/png"
    with open(image_path, "rb") as image_file:
        encoded = base64.b64encode(image_file.read()).decode("ascii")
    data_uri = f"data:{mime_type};base64,{encoded}"
    _LOCAL_IMAGE_CACHE[image_path] = data_uri
    return data_uri


def extract_page_summary(soup):
    candidates = []

    for prop in ("og:description",):
        meta = soup.find("meta", property=prop)
        if meta and meta.get("content"):
            candidates.append(meta["content"].strip())

    for name in ("twitter:description", "description"):
        meta = soup.find("meta", attrs={"name": name})
        if meta and meta.get("content"):
            candidates.append(meta["content"].strip())

    for script in soup.find_all("script", type="application/ld+json"):
        if not script.string:
            continue
        try:
            data = json.loads(script.string)
        except json.JSONDecodeError:
            continue
        items = data if isinstance(data, list) else [data]
        for item in items:
            if isinstance(item, dict):
                for key in ("description", "abstract"):
                    value = item.get(key)
                    if isinstance(value, str) and value.strip():
                        candidates.append(value.strip())

    for selector in (
        ".o-topper__standfirst",
        "[data-trackable='standfirst']",
        "div[data-trackable='body'] p",
        ".article__content-body p",
        ".article-body__content p",
        "[data-testid='paragraph']",
    ):
        element = soup.select_one(selector)
        if element:
            text = element.get_text(" ", strip=True)
            if len(text) >= 40:
                candidates.append(text)

    seen = set()
    best = ""
    for candidate in candidates:
        normalized = " ".join(candidate.split())
        key = normalized.lower()
        if key in seen or len(normalized) < 20:
            continue
        seen.add(key)
        if len(normalized) > len(best):
            best = normalized
    return best


def extract_bing_news_url(bing_link):
    query = parse_qs(urlparse(bing_link).query)
    return unquote(query.get("url", [""])[0]).strip()


def is_weak_summary(summary, title):
    if not summary:
        return True
    normalized_summary = " ".join(summary.lower().split())
    normalized_title = " ".join(title.lower().split())
    if normalized_summary == normalized_title:
        return True
    if normalized_summary == f"{normalized_title} reuters":
        return True
    if normalized_summary.startswith(normalized_title) and len(normalized_summary) <= len(
        normalized_title
    ) + 12:
        return True
    return False


def normalize_title(title):
    return " ".join(title.lower().split())


def is_duplicate_title(title):
    if not DEDUP_SIMILAR_TITLES:
        return normalize_title(title) in {normalize_title(t) for t in seen_titles}
    normalized = normalize_title(title)
    for seen in seen_titles:
        if SequenceMatcher(None, normalized, normalize_title(seen)).ratio() >= DEDUP_THRESHOLD:
            return True
    return False


def add_result(source, title, url, summary="", image_url=""):
    title = title.strip()
    url = url.strip()
    summary = summary.strip()
    image_url = image_url.strip()
    if not title or not url:
        return False
    if is_duplicate_title(title):
        return False
    seen_titles.append(title)
    results.append(
        {
            "source": source,
            "title": title,
            "url": url,
            "summary": summary,
            "image_url": image_url or DEFAULT_IMAGES.get(source, ""),
        }
    )
    return True


def fetch_soup(url, parser="html.parser"):
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return BeautifulSoup(response.content, parser)


def strip_html(text):
    if not text:
        return ""
    return BeautifulSoup(text, "html.parser").get_text(" ", strip=True)


def extract_story_image(story):
    for key in ("image", "thumbnail", "ledeImage"):
        value = story.get(key)
        if isinstance(value, dict):
            url = value.get("url") or value.get("src") or ""
            if url:
                return url
        elif isinstance(value, str) and value:
            return value
    return ""


def extract_story_summary(story):
    for key in ("summary", "abstract", "dek", "subheadline"):
        value = story.get(key) or ""
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def extract_story_url(story):
    url = story.get("longURL") or story.get("url") or story.get("link") or ""
    if url.startswith("/"):
        url = "https://www.bloomberg.com" + url
    return url.split("?")[0]


def extract_story_title(story):
    return (story.get("headline") or story.get("title") or "").strip()


def should_skip_bloomberg_item(title, url):
    for pattern in BLOOMBERG_SKIP_TITLE_PATTERNS:
        if re.search(pattern, title, re.IGNORECASE):
            return True
    for pattern in BLOOMBERG_SKIP_URL_PATTERNS:
        if re.search(pattern, url, re.IGNORECASE):
            return True
    return False


def normalize_bloomberg_url(url):
    if url.startswith("/"):
        url = "https://www.bloomberg.com" + url
    return url.split("?")[0]


def extract_headline_text(element):
    if not element:
        return ""
    headline = element.select_one("[data-component='headline']")
    if headline:
        return headline.get_text(strip=True)
    return element.get_text(strip=True)


def parse_bloomberg_lede_html(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    lede = soup.select_one(
        "section#lede, section[data-canonical='lede'], [data-component='LineupContentLede']"
    )
    if not lede:
        return []

    articles = []
    seen_urls = set()

    primary = lede.select_one("[data-testid='primary-story']")
    if primary:
        link = primary.select_one("a[data-component='story-link']")
        if link:
            url = normalize_bloomberg_url(link.get("href", ""))
            title = extract_headline_text(link)
            summary_el = link.select_one("[data-component='summary']")
            img_el = primary.select_one("img[data-component='image']")
            if title and url and url not in seen_urls:
                seen_urls.add(url)
                articles.append(
                    {
                        "title": title,
                        "url": url,
                        "summary": summary_el.get_text(strip=True) if summary_el else "",
                        "image_url": img_el.get("src", "") if img_el else "",
                    }
                )

    for link in lede.select("[data-testid='secondary-stories'] a[data-component='story-link']"):
        url = normalize_bloomberg_url(link.get("href", ""))
        title = extract_headline_text(link)
        if not title or not url or url in seen_urls:
            continue
        seen_urls.add(url)
        articles.append({"title": title, "url": url, "summary": "", "image_url": ""})

    return articles


def parse_bloomberg_lede_next_data(page_html):
    soup = BeautifulSoup(page_html, "html.parser")
    script = soup.find("script", id="__NEXT_DATA__")
    if not script or not script.string:
        return []

    data = json.loads(script.string)
    modules = data.get("props", {}).get("pageProps", {}).get("initialState", {}).get("modulesById", {})
    for module in modules.values():
        is_lede = (
            module.get("canonical") == "lede"
            or module.get("provenance") == "canonical"
            or (module.get("meta") or {}).get("canonical") == "lede"
            or (module.get("meta") or {}).get("trackingTitle") == "Lede"
        )
        stories = module.get("stories") or []
        if not is_lede or not stories:
            continue
        articles = []
        for story in stories:
            title = extract_story_title(story)
            url = extract_story_url(story)
            if not title or not url:
                continue
            articles.append(
                {
                    "title": title,
                    "url": url,
                    "summary": extract_story_summary(story),
                    "image_url": extract_story_image(story),
                }
            )
        if articles:
            return articles
    return []


def parse_bloomberg_lede(page_html):
    articles = parse_bloomberg_lede_html(page_html)
    if articles:
        return articles
    return parse_bloomberg_lede_next_data(page_html)


def _fetch_bloomberg_with_playwright(headless):
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=headless,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            user_agent=headers["User-Agent"],
            locale="en-US",
            viewport={"width": 1920, "height": 1080},
        )
        page = context.new_page()
        page.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        for page_url in BLOOMBERG_PAGE_URLS:
            page.goto(page_url, wait_until="domcontentloaded", timeout=90000)
            try:
                page.wait_for_selector(
                    "section#lede, [data-component='LineupContentLede']",
                    timeout=15000,
                )
            except Exception:
                page.wait_for_timeout(8000)
            html = page.content()
            if "Are you a robot" not in html and parse_bloomberg_lede(html):
                browser.close()
                return html
        html = page.content()
        browser.close()
        if "Are you a robot" in html:
            return None
        return html


def fetch_bloomberg_homepage_html():
    if not TRY_PLAYWRIGHT_FOR_BLOOMBERG:
        return None
    try:
        html = _fetch_bloomberg_with_playwright(headless=True)
        if html:
            return html
        return _fetch_bloomberg_with_playwright(headless=False)
    except Exception:
        return None


def collect_bloomberg_rss_items():
    feeds = (
        "https://feeds.bloomberg.com/markets/news.rss",
        "https://feeds.bloomberg.com/economics/news.rss",
        "https://feeds.bloomberg.com/politics/news.rss",
    )
    items = []
    seen_urls = set()
    for feed_url in feeds:
        soup = fetch_soup(feed_url, parser="xml")
        for item in soup.find_all("item"):
            url = item.link.text.strip() if item.link else ""
            if not url or url in seen_urls:
                continue
            seen_urls.add(url)
            items.append(
                {
                    "title": item.title.text.strip() if item.title else "",
                    "url": url,
                    "summary": item.description.text.strip() if item.description else "",
                }
            )
    return items


def scrape_bloomberg_from_rss(limit):
    """Fallback: top headlines from Bloomberg RSS feeds (markets first)."""
    items = collect_bloomberg_rss_items()
    count = 0
    for item in items:
        title = item["title"]
        url = item["url"]
        if should_skip_bloomberg_item(title, url):
            continue
        summary = strip_html(item.get("summary", ""))
        if add_result("Bloomberg", title, url, summary=summary):
            count += 1
        if count >= limit:
            break
    return count


def scrape_bloomberg(limit=None):
    """Scrape Bloomberg homepage lede when possible; otherwise use RSS."""
    limit = limit or SOURCE_LIMITS["Bloomberg"]

    if TRY_PLAYWRIGHT_FOR_BLOOMBERG:
        homepage_html = fetch_bloomberg_homepage_html()
        if homepage_html:
            articles = parse_bloomberg_lede(homepage_html)
            count = 0
            for article in articles[:limit]:
                if add_result(
                    "Bloomberg",
                    article["title"],
                    article["url"],
                    summary=article.get("summary", ""),
                    image_url=article.get("image_url", ""),
                ):
                    count += 1
            if count:
                return

    scrape_bloomberg_from_rss(limit)


def enrich_article_metadata(article):
    if article.get("summary") and article.get("image_url"):
        return article
    try:
        response = requests.get(article["url"], headers=headers, timeout=20)
        if response.status_code != 200:
            return article
        soup = BeautifulSoup(response.content, "html.parser")
        if not article.get("summary"):
            summary = extract_page_summary(soup)
            if summary:
                article["summary"] = summary
        if not article.get("image_url"):
            meta = soup.find("meta", property="og:image")
            if meta and meta.get("content"):
                article["image_url"] = meta["content"].strip()
    except Exception:
        pass
    return article


def scrape_reuters(limit=None):
    limit = limit or SOURCE_LIMITS["Reuters"]
    url = "https://www.bing.com/news/search?q=site:reuters.com+markets&format=rss"
    soup = fetch_soup(url, parser="xml")
    count = 0
    for item in soup.find_all("item"):
        title = item.title.text.strip() if item.title else ""
        bing_link = item.link.text.strip() if item.link else ""
        article_url = extract_bing_news_url(bing_link)
        if not title or not article_url or "reuters.com" not in article_url:
            continue
        summary = strip_html(item.description.text if item.description else "")
        summary = re.sub(r"^[\W\s\d]*", "", summary).strip()
        if is_weak_summary(summary, title):
            summary = ""
        if add_result("Reuters", title, article_url, summary=summary):
            count += 1
        if count >= limit:
            break


def scrape_financial_times(limit=None):
    limit = limit or SOURCE_LIMITS["Financial Times"]
    soup = fetch_soup("https://www.ft.com/")
    count = 0
    local_seen = set()
    for element in soup.select("[data-trackable='heading'], a[href*='/content/']"):
        link = element if element.name == "a" else element.find_parent("a") or element.find("a")
        if not link:
            continue
        title = link.get_text(strip=True)
        url = link.get("href", "")
        if not title or not url or title in local_seen:
            continue
        if "/content/" not in url:
            continue
        local_seen.add(title)
        if url.startswith("/"):
            url = "https://www.ft.com" + url.split("?")[0]
        article = enrich_article_metadata({"url": url, "summary": "", "image_url": ""})
        if add_result(
            "Financial Times",
            title,
            url,
            summary=article.get("summary", ""),
            image_url=article.get("image_url", ""),
        ):
            count += 1
        if count >= limit:
            break


def scrape_seeking_alpha(limit=None):
    limit = limit or SOURCE_LIMITS["Seeking Alpha"]
    soup = fetch_soup("https://seekingalpha.com/market-news/trending")
    count = 0
    local_seen = set()
    for link in soup.select('a[href*="/article/"], a[href*="/news/"]'):
        title = link.get_text(strip=True)
        url = link.get("href", "")
        if not title or not url or title in local_seen:
            continue
        local_seen.add(title)
        if url.startswith("/"):
            url = "https://seekingalpha.com" + url
        article = enrich_article_metadata({"url": url, "summary": "", "image_url": ""})
        if add_result(
            "Seeking Alpha",
            title,
            url,
            summary=article.get("summary", ""),
            image_url=article.get("image_url", ""),
        ):
            count += 1
        if count >= limit:
            break


def scrape_cnbc(limit=None):
    limit = limit or SOURCE_LIMITS["CNBC"]
    soup = fetch_soup(
        "https://search.cnbc.com/rs/search/combinedcms/view.xml?"
        "partnerId=wrss01&id=100003114",
        parser="xml",
    )
    count = 0
    for item in soup.find_all("item"):
        title = item.title.text.strip() if item.title else ""
        url = item.link.text.strip() if item.link else ""
        summary = strip_html(item.description.text if item.description else "")
        article = enrich_article_metadata({"url": url, "summary": summary, "image_url": ""})
        if add_result(
            "CNBC",
            title,
            url,
            summary=article.get("summary", summary),
            image_url=article.get("image_url", ""),
        ):
            count += 1
        if count >= limit:
            break


def fetch_wallstreetcn_article_detail(article_id):
    api_url = (
        f"https://api-one-wscn.awtmt.com/apiv1/content/articles/"
        f"{article_id}?extract=0"
    )
    try:
        response = requests.get(api_url, headers=headers, timeout=20)
        response.raise_for_status()
        data = response.json().get("data", {})
        summary = data.get("content_short", "").strip()
        if not summary:
            summary = strip_html(data.get("content", ""))[:300]
        image = (data.get("image") or {}).get("uri", "").strip()
        return summary, image
    except Exception:
        return "", ""


def scrape_wallstreetcn(limit=None):
    """Scrape 华尔街见闻 homepage 最热文章 ranked list."""
    limit = limit or SOURCE_LIMITS["華爾街見聞"]
    api_url = "https://api-one-wscn.awtmt.com/apiv1/content/articles/hot?period=all"
    response = requests.get(api_url, headers=headers, timeout=30)
    response.raise_for_status()
    data = response.json().get("data", {})
    items = data.get("day_items") or []
    count = 0
    for item in items:
        title = item.get("title", "").strip()
        url = item.get("uri", "").strip()
        article_id = item.get("id")
        if not title or not url:
            continue
        summary, image = fetch_wallstreetcn_article_detail(article_id)
        if add_result("華爾街見聞", title, url, summary=summary, image_url=image):
            count += 1
        if count >= limit:
            break


def scrape_hket(limit=None):
    limit = limit or SOURCE_LIMITS["香港經濟日報"]
    soup = fetch_soup("https://www.hket.com/")
    count = 0
    for link in soup.select("a.listing-overlay"):
        title = link.get_text(strip=True)
        url = link.get("href", "").strip()
        article = enrich_article_metadata({"url": url, "summary": "", "image_url": ""})
        if add_result(
            "香港經濟日報",
            title,
            url,
            summary=article.get("summary", ""),
            image_url=article.get("image_url", ""),
        ):
            count += 1
        if count >= limit:
            break


def scrape_hkej(limit=None):
    limit = limit or SOURCE_LIMITS["信報財經"]
    soup = fetch_soup("https://www1.hkej.com/dailynews/finnews")
    count = 0
    for paragraph in soup.select("p.top_news_right_detail"):
        link = paragraph.find("a")
        if not link:
            continue
        title = link.get_text(strip=True)
        url = link.get("href", "").strip()
        if url.startswith("/"):
            url = "https://www1.hkej.com" + url
        article = enrich_article_metadata({"url": url, "summary": "", "image_url": ""})
        if add_result(
            "信報財經",
            title,
            url,
            summary=article.get("summary", ""),
            image_url=article.get("image_url", ""),
        ):
            count += 1
        if count >= limit:
            break


def build_digest_email(articles, date_str):
    items_html = []

    for article in articles:
        safe_url = html.escape(article["url"], quote=True)
        safe_title = html.escape(article["title"])
        safe_source = html.escape(article["source"])
        safe_summary = html.escape(article.get("summary") or "")

        image_url = article.get("image_url", "")
        if is_remote_image_url(image_url):
            safe_image = html.escape(image_url, quote=True)
            alt_text = safe_source
        else:
            data_uri = local_image_to_data_uri(image_url)
            safe_image = data_uri
            alt_text = safe_source

        image_cell = ""
        if safe_image:
            image_cell = (
                f'<td width="200" valign="top" style="padding-right:16px;">'
                f'<a href="{safe_url}">'
                f'<img src="{safe_image}" width="200" alt="{alt_text}" '
                f'style="display:block;max-width:200px;height:auto;border:0;">'
                f"</a></td>"
            )
        else:
            image_cell = (
                f'<td width="200" valign="top" style="padding-right:16px;">'
                f'<div style="width:200px;height:100px;background:#f2f2f2;'
                f'color:#666666;font-size:12px;text-align:center;line-height:100px;">'
                f"{safe_source}</div></td>"
            )

        items_html.append(
            f"""
            <table width="100%" cellpadding="0" cellspacing="0"
                   style="margin-bottom:24px;border-bottom:1px solid #eeeeee;">
              <tr>
                {image_cell}
                <td valign="top">
                  <p style="margin:0 0 4px;font-size:11px;color:#888888;text-transform:uppercase;letter-spacing:0.5px;">
                    {safe_source}
                  </p>
                  <p style="margin:0 0 8px;font-size:18px;font-weight:bold;line-height:1.35;">
                    <a href="{safe_url}" style="color:#111111;text-decoration:none;">{safe_title}</a>
                  </p>
                  <p style="margin:0;font-size:14px;color:#555555;line-height:1.5;">{safe_summary}</p>
                </td>
              </tr>
            </table>
            """
        )

    return f"""
    <html>
      <body style="font-family:Arial,Helvetica,sans-serif;max-width:760px;margin:0 auto;padding:24px;color:#111111;">
        <p style="text-align:center;color:#888888;font-size:12px;margin:0 0 8px;"></p>
        <h1 style="text-align:center;font-size:28px;font-weight:bold;margin:0 0 8px;">
          {html.escape(EMAIL_SUBJECT_PREFIX)} - {html.escape(date_str)}
        </h1>
        <p style="text-align:center;color:#888888;font-size:13px;margin:0 0 32px;">
          
        </p>
        {''.join(items_html)}
      </body>
    </html>
    """


scrapers = [
    ("Bloomberg", scrape_bloomberg),
    ("Reuters", scrape_reuters),
    ("Financial Times", scrape_financial_times),
    ("CNBC", scrape_cnbc),
    ("WallstreetCN", scrape_wallstreetcn),
    ("HKET", scrape_hket),
    ("HKEJ", scrape_hkej),
    ("Seeking Alpha", scrape_seeking_alpha),
]

for source_name, scraper in scrapers:
    if len(results) >= MAX_TOTAL:
        print(f"Skipped: {source_name} (reached max {MAX_TOTAL} headlines)")
        continue
    before = len(results)
    try:
        scraper()
        added = len(results) - before
        print(f"OK: {source_name} ({added} headlines)")
    except Exception as exc:
        print(f"FAILED: {source_name} - {exc}")

if not results:
    raise SystemExit("No headlines were scraped. Please check your network connection.")

articles = results[:MAX_TOTAL]
export_name = f"Daily News Headlines_{today}"

df = pd.DataFrame(articles)
df = df.rename(
    columns={
        "source": "Sources",
        "title": "News Headlines",
        "url": "Links",
        "summary": "Summary",
        "image_url": "Image URL",
    }
)
df.index += 1

csv_path = os.path.join(filePath, export_name + ".csv")
df.to_csv(csv_path, encoding="utf_8_sig")
print(f"Saved CSV: {csv_path} ({len(df)} headlines)")


def send_email_smtp(subject, html_body):
    if not EMAIL_USER or not EMAIL_PASS:
        raise RuntimeError("EMAIL_USER / EMAIL_PASS not set")
    if not EMAIL_TO:
        raise RuntimeError("EMAIL_TO not set")

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = EMAIL_FROM or EMAIL_USER
    message["To"] = EMAIL_TO
    message.attach(MIMEText(html_body, "html", "utf-8"))

    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.sendmail(message["From"], [addr.strip() for addr in EMAIL_TO.split(",")], message.as_string())


def send_email_outlook(subject, html_body):
    if not OUTLOOK_AVAILABLE:
        raise RuntimeError("Outlook / pywin32 is not available on this system")
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = EMAIL_TO
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Send()


subject = f"{EMAIL_SUBJECT_PREFIX} - {today}"
html_body = build_digest_email(articles, today)

try:
    if EMAIL_USER and EMAIL_PASS:
        send_email_smtp(subject, html_body)
        print(f"Email sent via SMTP to {EMAIL_TO}.")
    else:
        send_email_outlook(subject, html_body)
        print(f"Email sent via Outlook to {EMAIL_TO}.")
except Exception as exc:
    channel = "SMTP" if EMAIL_USER and EMAIL_PASS else "Outlook"
    print(f"Email not sent ({channel} error): {exc}")
    print("The CSV was saved successfully. Fix email settings and run again, or send the CSV manually.")
    raise SystemExit(1) from exc
