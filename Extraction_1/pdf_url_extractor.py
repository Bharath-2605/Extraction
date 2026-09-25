#!/usr/bin/env python3
"""Extract URLs from one or more PDFs and save them to Excel.

Processing is local only so PII never leaves the machine.

Each page is scanned for digital text, hyperlinks, and annotations first,
then QR codes, logos, image text, and small embedded chunks. OCR stays in
memory and never writes image files.

Each output row includes Filename, Filename Date, Last 4 Digits, Client ID,
Form Type, Page Number, Source, the QR image, the URL, and live Status.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import threading
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import pymupdf as fitz
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TEXT_FLAGS = (
    fitz.TEXT_PRESERVE_LIGATURES
    | fitz.TEXT_PRESERVE_WHITESPACE
    | fitz.TEXT_DEHYPHENATE
    | fitz.TEXT_PRESERVE_SPANS
)

SOURCE_RANK = {
    "qr": 50,
    "hyperlink": 40,
    "logo": 35,
    "text": 25,
    "image": 10,
}
MIN_IMAGE_SIDE = 12
PAGE_RENDER_ZOOM = 2.0
MAX_RENDER_SIDE = 1600
QR_RENDER_ZOOM = 3.2
QR_MAX_RENDER_SIDE = 3000
QR_CROP_MIN_SIDE = 480
MAX_PAGE_WORKERS = 6
MAX_VISION_WORKERS = 4
MAX_QR_IMAGES = 120
MAX_OCR_IMAGES = 8
SPARSE_TEXT_CHARS = 120
URL_COLUMNS = [
    "Filename",
    "Filename Date",
    "Last 4 Digits",
    "Client ID",
    "Form Type",
    "Page Number",
    "Source",
    "QR Code Attachment",
    "URL",
    "Status",
]
RESULTS_SHEET = "QR Results"
FILE_MERGE_COLUMNS = (
    "Filename",
    "Filename Date",
    "Last 4 Digits",
    "Client ID",
    "Form Type",
)
WORKING_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")
WORKING_FONT = Font(color="006100", bold=True)
FAILED_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")
FAILED_FONT = Font(color="9C0006", bold=True)

_VISION_LOCK = threading.Lock()
_VISION_POOL: ThreadPoolExecutor | None = None


def _vision_pool() -> ThreadPoolExecutor:
    global _VISION_POOL
    with _VISION_LOCK:
        if _VISION_POOL is None:
            _VISION_POOL = ThreadPoolExecutor(max_workers=2, thread_name_prefix="vision")
        return _VISION_POOL


TLDS = frozenset(
    {
        "aaa", "aarp", "abogado", "ac", "academy", "accountant", "accountants",
        "actor", "ad", "ads", "adult", "ae", "aero", "af", "afl", "ag", "agency",
        "ai", "aig", "airforce", "al", "am", "amsterdam", "analytics", "android",
        "ao", "app", "apple", "aq", "ar", "archi", "army", "arpa", "art", "as",
        "asia", "associates", "at", "attorney", "au", "auction", "audio", "auto",
        "autos", "aw", "aw", "ax", "az", "ba", "baby", "band", "bank", "bar",
        "barcelona", "bargains", "baseball", "basketball", "bauhaus", "bayern",
        "bb", "bbc", "bbva", "bd", "be", "beauty", "beer", "berlin", "best",
        "bet", "bf", "bg", "bh", "bi", "bible", "bid", "bike", "bingo", "bio",
        "biz", "bj", "black", "blackfriday", "blog", "blue", "bm", "bms", "bmw",
        "bn", "bnpparibas", "bo", "boats", "bond", "boo", "book", "booking",
        "bosch", "bostik", "boston", "bot", "boutique", "box", "br", "bradesco",
        "bridgestone", "broker", "brother", "brussels", "bs", "bt", "build",
        "builders", "business", "buy", "buzz", "bw", "by", "bz", "bzh", "ca",
        "cab", "cafe", "cal", "call", "cam", "camera", "camp", "cancerresearch",
        "canon", "capetown", "capital", "car", "cards", "care", "career",
        "careers", "cars", "casa", "case", "cash", "casino", "cat", "catering",
        "catholic", "cba", "cbn", "cbre", "cc", "cd", "center", "ceo", "cern",
        "cf", "cfa", "cfd", "cg", "ch", "chanel", "channel", "charity", "chase",
        "chat", "cheap", "christmas", "chrome", "church", "ci", "cipriani",
        "circle", "cisco", "citadel", "citi", "citic", "city", "ck", "cl",
        "claims", "cleaning", "click", "clinic", "clinique", "clothing",
        "cloud", "club", "cm", "cn", "co", "coach", "codes", "coffee", "college",
        "cologne", "com", "community", "company", "compare", "computer",
        "comsec", "condos", "construction", "consulting", "contact",
        "contractors", "cooking", "cool", "coop", "corsica", "country",
        "coupon", "coupons", "courses", "cpa", "cr", "credit", "creditcard",
        "creditunion", "cricket", "cruise", "cruises", "cu", "cuisinella", "cv",
        "cw", "cx", "cy", "cymru", "cyou", "cz", "dabur", "dad", "dance", "data",
        "date", "dating", "de", "deal", "dealer", "deals", "degree", "delivery",
        "dell", "deloitte", "delta", "democrat", "dental", "dentist", "desi",
        "design", "dev", "dhl", "diamonds", "diet", "digital", "direct",
        "directory", "discount", "discover", "dish", "diy", "dj", "dk", "dm",
        "dnp", "do", "docs", "doctor", "dog", "domains", "dot", "download",
        "drive", "dtv", "dubai", "duck", "dunlop", "dupont", "durban", "dvag",
        "dvr", "dz", "earth", "eat", "ec", "eco", "edeka", "edu", "education",
        "ee", "eg", "email", "emerck", "energy", "engineer", "engineering",
        "enterprises", "epson", "equipment", "er", "ericsson", "erni", "es",
        "esq", "estate", "et", "eu", "eurovision", "eus", "events", "exchange",
        "expert", "exposed", "express", "extraspace", "fage", "fail", "fairwinds",
        "faith", "family", "fan", "fans", "farm", "farmers", "fashion", "fast",
        "fedex", "feedback", "ferrari", "ferrero", "fi", "fidelity", "fido",
        "film", "final", "finance", "financial", "fire", "firestone", "firmdale",
        "fish", "fishing", "fit", "fitness", "fj", "fk", "flickr", "flights",
        "flir", "florist", "flowers", "fly", "fm", "fo", "foo", "food",
        "football", "ford", "forex", "forsale", "forum", "foundation", "fox",
        "fr", "free", "fresenius", "frl", "frogans", "frontdoor", "frontier",
        "ftr", "fujitsu", "fun", "fund", "furniture", "futbol", "fyi", "ga",
        "gal", "gallery", "gallo", "gallup", "game", "games", "gap", "garden",
        "gay", "gb", "gbiz", "gd", "gdn", "ge", "gea", "gent", "genting",
        "george", "gf", "gg", "ggee", "gh", "gh", "gi", "gift", "gifts", "gives",
        "giving", "gl", "glass", "gle", "global", "globo", "gm", "gmail", "gmbh",
        "gmo", "gmx", "gn", "godaddy", "gold", "goldpoint", "golf", "goo", "goodyear",
        "goog", "google", "gop", "got", "gov", "gp", "gq", "gr", "grainger",
        "graphics", "gratis", "green", "gripe", "grocery", "group", "gs", "gt",
        "gu", "guardian", "gucci", "guge", "guide", "guitars", "guru", "gw", "gy",
        "hair", "hamburg", "hangout", "haus", "hbo", "hdfc", "hdfcbank", "health",
        "healthcare", "help", "helsinki", "here", "hermes", "hiphop", "hisamitsu",
        "hitachi", "hiv", "hk", "hkt", "hm", "hn", "hockey", "holdings", "holiday",
        "homedepot", "homegoods", "homes", "homesense", "honda", "horse",
        "hospital", "host", "hosting", "hot", "hotels", "hotmail", "house", "how",
        "hr", "hsbc", "ht", "hu", "hughes", "hyatt", "hyundai", "ibm", "icbc",
        "ice", "icu", "id", "ie", "ieee", "ifm", "ikano", "il", "im", "imamat",
        "imdb", "immo", "immobilien", "in", "im", "inc", "industries", "infiniti",
        "info", "ing", "ink", "institute", "insurance", "insure", "int",
        "international", "intuit", "investments", "io", "ipiranga", "iq", "ir",
        "irish", "is", "ismaili", "ist", "istanbul", "it", "itau", "itv",
        "jaguar", "java", "jcb", "je", "jeep", "jetzt", "jewelry", "jio", "jll",
        "jm", "jmp", "jnj", "jo", "jobs", "joburg", "jot", "joy", "jp", "jpmorgan",
        "jprs", "juegos", "juniper", "kaufen", "kddi", "ke", "kerryhotels",
        "kerrylogistics", "kerryproperties", "kfh", "kg", "kh", "ki", "kia",
        "kids", "kim", "kinder", "kindle", "kitchen", "kiwi", "km", "kn", "koeln",
        "komatsu", "kosher", "kp", "kpmg", "kpn", "kr", "krd", "kred", "kuokgroup",
        "kw", "ky", "kyoto", "kz", "la", "lacaixa", "lamborghini", "lamer",
        "lancaster", "lancia", "land", "landrover", "lanxess", "lasalle", "lat",
        "latino", "latrobe", "law", "lawyer", "lb", "lc", "lds", "lease",
        "leclerc", "lefrak", "legal", "lego", "lexus", "lgbt", "li", "lidl",
        "life", "lifeinsurance", "lifestyle", "lighting", "like", "lilly",
        "limited", "limo", "lincoln", "link", "lipsy", "live", "living", "lk",
        "llc", "llp", "loan", "loans", "locker", "locus", "lol", "london",
        "lotte", "lotto", "love", "lpl", "lplfinancial", "lr", "ls", "lt", "ltd",
        "ltda", "lu", "lundbeck", "luxe", "luxury", "lv", "ly", "ma", "madrid",
        "maif", "maison", "makeup", "man", "management", "mango", "map",
        "market", "marketing", "markets", "marriott", "marshalls", "mattel",
        "mba", "mc", "mckinsey", "md", "me", "med", "media", "meet", "melbourne",
        "meme", "memorial", "men", "menu", "merckmsd", "mg", "mh", "miami",
        "microsoft", "mil", "mini", "mint", "mit", "mitsubishi", "mk", "ml",
        "mlb", "mls", "mm", "mma", "mn", "mo", "mobi", "mobile", "moda", "moe",
        "moi", "mom", "monash", "money", "monster", "mormon", "mortgage",
        "moscow", "moto", "motorcycles", "mov", "movie", "mp", "mq", "mr", "ms",
        "msd", "mt", "mtn", "mtr", "mu", "museum", "music", "mutual", "mv", "mw",
        "mx", "my", "mz", "na", "nab", "nagoya", "name", "natura", "navy", "nba",
        "nc", "ne", "nec", "net", "netbank", "netflix", "network", "neustar",
        "new", "news", "next", "nextdirect", "nexus", "nf", "nfl", "ng", "ngo",
        "nhk", "ni", "nico", "nike", "nikon", "ninja", "nissan", "nissay", "nl",
        "no", "nokia", "norton", "now", "nowruz", "nowtv", "np", "nr", "nra",
        "nrw", "ntt", "nu", "nyc", "nz", "obi", "observer", "office", "okinawa",
        "olayan", "olayangroup", "oldnavy", "ollo", "om", "omega", "one", "ong",
        "onl", "online", "ooo", "open", "oracle", "orange", "org", "organic",
        "origins", "osaka", "otsuka", "ott", "ovh", "pa", "page", "panasonic",
        "paris", "pars", "partners", "parts", "party", "passagens", "pay",
        "pccw", "pe", "pet", "pf", "pfizer", "pg", "ph", "pharmacy", "phd",
        "philips", "phone", "photo", "photography", "photos", "phyto", "pics",
        "pictet", "pictures", "pid", "pin", "ping", "pink", "pioneer", "pizza",
        "pk", "pl", "place", "play", "playstation", "plumbing", "plus", "pm",
        "pn", "pnc", "pohl", "poker", "politie", "porn", "post", "pr", "pramerica",
        "praxi", "press", "prime", "pro", "prod", "productions", "prof",
        "progressive", "promo", "properties", "property", "protection", "pru",
        "prudential", "ps", "pt", "pub", "pw", "pwc", "py", "qa", "qpon",
        "quebec", "quest", "racing", "radio", "re", "read", "realestate",
        "realtor", "realty", "recipes", "red", "redstone", "redumbrella",
        "rehab", "reise", "reisen", "reit", "reliance", "ren", "rent", "rentals",
        "repair", "report", "republican", "rest", "restaurant", "review",
        "reviews", "rexroth", "rich", "richardli", "ricoh", "ril", "rio", "rip",
        "ro", "rocks", "rodeo", "rogers", "room", "rs", "rsvp", "ru", "rugby",
        "ruhr", "run", "rw", "rwe", "ryukyu", "sa", "saarland", "safe", "safety",
        "sakura", "sale", "salon", "samsclub", "samsung", "sandvik",
        "sandvikcoromant", "sanofi", "sap", "sarl", "sas", "save", "saxo", "sb",
        "sbi", "sbs", "sc", "sca", "scb", "schaeffler", "schmidt", "scholarships",
        "school", "schule", "schwarz", "science", "scot", "sd", "se", "search",
        "seat", "secure", "security", "seek", "select", "sener", "services",
        "ses", "seven", "sew", "sex", "sexy", "sfr", "sg", "sh", "shangrila",
        "sharp", "shaw", "shell", "shia", "shiksha", "shoes", "shop", "shopping",
        "shouji", "show", "showtime", "si", "silk", "sina", "singles", "site",
        "sj", "sk", "ski", "skin", "sky", "skype", "sl", "sling", "sm", "smart",
        "smile", "sn", "sncf", "so", "soccer", "social", "softbank", "software",
        "sohu", "solar", "solutions", "song", "sony", "soy", "spa", "space",
        "sport", "spot", "sr", "srl", "ss", "st", "stada", "staples", "star",
        "statebank", "statefarm", "stc", "stcgroup", "stockholm", "storage",
        "store", "stream", "studio", "study", "style", "su", "sucks", "supplies",
        "supply", "support", "surf", "surgery", "suzuki", "sv", "swatch",
        "swiss", "sx", "sy", "sydney", "systems", "sz", "tab", "taipei", "talk",
        "taobao", "target", "tatamotors", "tatar", "tattoo", "tax", "taxi", "tc",
        "tci", "td", "tdk", "team", "tech", "technology", "tel", "temasek",
        "tennis", "teva", "tf", "tg", "th", "thd", "theater", "theatre", "tiaa",
        "tickets", "tienda", "tips", "tires", "tirol", "tj", "tjmaxx", "tjx",
        "tk", "tkmaxx", "tl", "tm", "tmall", "tn", "to", "today", "tokyo",
        "tools", "top", "toray", "toshiba", "total", "tours", "town", "toyota",
        "toys", "tr", "trade", "trading", "training", "travel", "travelers",
        "travelersinsurance", "trust", "trv", "tt", "tube", "tui", "tunes",
        "tunes", "tushu", "tv", "tvs", "tw", "tz", "ua", "ubank", "ubs", "ug",
        "uk", "unicom", "university", "uno", "uol", "ups", "us", "uy", "uz",
        "va", "vacations", "vana", "vanguard", "vc", "ve", "vegas", "ventures",
        "verisign", "versicherung", "vet", "vg", "vi", "viajes", "video",
        "vig", "viking", "villas", "vin", "vip", "virgin", "visa", "vision",
        "viva", "vivo", "vlaanderen", "vn", "vodka", "volkswagen", "volvo",
        "vote", "voting", "voto", "voyage", "vu", "vuelos", "wales", "walmart",
        "walter", "wang", "wanggou", "watch", "watches", "weather",
        "weatherchannel", "webcam", "weber", "website", "wed", "wedding",
        "weibo", "weir", "wf", "whoswho", "wien", "wiki", "williamhill", "win",
        "windows", "wine", "winners", "wme", "wolterskluwer", "woodside", "work",
        "works", "world", "wow", "ws", "wtc", "wtf", "xbox", "xerox", "xfinity",
        "xihuan", "xin", "xxx", "xyz", "yachts", "yahoo", "yamaxun", "yandex",
        "ye", "yodobashi", "yoga", "yokohama", "you", "youtube", "yt", "yun",
        "za", "zappos", "zara", "zero", "zip", "zm", "zone", "zuerich", "zw",
    }
)

# Multi-part country suffixes such as co.uk so the host is not treated as "co".
MULTI_PART_TLDS = frozenset(
    {
        "ac.uk", "co.uk", "gov.uk", "ac.jp", "co.jp", "or.jp", "ne.jp",
        "com.au", "net.au", "org.au", "edu.au", "gov.au", "co.nz", "org.nz",
        "co.in", "net.in", "org.in", "gov.in", "ac.in", "res.in", "co.za",
        "org.za", "gov.za", "com.br", "org.br", "gov.br", "com.mx", "org.mx",
        "com.cn", "net.cn", "org.cn", "gov.cn", "com.hk", "com.sg", "com.tw",
        "co.kr", "co.id", "com.ar", "com.tr", "com.ua", "com.pl",
    }
)


SCHEME_URL_RE = re.compile(
    r"(?:https?|ftp)://[^\s<>\[\]()\"']+",
    re.IGNORECASE,
)
WWW_URL_RE = re.compile(
    r"(?<![\w./])www\.(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,}"
    r"(?::\d{2,5})?(?:/[^\s<>\[\]()\"']*)?",
    re.IGNORECASE,
)
BARE_DOMAIN_RE = re.compile(
    r"(?<![@\w./-])(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,}"
    r"(?::\d{2,5})?(?:/[^\s<>\[\]()\"']*)?",
    re.IGNORECASE,
)
TRAILING_PUNCT_RE = re.compile(r"[.,;:!?)]\}>\"'`“”‘’«»′″\]]+$")
TRAILING_NUMERIC_JUNK_RE = re.compile(r"(?:[.,]\d+%?)+$")
QUOTE_RE = re.compile(r"[\"'`“”‘’«»′″]")
APOSTROPHE_JOIN_RE = re.compile(r"(?<=[A-Za-z0-9])[\s]*[''`‘’]+[\s]*(?=[A-Za-z0-9])")
SCHEME_SPACE_RE = re.compile(r"(?i)\b(https?|ftp|upi)\s*:\s*/\s*/\s*")
WWW_SPACE_RE = re.compile(r"(?i)\bwww\s*\.\s*")
DOT_SPACE_RE = re.compile(r"(?<=[A-Za-z0-9])\s*\.\s*([A-Za-z0-9]+)")
SLASH_SPACE_RE = re.compile(r"(?<=[A-Za-z0-9.])\s*/\s*(?=[A-Za-z0-9])")
INNER_URL_SPACE_RE = re.compile(r"\s+")
GLUED_AFTER_TLD_RE = re.compile(
    r"(?P<head>(?:https?://|ftp://)?(?:www\.)?(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,})"
    r"\.(?P<tail>[A-Za-z][A-Za-z0-9-]*)"
)
EMAIL_PREFIX_RE = re.compile(r"[A-Z0-9._%+-]+@$", re.IGNORECASE)
AMBIGUOUS_TLDS = frozenset(
    {
        "al", "am", "an", "as", "at", "be", "by", "do", "et", "go", "he", "hi",
        "id", "ie", "if", "in", "is", "it", "me", "my", "no", "ok", "on", "or",
        "pm", "re", "so", "st", "th", "to", "up", "us", "vs", "we",
    }
)

FUNCTION_WORDS = frozenset(
    {
        "a", "al", "an", "and", "as", "at", "by", "et", "for", "from", "if",
        "in", "is", "it", "no", "of", "on", "or", "our", "so", "the", "this",
        "that", "to", "us", "we", "with", "your",
    }
)

SENTENCE_WORDS = frozenset(
    {
        "about", "above", "account", "accounts", "additional", "address",
        "after", "all", "also", "any", "apply", "available", "before", "below",
        "between", "both", "but", "call", "can", "card", "cards", "click",
        "contact", "customer", "customers", "details", "during", "each",
        "email", "every", "find", "following", "further", "get", "go", "have",
        "here", "how", "important", "including", "information", "into", "its",
        "just", "learn", "link", "login", "mail", "make", "more", "most", "note", "now",
        "online", "only", "open", "other", "out", "over", "page", "please", "policy",
        "privacy", "questions", "read", "rewards", "see", "service", "services",
        "should", "site", "some", "step", "such", "take", "terms", "than", "then",
        "there", "these", "through", "under", "use", "using", "view", "visit",
        "was", "website", "were", "what", "when", "where", "which", "who",
        "why", "will", "you",
    }
)

TRUSTED_TLDS = frozenset(
    {
        "com", "org", "net", "edu", "gov", "mil", "int", "io", "ai", "biz",
        "co", "app", "dev", "xyz", "one", "bank",
    }
)

# Words that follow a finished link and must not stay attached to it.
TAIL_WORDS = SENTENCE_WORDS | FUNCTION_WORDS | frozenset(
    {
        "amount", "balance", "bill", "due", "order", "payment", "rupee",
        "rupees", "steps", "today", "total",
    }
)
FILE_EXTS = frozenset(
    {
        "asp", "aspx", "csv", "css", "gif", "htm", "html", "jpeg", "jpg", "js",
        "json", "pdf", "php", "png", "svg", "txt", "webp", "xml", "zip",
    }
)


_FORM_STOP = frozenset(
    {
        "PDF", "FAX", "TEL", "TAX", "USA", "THE", "AND", "FOR", "WWW", "COM", "NET",
        "ORG", "INC", "LLC", "LTD", "BOX", "PAY", "DUE", "APR", "MAY", "JUN", "JUL",
        "AUG", "SEP", "OCT", "NOV", "DEC", "JAN", "FEB", "MAR", "SUN", "MON", "TUE",
        "WED", "THU", "FRI", "SAT", "SYF", "URL", "IMG", "PNG", "JPG", "EOF", "ALL",
        "NOT", "YES", "TBD", "ETA", "GST", "EIN", "SSN", "DOB", "ZIP", "APT", "STE",
        "AVE", "QR", "PMT", "ACH", "ABA", "PO",
    }
)


def filename_meta(file_name: str) -> tuple[str, str]:
    """Date and account last-4 from names such as 09-02-26 5560531101016173.pdf."""
    stem = Path(str(file_name or "")).stem.strip()
    match = re.search(r"(\d{2}-\d{2}-\d{2})(\d{2})?", stem)
    if not match:
        digits = re.sub(r"\D", "", stem)
        return "", digits[-4:] if len(digits) >= 4 else digits
    year2 = match.group(1)
    extra = match.group(2) or ""
    after = stem[match.end():]
    if extra and (after == "" or not after[:1].isdigit()):
        file_date = year2 + extra
        rest = stem[: match.start()] + after
    else:
        file_date = year2
        rest = stem[: match.start()] + extra + after
    digits = re.sub(r"\D", "", rest)
    last4 = digits[-4:] if len(digits) >= 4 else digits
    return file_date, last4


def _valid_form(token: str) -> bool:
    """Form type is at most 3 characters, such as KTJ or A2J. Digits may sit inside it."""
    token = str(token or "").upper()
    if len(token) != 3 or token in _FORM_STOP:
        return False
    return any(char.isalpha() for char in token)


def _ids_from_footer_text(text: str) -> tuple[str, str]:
    raw = str(text or "")
    # 5433 3662 A2J, or 6709 / HJJ. The code is the 3 characters after the client id.
    for match in re.finditer(
        r"(?<!\d)(\d{4})(?!\d)(?:\s+(\d{4})(?!\d))?(?:\s*[/|]\s*|\s+)([A-Z0-9]{3})(?![A-Z0-9])",
        raw,
    ):
        client_tok = match.group(1)
        form_tok = match.group(3)
        if client_tok.startswith(("19", "20")):
            continue
        if _valid_form(form_tok):
            return client_tok, form_tok.upper()
    digits = [match.group(1) for match in re.finditer(r"(?<!\d)(\d{4})(?!\d)", raw)]
    client = ""
    for token in digits:
        if not token.startswith(("19", "20")):
            client = token
            break
    if not client and digits:
        client = digits[0]
    forms = [
        match.group(1).upper()
        for match in re.finditer(r"(?<![A-Z0-9])([A-Z0-9]{3})(?![A-Z0-9])", raw)
        if _valid_form(match.group(1))
    ]
    # The footer code is the lowest one. An earlier word such as OFF is not the form.
    form = forms[-1] if forms else ""
    return client, form


def _footer_ids(page: fitz.Page, page_text: str) -> tuple[str, str]:
    """Client ID is the left-footer 4-digit code. Form type is the 3-character footer code."""
    height = float(page.rect.height)
    width = float(page.rect.width)
    top = height * 0.78
    left_text = ""
    full_text = ""
    try:
        left_text = page.get_text("text", clip=fitz.Rect(0, top, width * 0.7, height)) or ""
        full_text = page.get_text("text", clip=fitz.Rect(0, top, width, height)) or ""
    except Exception:
        left_text = ""
        full_text = ""
    client, form_left = _ids_from_footer_text(left_text)
    _client_full, form_full = _ids_from_footer_text(full_text)
    form = form_left or form_full
    if not client:
        client = _client_full
    if not (client and form):
        lines = [line.strip() for line in str(page_text or "").splitlines() if line.strip()]
        client2, form2 = _ids_from_footer_text("\n".join(lines[-5:]))
        client = client or client2
        form = form or form2
    if client and form:
        return client, form
    if len(re.sub(r"\s+", "", left_text + full_text)) >= 8:
        return client, form
    try:
        from vision_scan import recognize_text

        image = render_region(page, fitz.Rect(0, top, width, height), 2.0)
        client3, form3 = _ids_from_footer_text(recognize_text(image, True))
        client = client or client3
        form = form or form3
    except Exception:
        pass
    return client, form


def footer_for_pdf(pdf_path: Path | str, password: str | None = None) -> tuple[str, str]:
    """Client ID and form type are the same on every page, so read the last page once."""
    document = _open_document(Path(pdf_path), password)
    try:
        if document.page_count < 1:
            return "", ""
        return _footer_ids(document.load_page(document.page_count - 1), "")
    finally:
        document.close()


def collect_pdf_paths(inputs: list[str]) -> list[Path]:
    paths: list[Path] = []
    seen: set[Path] = set()
    for raw in inputs:
        item = Path(raw).expanduser()
        if item.is_dir():
            candidates = sorted(item.rglob("*.pdf"))
        elif item.is_file():
            candidates = [item]
        else:
            raise FileNotFoundError(f"PDF not found: {item}")
        for path in candidates:
            resolved = path.resolve()
            if resolved.suffix.lower() != ".pdf":
                continue
            if resolved in seen:
                continue
            seen.add(resolved)
            paths.append(path)
    if not paths:
        raise FileNotFoundError("No PDF files were found in the given paths.")
    return paths


def _split_glued_after_tld(text: str) -> str:
    """www.syf.com.Step → www.syf.com. Step so the sentence is not part of the host."""

    def repl(match: re.Match) -> str:
        head = match.group("head")
        tail = match.group("tail")
        host = re.sub(r"^(?:https?|ftp)://", "", head, flags=re.IGNORECASE)
        labels = [part for part in host.split(".") if part]
        if len(labels) < 2:
            return match.group(0)
        tld = labels[-1].lower()
        tail_l = tail.lower()
        multi = ".".join(labels[-2:]).lower()
        if tld not in TLDS and multi not in MULTI_PART_TLDS:
            return match.group(0)
        if tail_l in TLDS or f"{tld}.{tail_l}" in MULTI_PART_TLDS:
            return match.group(0)
        return f"{head}. {tail}"

    return GLUED_AFTER_TLD_RE.sub(repl, text)


def sanitize_url_source(text: str) -> str:
    """Drop quotes, apostrophes, and extra spaces that split a link.

    DICK'S.com/ScoreCardTerms → DICKS.com/ScoreCardTerms
    DICKS . com / ScoreCard → DICKS.com/ScoreCard
    """
    cleaned = APOSTROPHE_JOIN_RE.sub("", str(text or ""))
    cleaned = QUOTE_RE.sub("", cleaned)
    cleaned = SCHEME_SPACE_RE.sub(lambda match: f"{match.group(1)}://", cleaned)
    cleaned = WWW_SPACE_RE.sub("www.", cleaned)

    def join_dot(match: re.Match) -> str:
        right = match.group(1)
        right_l = right.lower()
        start = match.start()
        index = start - 1
        source = match.string
        while index >= 0 and (source[index].isalnum() or source[index] == "-"):
            index -= 1
        left = source[index + 1 : start].lower()
        # Keep real suffixes such as co.uk and .com. Do not split DSG.SYF.COM.
        if f"{left}.{right_l}" in MULTI_PART_TLDS or right_l in TLDS:
            return f".{right}"
        # "www.syf.com. Step 1" and "www.syf.com.Step" are the domain plus the next sentence.
        if left in TLDS or right_l in SENTENCE_WORDS or right_l in FUNCTION_WORDS:
            return f". {right}"
        left_is_word = left in FUNCTION_WORDS or left in SENTENCE_WORDS
        right_is_word = right_l in FUNCTION_WORDS or right_l in SENTENCE_WORDS
        if left_is_word and (right_is_word or right_l not in TLDS):
            return match.group(0)
        return f".{right}"

    cleaned = DOT_SPACE_RE.sub(join_dot, cleaned)
    cleaned = SLASH_SPACE_RE.sub("/", cleaned)
    return _split_glued_after_tld(cleaned)


def _trim_glued_sentence_suffix(url: str) -> str:
    """www.syf.com.Step / www.syf.com. Step → www.syf.com"""
    if not url or url.lower().startswith(("upi://", "mailto:", "javascript:")):
        return url
    if " " in url:
        url = url.split()[0]
    scheme = ""
    rest = url
    lowered = url.lower()
    for prefix in ("https://", "http://", "ftp://"):
        if lowered.startswith(prefix):
            scheme = url[: len(prefix)]
            rest = url[len(prefix) :]
            break
    hostport, slash, path = rest.partition("/")
    extra = ""
    if "?" in hostport:
        hostport, mark, tail = hostport.partition("?")
        extra = mark + tail
    elif "#" in hostport:
        hostport, mark, tail = hostport.partition("#")
        extra = mark + tail
    host, colon, port = hostport.partition(":")
    if colon and not str(port).isdigit():
        colon, port = "", ""
    if slash:
        return scheme + host + ((colon + port) if colon else "") + slash + path + extra
    labels = [part for part in host.split(".") if part]
    while len(labels) >= 3:
        tld = labels[-1].lower()
        prev = labels[-2].lower()
        last_raw = labels[-1]
        multi = ".".join(labels[-2:]).lower()
        valid_tld = tld in TLDS or multi in MULTI_PART_TLDS
        glued = prev in TLDS and multi not in MULTI_PART_TLDS and (
            (last_raw[:1].isupper() and last_raw[1:].islower())
            or tld in SENTENCE_WORDS
            or tld in FUNCTION_WORDS
            or tld not in TLDS
        )
        if valid_tld and not glued:
            break
        labels.pop()
    new_host = ".".join(labels)
    if not new_host:
        return url
    return scheme + new_host + ((colon + port) if colon else "") + extra


def _peel_dot_slash_sentence(url: str) -> str:
    """www.syf.com./Step → www.syf.com. A real slash path such as /ScoreCard stays."""
    match = re.search(r"\./([A-Za-z][A-Za-z0-9-]*)/?$", url)
    if not match:
        return url
    word = match.group(1)
    lowered = word.lower()
    sentence = (
        lowered in SENTENCE_WORDS
        or lowered in FUNCTION_WORDS
        or (word[:1].isupper() and lowered not in TLDS)
    )
    if not sentence:
        return url
    return url[: match.start()].rstrip(".")


_CAMEL_TAIL_RE = re.compile(r"^(?P<head>[a-z0-9][a-z0-9-]*)(?P<tail>[A-Z][A-Za-z]{2,})$")
_DOT_TAIL_RE = re.compile(r"\.([A-Za-z][A-Za-z-]{1,})$")


def _peel_added_words(url: str) -> str:
    """Remove a sentence stuck to the end of a link. Real paths and file names stay."""
    if not url or url.lower().startswith(("upi://", "mailto:", "javascript:")):
        return url
    scheme = ""
    rest = url
    lowered = url.lower()
    for prefix in ("https://", "http://", "ftp://"):
        if lowered.startswith(prefix):
            scheme = url[: len(prefix)]
            rest = url[len(prefix) :]
            break
    rest, mark, query = rest.partition("?")
    rest, hashmark, fragment = rest.partition("#")
    host, slash, path = rest.partition("/")

    def peel_dot(text: str) -> str:
        match = _DOT_TAIL_RE.search(text)
        if not match:
            return text
        word = match.group(1)
        word_l = word.lower()
        if word_l in FILE_EXTS:
            return text
        titled = word[:1].isupper() and (len(word) == 1 or word[1:].islower())
        if word_l in TAIL_WORDS or (titled and word_l not in TLDS):
            return text[: match.start()]
        return text

    labels = [part for part in host.split(".") if part]
    while len(labels) >= 3:
        last = labels[-1].lower()
        prev = labels[-2].lower()
        multi = ".".join(labels[-2:]).lower()
        if prev in TLDS and multi not in MULTI_PART_TLDS and last in TAIL_WORDS:
            labels.pop()
            continue
        break
    host = ".".join(labels)
    if slash and path:
        segments = path.split("/")
        last_seg = segments[-1]
        match = _CAMEL_TAIL_RE.match(last_seg)
        if match and match.group("tail").lower() in TAIL_WORDS:
            segments[-1] = match.group("head")
            path = "/".join(segments)
        path = peel_dot(path)
        rest = host + slash + path
    else:
        rest = peel_dot(host)
    if hashmark:
        rest += hashmark + fragment
    if mark:
        rest += mark + peel_dot(query)
    return scheme + rest


def strip_trailing_punctuation(value: str) -> str:
    cleaned = sanitize_url_source(value).strip()
    if " " in cleaned and cleaned.lower().startswith(("http://", "https://", "ftp://", "www.", "upi://")):
        cleaned = cleaned.split()[0]
    while cleaned:
        next_value = sanitize_url_source(cleaned).strip()
        if " " in next_value and next_value.lower().startswith(("http://", "https://", "ftp://", "www.", "upi://")):
            next_value = next_value.split()[0]
        if "?" not in next_value and "#" not in next_value:
            next_value = TRAILING_NUMERIC_JUNK_RE.sub("", next_value)
        next_value = TRAILING_PUNCT_RE.sub("", next_value).rstrip(".,;:!?)]}>\"'")
        next_value = _peel_dot_slash_sentence(next_value)
        next_value = _trim_glued_sentence_suffix(next_value)
        next_value = _peel_added_words(next_value)
        if next_value.endswith("/") and not next_value.lower().startswith(("http://", "https://", "ftp://")):
            next_value = next_value.rstrip("/")
        if next_value == cleaned:
            break
        cleaned = next_value
    return cleaned


def domain_and_tld(host: str) -> tuple[str, str] | None:
    host = host.strip(".").lower()
    if ":" in host:
        host = host.split(":", 1)[0]
    labels = [part for part in host.split(".") if part]
    if len(labels) < 2:
        return None
    last_two = ".".join(labels[-2:])
    if last_two in MULTI_PART_TLDS and len(labels) >= 3:
        return host, last_two
    tld = labels[-1]
    if tld not in TLDS:
        return None
    return host, tld


def _host_and_path(url: str) -> tuple[str, str]:
    text = str(url or "").strip()
    lowered = text.lower()
    for prefix in ("https://", "http://", "ftp://"):
        if lowered.startswith(prefix):
            text = text[len(prefix):]
            break
    host_part, _, remainder = text.partition("/")
    if "@" in host_part:
        host_part = host_part.rsplit("@", 1)[-1]
    host = host_part.split(":")[0].strip().strip(".")
    path = remainder.split("?", 1)[0].split("#", 1)[0]
    return host, path


def is_noise_url(url: str) -> bool:
    """True for sentence fragments that look like domains, e.g. information.contact."""
    host, path = _host_and_path(url)
    if not host:
        return True
    if re.match(r"^\d{1,3}(?:\.\d{1,3}){3}$", host):
        return False
    had_www = host.lower().startswith("www.")
    core = host[4:] if had_www else host
    parsed = domain_and_tld(core)
    if parsed is None:
        return True
    full_host, tld = parsed
    labels = [part for part in full_host.split(".") if part]
    if len(labels) < 2:
        return True
    sld = labels[-3] if tld in MULTI_PART_TLDS and len(labels) >= 3 else labels[-2]
    if sld in FUNCTION_WORDS:
        return True
    if had_www or bool(path.strip("/")):
        return False
    if tld in TRUSTED_TLDS:
        return False
    # Bare fragments such as linea.de. A real link has www, a path, or a scheme.
    lowered_url = str(url or "").lower()
    has_scheme = lowered_url.startswith(("http://", "https://", "ftp://"))
    if (
        not had_www
        and not has_scheme
        and not path.strip("/")
        and "." not in tld
        and len(tld) == 2
        and sld.isalpha()
        and "-" not in sld
    ):
        return True
    core_noise = (
        sld in SENTENCE_WORDS
        or tld in SENTENCE_WORDS
        or (len(sld) <= 2 and sld.isalpha())
    )
    if not core_noise:
        return False
    extra_count = len(labels) - (3 if tld in MULTI_PART_TLDS else 2)
    if extra_count <= 0:
        return True
    extra = labels[:extra_count]
    if all(
        part in SENTENCE_WORDS or part in FUNCTION_WORDS or (len(part) <= 2 and part.isalpha())
        for part in extra
    ):
        return True
    return False


def looks_like_url(value: str, source_text: str, start: int) -> bool:
    url = strip_trailing_punctuation(value)
    if not url or len(url) < 4 or len(url) > 250:
        return False
    lowered = url.lower()
    if lowered.startswith(("mailto:", "javascript:", "tel:")):
        return False
    if EMAIL_PREFIX_RE.search(source_text[:start]):
        return False
    if lowered.startswith(("http://", "https://", "ftp://")):
        remainder = re.split(r"://", url, maxsplit=1)[1]
        host = remainder.split("/", 1)[0]
        if re.match(r"\d{1,3}(?:\.\d{1,3}){3}", host.split(":")[0]):
            return True
        if domain_and_tld(host) is None:
            return False
        return not is_noise_url(url)

    host_raw = url.split("/", 1)[0]
    labels_raw = host_raw.split(":")[0].split(".")
    tld_raw = labels_raw[-1] if labels_raw else ""
    if tld_raw[:1].isupper() and tld_raw[1:].islower():
        return False
    if len(labels_raw) == 2:
        left, right = labels_raw[0], labels_raw[1]
        if left[:1].isupper() and left[1:].islower() and right.isupper() and len(right) >= 3:
            return False
    if lowered.startswith("www."):
        return domain_and_tld(host_raw) is not None and not is_noise_url(url)
    parsed = domain_and_tld(host_raw)
    if parsed is None:
        return False
    _full_host, tld = parsed
    if tld in {"py", "java", "c", "h", "json", "xml", "csv", "txt", "log"}:
        return False
    if tld in AMBIGUOUS_TLDS:
        return False
    return not is_noise_url(url)


def extract_url_occurrences(text: str) -> list[tuple[str, int]]:
    """Return every complete URL match, including repeats at different positions.

    A host like www.airtel.in is not counted when it is only the start of a
    longer path URL such as www.airtel.in/business/login/Page.
    """
    if not text:
        return []
    text = sanitize_url_source(text)
    raw: list[tuple[int, int, str]] = []
    for regex in (SCHEME_URL_RE, WWW_URL_RE, BARE_DOMAIN_RE):
        for match in regex.finditer(text):
            cleaned = strip_trailing_punctuation(match.group(0))
            if not looks_like_url(cleaned, text, match.start()):
                continue
            start = match.start()
            rest = text[start + len(cleaned) :]
            if re.match(r"^/[A-Za-z0-9._~%-]", rest):
                continue
            raw.append((start, start + len(cleaned), cleaned))
    raw.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    kept: list[tuple[str, int, int]] = []
    last_end = -1
    for start, end, cleaned in raw:
        if start < last_end:
            continue
        kept.append((cleaned, start, end))
        last_end = end
    filtered: list[tuple[str, int]] = []
    for url, start, end in kept:
        if any(
            start >= other_start and end <= other_end and (start, end) != (other_start, other_end)
            for _other, other_start, other_end in kept
        ):
            continue
        filtered.append((url, start))
    return filtered


def extract_urls_from_text(text: str) -> list[str]:
    return [url for url, _start in extract_url_occurrences(text)]


def span_gap_is_space(previous: dict, current: dict) -> bool:
    prev_text = previous.get("text") or ""
    cur_text = current.get("text") or ""
    if prev_text.endswith(("'", "’", "`")) or cur_text.startswith(("'", "’", "`")):
        return False
    prev_x1 = previous["bbox"][2]
    cur_x0 = current["bbox"][0]
    size = current.get("size") or previous.get("size") or 10
    return (cur_x0 - prev_x1) > (size * 0.18)


def reconstruct_page_text(page: fitz.Page) -> tuple[str, list[dict]]:
    page_rect = page.rect
    payload = page.get_text("dict", flags=TEXT_FLAGS, clip=page_rect)
    spans_out: list[dict] = []
    lines: list[str] = []
    for block in payload.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = [span for span in line.get("spans", []) if (span.get("text") or "")]
            if not spans:
                continue
            pieces = [spans[0]["text"]]
            for previous, current in zip(spans, spans[1:]):
                prev_text = previous.get("text") or ""
                cur_text = current.get("text") or ""
                needs_space = span_gap_is_space(previous, current)
                if prev_text.rstrip().endswith(".") and cur_text[:1].isupper():
                    needs_space = True
                if needs_space:
                    pieces.append(" ")
                pieces.append(current["text"])
            line_text = "".join(pieces)
            lines.append(line_text)
            for span in spans:
                spans_out.append(
                    {
                        "text": span["text"],
                        "size": float(span.get("size") or 0),
                        "bbox": tuple(span["bbox"]),
                    }
                )
    rebuilt = "\n".join(lines).strip()
    fallback = page.get_text("text", flags=TEXT_FLAGS, clip=page_rect).strip()
    if len(fallback) > len(rebuilt):
        rebuilt = fallback
    raw_text = _rawdict_text(page, page_rect)
    if raw_text and len(raw_text) > len(rebuilt):
        rebuilt = raw_text
    return rebuilt, spans_out


def _rawdict_text(page: fitz.Page, page_rect: fitz.Rect) -> str:
    """Pull extra character streams (vector text) without using get_drawings()."""
    try:
        payload = page.get_text("rawdict", flags=TEXT_FLAGS, clip=page_rect)
    except Exception:
        return ""
    pieces: list[str] = []
    for block in payload.get("blocks", []) or []:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []) or []:
            line_chars: list[str] = []
            for span in line.get("spans", []) or []:
                chars = span.get("chars") or []
                if chars:
                    line_chars.append("".join(str(ch.get("c") or "") for ch in chars))
                elif span.get("text"):
                    line_chars.append(str(span["text"]))
            if line_chars:
                pieces.append("".join(line_chars))
    return "\n".join(pieces).strip()


def location_for_bbox(bbox: tuple[float, float, float, float], page_height: float) -> str:
    y0, y1 = bbox[1], bbox[3]
    header_cut = page_height * 0.12
    footer_cut = page_height * 0.88
    if y1 <= header_cut:
        return "header"
    if y0 >= footer_cut:
        return "footer"
    return "body"


def font_size_for_url(url: str, spans: list[dict], occurrence: int = 0) -> float | None:
    needle = url.lower()
    hits = 0
    for span in spans:
        if needle in span["text"].lower():
            if hits == occurrence:
                return round(span["size"], 2)
            hits += 1
    compact_url = needle.replace(" ", "")
    compact_spans = "".join(span["text"] for span in spans).lower().replace(" ", "")
    if compact_url in compact_spans:
        sizes = [span["size"] for span in spans if span["size"] > 0]
        if sizes:
            return round(min(sizes), 2)
    return None


def location_for_url(url: str, spans: list[dict], page_height: float, occurrence: int = 0) -> str:
    needle = url.lower()
    hits = 0
    for span in spans:
        if needle in span["text"].lower():
            if hits == occurrence:
                return location_for_bbox(span["bbox"], page_height)
            hits += 1
    return "body"


def extract_hyperlink_uris(page: fitz.Page) -> list[dict]:
    results: list[dict] = []
    page_height = float(page.rect.height)
    for link in page.get_links():
        uri = link.get("uri")
        if not uri:
            continue
        cleaned = strip_trailing_punctuation(str(uri))
        rect = link.get("from")
        if rect is not None:
            bbox = (float(rect.x0), float(rect.y0), float(rect.x1), float(rect.y1))
            location = location_for_bbox(bbox, page_height)
        else:
            location = "body"
        results.append({"url": cleaned, "location": location, "source": "hyperlink"})
    return results


def ocr_text_variants(text: str) -> list[str]:
    if not text:
        return []
    text = sanitize_url_source(text)
    variants = [text]
    dotted = re.sub(r"https?\s*:\s*//", lambda match: re.sub(r"\s+", "", match.group(0)), text, flags=re.IGNORECASE)
    dotted = re.sub(r"\bwww\s*\.\s*", "www.", dotted, flags=re.IGNORECASE)
    if dotted not in variants:
        variants.append(dotted)
    return variants


def extract_urls_from_any_text(text: str) -> list[str]:
    if not str(text or "").strip():
        return []
    for variant in ocr_text_variants(text):
        found = extract_urls_from_text(variant)
        if found:
            return found
    found: list[str] = []
    for line in str(text).splitlines():
        line = line.strip()
        if line:
            found.extend(extract_urls_from_text(line))
    if found:
        return found
    for part in re.split(r"[\s|;]+", text):
        if len(part.strip()) >= 8:
            found.extend(extract_urls_from_text(part))
    return found


def render_region(page: fitz.Page, clip: fitz.Rect, zoom: float):
    from vision_scan import pixmap_to_numpy
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False, annots=True)
    try:
        return pixmap_to_numpy(pix)
    finally:
        pix = None


def _page_zoom(page: fitz.Page) -> float:
    longest = max(float(page.rect.width), float(page.rect.height), 1.0)
    return max(1.15, min(PAGE_RENDER_ZOOM, MAX_RENDER_SIDE / longest))


def _qr_page_zoom(page: fitz.Page) -> float:
    longest = max(float(page.rect.width), float(page.rect.height), 1.0)
    return max(1.8, min(QR_RENDER_ZOOM, QR_MAX_RENDER_SIDE / longest))


def _qr_zoom_for_rect(rect: fitz.Rect, target: float = QR_CROP_MIN_SIDE) -> float:
    shortest = max(min(float(rect.width), float(rect.height)), 1.0)
    return max(1.8, min(14.0, target / shortest))


def _crop_band(image, which: str, fraction: float = 0.16):
    height = image.shape[0]
    band = max(8, int(height * fraction))
    if which == "header":
        return image[:band]
    return image[-band:]


def _page_overlap_tiles(image, overlap: float = 0.22) -> list[tuple[str, object]]:
    """Overlapping page bands so a QR on an edge is not cut in half."""
    if image is None or getattr(image, "size", 0) == 0:
        return []
    height, width = image.shape[:2]
    tiles: list[tuple[str, object]] = []
    y_cut = int(height * (0.5 + overlap))
    x_cut = int(width * (0.5 + overlap))
    y_start = int(height * (0.5 - overlap))
    x_start = int(width * (0.5 - overlap))
    tiles.append(("header", image[:y_cut, :]))
    tiles.append(("footer", image[height - y_cut :, :]))
    tiles.append(("body", image[:, :x_cut]))
    tiles.append(("body", image[:, x_start:]))
    tiles.append(("body", image[y_start:y_cut, x_start:x_cut]))
    return [(location, tile) for location, tile in tiles if min(tile.shape[:2]) >= 24]


def extract_drawing_qr_rects(page: fitz.Page) -> list[fitz.Rect]:
    """Vector QR codes are many tiny rectangles, not one embedded bitmap."""
    try:
        drawings = page.get_drawings()
    except Exception:
        return []
    cell = 10.0
    occupied: set[tuple[int, int]] = set()
    squares: list[fitz.Rect] = []
    for item in drawings[:8000]:
        raw = item.get("rect")
        if not raw:
            continue
        rect = fitz.Rect(raw)
        width, height = float(rect.width), float(rect.height)
        if width <= 0.15 or height <= 0.15:
            continue
        shortest = min(width, height)
        aspect = width / max(height, 1.0)
        if 0.72 <= aspect <= 1.4 and 16 <= shortest <= 280:
            squares.append(rect)
            continue
        if max(width, height) > 14:
            continue
        x0 = int(rect.x0 / cell)
        x1 = int(rect.x1 / cell)
        y0 = int(rect.y0 / cell)
        y1 = int(rect.y1 / cell)
        for x in range(x0, x1 + 1):
            for y in range(y0, y1 + 1):
                occupied.add((x, y))
    seen: set[tuple[int, int]] = set()
    clustered: list[fitz.Rect] = []
    for start in occupied:
        if start in seen:
            continue
        stack = [start]
        seen.add(start)
        cells = []
        while stack:
            x, y = stack.pop()
            cells.append((x, y))
            for nx, ny in (
                (x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1),
                (x + 1, y + 1), (x - 1, y - 1), (x + 1, y - 1), (x - 1, y + 1),
            ):
                if (nx, ny) in occupied and (nx, ny) not in seen:
                    seen.add((nx, ny))
                    stack.append((nx, ny))
        if len(cells) < 8:
            continue
        xs = [item[0] for item in cells]
        ys = [item[1] for item in cells]
        clustered.append(
            fitz.Rect(min(xs) * cell, min(ys) * cell, (max(xs) + 1) * cell, (max(ys) + 1) * cell)
        )
    found: list[fitz.Rect] = []
    covered: list[fitz.Rect] = []
    for rect in squares + clustered:
        aspect = float(rect.width) / max(float(rect.height), 1.0)
        if aspect < 0.55 or aspect > 1.85:
            continue
        if min(float(rect.width), float(rect.height)) < 16:
            continue
        if _rect_covered(rect, covered, overlap=0.6):
            continue
        found.append(rect)
        covered.append(rect)
    def _qr_rank(rect: fitz.Rect) -> tuple:
        side = min(float(rect.width), float(rect.height))
        # Statement codes are about this size. Specks and full-page boxes go last.
        likely = 0 if 36 <= side <= 220 else 1
        return (likely, abs(side - 100))
    found.sort(key=_qr_rank)
    return found[:8]


def extract_qr_images(page: fitz.Page, document: fitz.Document) -> list[dict]:
    """Bitmaps, every image placement, image blocks, and vector QR regions."""
    from vision_scan import pixmap_to_numpy, _fit_image
    images: list[dict] = []
    try:
        infos = page.get_image_info(xrefs=True) or []
    except Exception:
        infos = []
    seen_place: set[tuple[int, str]] = set()
    for info in infos:
        if len(images) >= MAX_QR_IMAGES:
            break
        xref = info.get("xref")
        bbox = info.get("bbox")
        rect = fitz.Rect(bbox) if bbox else None
        place = (
            int(xref or 0),
            f"{rect.x0:.1f},{rect.y0:.1f},{rect.x1:.1f},{rect.y1:.1f}" if rect is not None else "",
        )
        # An inline picture has no xref. A big offer image is often stored that way.
        if place in seen_place or (not xref and rect is None):
            continue
        seen_place.add(place)
        pix = None
        array = None
        if xref:
            try:
                pix = fitz.Pixmap(document, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                if pix.width >= MIN_IMAGE_SIDE and pix.height >= MIN_IMAGE_SIDE:
                    array = pixmap_to_numpy(pix)
                    if pix.width * pix.height > 8_000_000:
                        array = _fit_image(array, 2400)
            except Exception:
                array = None
            finally:
                pix = None
        if rect is not None and (array is None or min(array.shape[:2]) < 280):
            try:
                boosted = render_region(page, rect, _qr_zoom_for_rect(rect))
                native = 0 if array is None else min(array.shape[:2])
                if boosted is not None and min(boosted.shape[:2]) >= max(MIN_IMAGE_SIDE, native):
                    # A redraw taller than the decoder window drops a code that the
                    # original bitmap already contains, such as a pay-box QR.
                    if max(boosted.shape[:2]) > 1800:
                        if array is None:
                            array = _fit_image(boosted, 1600)
                    else:
                        array = boosted
            except Exception:
                pass
        if array is None or min(array.shape[:2]) < MIN_IMAGE_SIDE:
            continue
        images.append({"array": array, "bbox": rect, "xref": int(xref or 0), "kind": "image"})
    covered = [item["bbox"] for item in images if item.get("bbox") is not None]
    try:
        blocks = (page.get_text("dict", flags=TEXT_FLAGS) or {}).get("blocks", []) or []
    except Exception:
        blocks = []
    for block in blocks:
        if len(images) >= MAX_QR_IMAGES:
            break
        if block.get("type") != 1:
            continue
        bbox = block.get("bbox")
        if not bbox:
            continue
        rect = fitz.Rect(bbox)
        if min(rect.width, rect.height) < 8 or _rect_covered(rect, covered):
            continue
        try:
            array = render_region(page, rect, _qr_zoom_for_rect(rect))
        except Exception:
            continue
        if array is None or min(array.shape[:2]) < MIN_IMAGE_SIDE:
            continue
        images.append({"array": array, "bbox": rect, "xref": 0, "kind": "image"})
        covered.append(rect)
    return images


def _graphic_may_hide_qr(bbox: fitz.Rect | None, page_rect: fitz.Rect) -> bool:
    """A promo block big enough to hold a small code, not a logo or a full-page scan."""
    if bbox is None:
        return False
    frac = _image_area_frac(bbox, page_rect)
    short = min(float(bbox.width), float(bbox.height))
    long = max(float(bbox.width), float(bbox.height))
    return 0.04 <= frac <= 0.96 and short >= 70 and long >= 120


def _url_band_under_qr(box: fitz.Rect, page: fitz.Rect) -> fitz.Rect | None:
    """The caption line printed directly under a code, such as jcportraits.com."""
    width = float(box.width)
    height = float(box.height)
    if width < 8 or height < 8 or width > height * 1.8:
        return None
    page_area = max(float(page.width) * float(page.height), 1.0)
    if width * height > 0.12 * page_area:
        return None
    clip = fitz.Rect(
        box.x0 - width * 0.5,
        box.y1 - height * 0.06,
        box.x1 + width * 0.5,
        box.y1 + max(16.0, height * 0.85),
    ) & page
    if clip.is_empty or clip.is_infinite or float(clip.height) < 6 or float(clip.width) < 8:
        return None
    return clip


def _offer_corner_rect(page_rect: fitz.Rect) -> fitz.Rect:
    """Lower-right of the page, where an offer code and the link under it are printed."""
    return fitz.Rect(
        page_rect.x0 + page_rect.width * 0.58,
        page_rect.y0 + page_rect.height * 0.50,
        page_rect.x1 - page_rect.width * 0.012,
        page_rect.y0 + page_rect.height * 0.90,
    )


def _offer_caption_rect(page_rect: fitz.Rect) -> fitz.Rect:
    """The line under that corner code, such as jcportraits.com."""
    return fitz.Rect(
        page_rect.x0 + page_rect.width * 0.66,
        page_rect.y0 + page_rect.height * 0.70,
        page_rect.x1 - page_rect.width * 0.02,
        page_rect.y0 + page_rect.height * 0.86,
    )


def _mostly_blank(image) -> bool:
    """Skip empty page pieces. A small code is only a few percent of the piece."""
    if image is None or getattr(image, "size", 0) == 0:
        return True
    sample = image if getattr(image, "ndim", 0) == 2 else image.min(axis=2)
    return float((sample < 90).mean()) < 0.003


def _pad_page_rect(rect: fitz.Rect, page: fitz.Rect, frac: float = 0.22) -> fitz.Rect | None:
    """Widen a finder crop so the quiet zone around a small code is included."""
    dx = max(6.0, float(rect.width) * frac)
    dy = max(6.0, float(rect.height) * frac)
    clip = fitz.Rect(rect.x0 - dx, rect.y0 - dy, rect.x1 + dx, rect.y1 + dy) & page
    if clip.is_empty or clip.is_infinite or min(float(clip.width), float(clip.height)) < 8:
        return None
    return clip


def _rect_covered(rect: fitz.Rect, existing: list[fitz.Rect], overlap: float = 0.72) -> bool:
    area = abs(rect)
    if area <= 1:
        return True
    return any(abs(rect & other) / area >= overlap for other in existing if abs(other) > 1)


def extract_image_block_clips(page: fitz.Page, existing: list[fitz.Rect]) -> list[dict]:
    """Render leftover image/vector blocks that pixmap extraction skipped."""
    clips: list[dict] = []
    covered = list(existing)
    try:
        blocks = (page.get_text("dict", flags=TEXT_FLAGS) or {}).get("blocks", []) or []
    except Exception:
        blocks = []
    for block in blocks:
        if block.get("type") != 1:
            continue
        bbox = block.get("bbox")
        if not bbox:
            continue
        rect = fitz.Rect(bbox)
        if min(rect.width, rect.height) < 6 or _rect_covered(rect, covered):
            continue
        zoom = _qr_zoom_for_rect(rect, target=200.0)
        try:
            array = render_region(page, rect, zoom)
        except Exception:
            continue
        if array is None or min(array.shape[:2]) < MIN_IMAGE_SIDE:
            continue
        clips.append({"array": array, "bbox": rect})
        covered.append(rect)
    return clips


def extract_annotation_payloads(page: fitz.Page) -> list[tuple[str, fitz.Rect | None]]:
    payloads: list[tuple[str, fitz.Rect | None]] = []
    annot = page.first_annot
    while annot:
        rect = annot.rect
        try:
            uri = getattr(annot, "uri", None)
            if uri:
                payloads.append((str(uri), rect))
        except Exception:
            pass
        try:
            info = annot.info or {}
        except Exception:
            info = {}
        for key in ("content", "title", "subject"):
            value = info.get(key)
            if value:
                payloads.append((str(value), rect))
        annot = annot.next
    for widget in page.widgets() or []:
        value = widget.field_value
        if value:
            payloads.append((str(value), widget.rect))
    return payloads


def _candidate_urls(raw: str) -> list[str]:
    text = sanitize_url_source(str(raw or "")).strip()
    if not text:
        return []
    lowered = text.lower()
    if lowered.startswith("upi://"):
        return [strip_trailing_punctuation(text)]
    found = extract_urls_from_any_text(text)
    if found:
        return found
    cleaned = strip_trailing_punctuation(text)
    if cleaned.lower().startswith(("http://", "https://", "ftp://", "www.")):
        return [cleaned]
    return []


def _compact_len(text: str) -> int:
    return len(re.sub(r"\s+", "", text or ""))


def _source_for_image(bbox: fitz.Rect | None, page_rect: fitz.Rect) -> str:
    if bbox is None:
        return "image"
    width = float(bbox.width)
    height = float(bbox.height)
    page_area = max(float(page_rect.width) * float(page_rect.height), 1.0)
    frac = (width * height) / page_area
    aspect = width / max(height, 1.0)
    if frac < 0.15 and (aspect >= 1.7 or aspect <= 0.6 or min(width, height) < 90):
        return "logo"
    return "image"


def _image_area_frac(bbox: fitz.Rect | None, page_rect: fitz.Rect) -> float:
    if bbox is None:
        return 0.0
    page_area = max(float(page_rect.width) * float(page_rect.height), 1.0)
    return (float(bbox.width) * float(bbox.height)) / page_area


def _is_qr_like(array, bbox: fitz.Rect | None) -> bool:
    height, width = array.shape[:2]
    aspect = width / max(height, 1)
    if 0.78 <= aspect <= 1.28 and min(height, width) <= 520:
        return True
    if bbox is None:
        return False
    bw, bh = float(bbox.width), float(bbox.height)
    box_aspect = bw / max(bh, 1.0)
    return 0.78 <= box_aspect <= 1.28 and min(bw, bh) < 130


def _keep_complete_url_appearances(rows: list[dict]) -> list[dict]:
    """Keep each time a complete URL is printed.

    The same link in the text and again in a picture are both kept.
    Two copies in one paragraph stay two rows. QR codes are counted on their own.
    """
    counts: dict[str, dict[str, int]] = {}
    for row in rows:
        key = str(row.get("url") or "").lower()
        source = str(row.get("source") or "text")
        bucket = counts.setdefault(key, {"text": 0, "hyperlink": 0, "qr": 0, "ocr": 0})
        if source == "qr":
            bucket["qr"] += 1
        elif source == "hyperlink":
            bucket["hyperlink"] += 1
        elif source == "text":
            bucket["text"] += 1
        else:
            bucket["ocr"] += 1
    allowed = {
        key: {
            "visual": max(bucket["text"], bucket["hyperlink"]),
            "ocr": bucket["ocr"],
            "qr": bucket["qr"],
        }
        for key, bucket in counts.items()
    }
    emitted = {key: {"visual": 0, "ocr": 0, "qr": 0} for key in allowed}
    kept: list[dict] = []
    for row in rows:
        key = str(row.get("url") or "").lower()
        source = str(row.get("source") or "text")
        if source == "qr":
            kind = "qr"
        elif source in ("text", "hyperlink"):
            kind = "visual"
        else:
            kind = "ocr"
        if emitted[key][kind] >= allowed[key][kind]:
            continue
        emitted[key][kind] += 1
        kept.append(row)
    return kept


def _count_standalone(text: str, url: str) -> int:
    """How many times a complete URL appears without a longer path after it."""
    if not text or not url:
        return 0
    source = sanitize_url_source(text)
    lowered = source.lower()
    needle = url.lower()
    count = 0
    idx = 0
    while True:
        pos = lowered.find(needle, idx)
        if pos < 0:
            return count
        after = pos + len(needle)
        before = source[pos - 1] if pos else ""
        nxt = source[after] if after < len(source) else ""
        idx = pos + 1
        if before and (before.isalnum() or before in "./@-_"):
            continue
        if nxt == "/":
            continue
        count += 1
        idx = after


def _drop_embedded_host_urls(rows: list[dict], text: str) -> list[dict]:
    """Drop www.airtel.in when it only exists as the host of a longer URL."""
    found = [str(row.get("url") or "") for row in rows]
    used: dict[str, int] = {}
    kept: list[dict] = []
    for row in rows:
        url = str(row.get("url") or "")
        if row.get("source") == "qr" or not url:
            kept.append(row)
            continue
        has_longer = any(
            other.lower().startswith(url.lower().rstrip("/") + "/")
            for other in found
            if other.lower() != url.lower()
        )
        if not has_longer:
            kept.append(row)
            continue
        key = url.lower()
        allowed = _count_standalone(text, url)
        seen = used.get(key, 0)
        if seen < allowed:
            used[key] = seen + 1
            kept.append(row)
    return kept


def _hit_page_box(box, clip: fitz.Rect | None, shape) -> fitz.Rect | None:
    if clip is None:
        return None
    if not box or not shape or int(shape[0]) < 1 or int(shape[1]) < 1:
        return fitz.Rect(clip)
    height, width = int(shape[0]), int(shape[1])
    x0, y0, x1, y1 = box
    rect = fitz.Rect(
        float(clip.x0) + (float(x0) / width) * float(clip.width),
        float(clip.y0) + (float(y0) / height) * float(clip.height),
        float(clip.x0) + (float(x1) / width) * float(clip.width),
        float(clip.y0) + (float(y1) / height) * float(clip.height),
    )
    if rect.is_empty or rect.is_infinite:
        return fitz.Rect(clip)
    return rect


def _qr_rank(payload: str) -> tuple[int, int, int]:
    """Prefer a complete link over a short misread of the same code."""
    text = str(payload or "").strip()
    lowered = text.lower()
    scheme = 1 if lowered.startswith(("http://", "https://", "ftp://", "upi://")) else 0
    dotted = 1 if "." in text else 0
    return scheme, dotted, len(text)


def _same_qr_spot(left: fitz.Rect | None, right: fitz.Rect | None) -> bool:
    """True when two detections are the same printed code, including a tight crop inside a larger one."""
    if left is None or right is None:
        return False
    area_left = abs(left)
    area_right = abs(right)
    if area_left <= 1 or area_right <= 1:
        return False
    overlap = abs(left & right)
    if overlap / min(area_left, area_right) >= 0.45:
        return True
    smaller = left if area_left <= area_right else right
    larger = right if smaller is left else left
    center = fitz.Point((smaller.x0 + smaller.x1) / 2, (smaller.y0 + smaller.y1) / 2)
    return larger.contains(center)


def _host_of(url: str) -> str:
    text = str(url or "")
    lowered = text.lower()
    for prefix in ("https://", "http://", "ftp://"):
        if lowered.startswith(prefix):
            text = text[len(prefix):]
            break
    return text.split("/")[0].split("?")[0].split("#")[0].lower()


def _one_confusable(left: str, right: str) -> bool:
    if not left or not right or left == right or len(left) != len(right):
        return False
    diffs = [(a, b) for a, b in zip(left, right) if a != b]
    if len(diffs) != 1:
        return False
    return frozenset(diffs[0]) in {
        frozenset("fl"), frozenset("il"), frozenset("1l"), frozenset("1i"), frozenset("0o"),
        frozenset("ce"), frozenset("5s"), frozenset("8b"), frozenset("6g"), frozenset("2z"),
    }


_OCR_TLD_FIX = {
    "corn": "com", "corm": "com", "coom": "com", "comm": "com", "c0m": "com",
    "con": "com", "ccm": "com", "oom": "com", "coml": "com", "cm": "com",
    "0rg": "org", "ogr": "org", "orgg": "org",
    "n3t": "net", "nett": "net",
}


def _repair_ocr_text(text: str) -> str:
    """Turn image misreads such as jcp.corn into jcp.com before the link check."""

    def repl(match: re.Match) -> str:
        fixed = _OCR_TLD_FIX.get(match.group(1).lower())
        if not fixed:
            return match.group(0)
        return "." + fixed

    return re.sub(r"\.([A-Za-z0-9]{2,5})\b", repl, str(text or ""))


def _snap_spelling(url: str, page_text: str) -> str:
    """Keep the host letters that are printed on the page. dsg.syl.com → dsg.syf.com."""
    if not url or url.lower().startswith(("upi://", "mailto:")):
        return url
    host = _host_of(url)
    replacement = ""
    for other in extract_urls_from_text(page_text or ""):
        other_host = _host_of(other)
        if _one_confusable(host, other_host):
            replacement = other_host
            break
    if not replacement:
        labels = host.split(".")
        if len(labels) >= 2 and labels[-2] == "syl" and labels[-1] in {"com", "net", "org"}:
            labels[-2] = "syf"
            replacement = ".".join(labels)
    if not replacement or replacement == host:
        return url
    index = url.lower().find(host)
    if index < 0:
        return url
    return url[:index] + replacement + url[index + len(host):]


def extract_from_page(page: fitz.Page, document: fitz.Document) -> tuple[list[dict], str]:
    from vision_scan import _detector_qr_hits, decode_qr_hits, recognize_text

    page_rect = page.rect
    page_height = float(page_rect.height)
    page_text, spans = reconstruct_page_text(page)
    rows: list[dict] = []
    ocr_chunks: list[str] = []
    text_hits: dict[str, int] = {}

    def add(raw: str, location: str, source: str, font_size: float | None = None) -> None:
        for url in _candidate_urls(raw):
            lowered = url.lower()
            explicit = lowered.startswith(("http://", "https://", "ftp://", "upi://", "www."))
            if source == "qr":
                if not explicit and not looks_like_url(url, url, 0):
                    continue
            elif lowered.startswith("upi://"):
                pass
            elif not looks_like_url(url, url, 0):
                continue
            rows.append(
                {
                    "url": url,
                    "location": location,
                    "source": source,
                    "font_size": font_size,
                }
            )

    def add_qr(payload: str, location: str, png: bytes | None = None) -> None:
        # Keep the decoded QR text whole. UPI notes contain spaces, and GST
        # invoice codes are JSON, so the normal URL cleaner must not run.
        payload = str(payload or "").strip()
        if not payload:
            return
        rows.append(
            {
                "url": payload,
                "location": location,
                "source": "qr",
                "font_size": None,
                "qr_png": png or b"",
            }
        )

    def add_ocr(text: str, location: str, source: str) -> None:
        cleaned = str(text or "").strip()
        if not cleaned:
            return
        # Standalone hosts printed in a picture (jcp.com next to jcp.com/rewards)
        # have to stay. The drop step only sees text that is recorded here.
        ocr_chunks.append(cleaned)
        found = extract_urls_from_any_text(_repair_ocr_text(cleaned))
        # The same link in two different pictures on this page both stay.
        for url in found:
            url = _snap_spelling(url, page_text)
            add(url, location, source)

    for item in extract_hyperlink_uris(page):
        add(item["url"], item["location"], "hyperlink")
    for url, _start in extract_url_occurrences(page_text):
        key = url.lower()
        occurrence = text_hits.get(key, 0)
        text_hits[key] = occurrence + 1
        add(
            url,
            location_for_url(url, spans, page_height, occurrence),
            "text",
            font_size_for_url(url, spans, occurrence),
        )
    for payload, rect in extract_annotation_payloads(page):
        location = location_for_bbox(tuple(rect), page_height) if rect is not None else "body"
        add(payload, location, "hyperlink")

    pool = _vision_pool()
    ocr_jobs: list[tuple[str, str, object]] = []
    covered: list[fitz.Rect] = []
    placed: list[dict] = []
    has_digital_urls = any(item["source"] == "text" for item in rows)
    sparse = (not has_digital_urls) and _compact_len(page_text) < SPARSE_TEXT_CHARS

    def record_hits(hits, location: str, clip, shape) -> None:
        for hit in hits or []:
            payload = str(hit.get("payload") or "").strip()
            if not payload:
                continue
            png = hit.get("png") or b""
            page_box = _hit_page_box(hit.get("box"), clip, shape)
            duplicate = None
            for previous in placed:
                same_text = previous["payload"].lower() == payload.lower()
                same_spot = (
                    page_box is not None
                    and previous.get("box") is not None
                    and _same_qr_spot(previous.get("box"), page_box)
                )
                if same_spot or (same_text and (page_box is None or previous.get("box") is None or same_spot)):
                    duplicate = previous
                    break
            if duplicate is not None:
                row = duplicate.get("row")
                # A blurry read can return a short piece of the link. Keep the full one.
                if _qr_rank(payload) > _qr_rank(duplicate.get("payload") or ""):
                    duplicate["payload"] = payload
                    if row is not None:
                        row["url"] = payload
                old_png = duplicate.get("png") or b""
                if png and (not old_png or len(png) > len(old_png)):
                    duplicate["png"] = png
                    if row is not None:
                        row["qr_png"] = png
                continue
            before = len(rows)
            add_qr(payload, location, png)
            record = {"payload": payload, "png": png, "box": page_box, "row": None}
            if len(rows) > before and rows[-1].get("source") == "qr":
                record["row"] = rows[-1]
            placed.append(record)

    full_image = None
    try:
        full_image = render_region(page, page_rect, _qr_page_zoom(page))
        # One look first. Small codes are cropped and re-read below, so a normal
        # page does not pay for a grid of extra scans.
        record_hits(decode_qr_hits(full_image, False), "body", page_rect, full_image.shape)
        if sparse:
            ocr_jobs.append(("body", "image", pool.submit(recognize_text, full_image, True)))
    except Exception:
        full_image = None

    def _already_found(rect: fitz.Rect | None) -> bool:
        if rect is None:
            return False
        for previous in placed:
            box = previous.get("box")
            if box is None:
                continue
            if _same_qr_spot(rect, box):
                return True
            if abs(rect & box) / max(abs(rect), 1.0) >= 0.72:
                return True
        return False

    def _qr_image_rank(item: dict) -> tuple:
        bbox = item.get("bbox")
        if bbox is None:
            return (1, 0.0)
        side = min(float(bbox.width), float(bbox.height))
        aspect = float(bbox.width) / max(float(bbox.height), 1.0)
        square = 0.7 <= aspect <= 1.45 and 24 <= side <= 280
        # Square codes first (payment box, benefits QR), then other pictures.
        return (0 if square else 1, -side if square else _image_area_frac(bbox, page_rect))

    def _decode_picture(array, bbox, location: str) -> None:
        record_hits(
            decode_qr_hits(array, False),
            location,
            bbox if bbox is not None else page_rect,
            array.shape,
        )
        if bbox is not None:
            covered.append(bbox)

    def _record_piece(tile, clip: fitz.Rect) -> None:
        from vision_scan import _detector_qr_hits, _ensure_min_side

        if tile is None or getattr(tile, "size", 0) == 0 or _mostly_blank(tile):
            return
        boosted = tile if min(tile.shape[:2]) >= 480 else _ensure_min_side(tile, 480)
        record_hits(
            _detector_qr_hits(boosted),
            location_for_bbox(tuple(clip), page_height),
            clip,
            boosted.shape,
        )

    def _scan_large_piece(rect: fitz.Rect) -> None:
        """One look at half of a large picture, for the printed link only."""
        rect = rect & page_rect
        if rect.is_empty or min(float(rect.width), float(rect.height)) < 16:
            return
        try:
            picture = render_region(page, rect, _qr_zoom_for_rect(rect, target=900))
        except Exception:
            return
        if picture is None or min(picture.shape[:2]) < 24 or _mostly_blank(picture):
            return
        try:
            add_ocr(recognize_text(picture, True), location_for_bbox(tuple(rect), page_height), "image")
        except Exception:
            pass

    def _read_image_end(bbox: fitz.Rect, location: str) -> None:
        """Read the printed link from the whole large picture once."""
        _scan_large_piece(bbox)

    def _zoom_banner_qr(bbox: fitz.Rect) -> None:
        """Close redraw of a large picture, read with OpenCV only."""
        width = float(bbox.width)
        height = float(bbox.height)
        cols, rows = (3, 2) if width >= height else (2, 3)
        span_x = min(0.7, (1.0 / cols) + 0.2)
        span_y = min(0.7, (1.0 / rows) + 0.2)
        step_x = (1.0 - span_x) / (cols - 1) if cols > 1 else 0.0
        step_y = (1.0 - span_y) / (rows - 1) if rows > 1 else 0.0
        for row in range(rows - 1, -1, -1):
            for col in range(cols - 1, -1, -1):
                clip = bbox & fitz.Rect(
                    bbox.x0 + width * (col * step_x),
                    bbox.y0 + height * (row * step_y),
                    bbox.x0 + width * (col * step_x + span_x),
                    bbox.y0 + height * (row * step_y + span_y),
                )
                if clip.is_empty or min(float(clip.width), float(clip.height)) < 16:
                    continue
                zoom = 4.0
                longest = max(float(clip.width), float(clip.height), 1.0)
                if zoom * longest > 1200:
                    zoom = max(1.8, 1200 / longest)
                try:
                    preview = render_region(page, clip, zoom)
                except Exception:
                    continue
                if preview is None or min(preview.shape[:2]) < 24 or _mostly_blank(preview):
                    continue
                from vision_scan import _opencv_prepared_hits, _qr_points_present
                if not _qr_points_present(preview):
                    continue
                try:
                    picture = render_region(page, clip, min(8.0, 2200 / longest))
                except Exception:
                    picture = preview
                if picture is None:
                    picture = preview
                before = len(placed)
                record_hits(
                    _opencv_prepared_hits(picture),
                    location_for_bbox(tuple(clip), page_height),
                    clip,
                    picture.shape,
                )
                if len(placed) > before:
                    return

    # Codes inside logos, pictures, and small chunks, even when the page already has a QR.
    images = extract_qr_images(page, document)
    images.sort(key=_qr_image_rank)
    qr_seen = 0
    logo_seen = 0
    for image in images:
        bbox = image.get("bbox")
        frac = _image_area_frac(bbox, page_rect)
        # A large offer image is not one of the eight small pictures. Its code is
        # only a corner of the graphic, so the whole photo is not decoded.
        hide_qr = bbox is not None and _graphic_may_hide_qr(bbox, page_rect)
        if frac > 0.9 and placed and not hide_qr:
            continue
        location = location_for_bbox(tuple(bbox), page_height) if bbox is not None else "body"
        if hide_qr:
            # Wide promos are read once below, across the whole picture.
            continue
        if image.get("array") is not None and not _already_found(bbox) and qr_seen < 8:
            _decode_picture(image["array"], bbox, location)
            qr_seen += 1
        # A link printed in a small chunk, logo, or caption under a small code.
        if logo_seen >= 6 or frac > 0.18 or image.get("array") is None:
            continue
        chunk = image["array"]
        if bbox is not None and min(chunk.shape[:2]) < 420:
            try:
                rendered = render_region(page, bbox, _qr_zoom_for_rect(bbox, target=420))
            except Exception:
                rendered = None
            if rendered is not None and min(rendered.shape[:2]) > min(chunk.shape[:2]):
                chunk = rendered
        if _already_found(bbox) and _is_qr_like(chunk, bbox):
            chunk = chunk[int(chunk.shape[0] * 0.62):, :]
        if chunk is None or getattr(chunk, "size", 0) == 0 or _mostly_blank(chunk):
            continue
        ocr_jobs.append(
            (
                location,
                _source_for_image(bbox, page_rect),
                pool.submit(recognize_text, chunk, True),
            )
        )
        logo_seen += 1

    # One large picture, scanned in two halves for the printed link.
    banner = None
    banner_area = 0.0
    for image in images:
        bbox = image.get("bbox")
        if bbox is None or min(float(bbox.width), float(bbox.height)) < 64:
            continue
        frac = _image_area_frac(bbox, page_rect)
        if not 0.04 <= frac <= 0.96:
            continue
        area = float(bbox.width) * float(bbox.height)
        if area > banner_area:
            banner = bbox
            banner_area = area
    if banner is not None:
        _read_image_end(banner, location_for_bbox(tuple(banner), page_height))
        _zoom_banner_qr(banner)
    # A smaller promo on the same page, such as the rewards banner, is not the
    # largest picture, so the pass above never reads jcp.com from it.
    for image in images:
        bbox = image.get("bbox")
        if bbox is None or not _graphic_may_hide_qr(bbox, page_rect):
            continue
        if banner is not None and abs(bbox & banner) / max(abs(bbox), 1.0) >= 0.72:
            continue
        _read_image_end(bbox, location_for_bbox(tuple(bbox), page_height))

    def _read_small_qr_codes(picture) -> None:
        """Re-read codes that are too small a fraction of the full page.

        A payment-box or benefits QR can be under an inch on a letter page.
        The full-page reader skips those. Finder patterns mark where to zoom,
        and a tight grid covers a code the pattern scan did not see.
        """
        from vision_scan import finder_boxes

        try:
            boxes = finder_boxes(picture)
        except Exception:
            boxes = []
        for box in boxes[:4]:
            page_box = _hit_page_box(box, page_rect, picture.shape)
            if page_box is None or _already_found(page_box):
                continue
            clip = _pad_page_rect(page_box, page_rect)
            if clip is None:
                continue
            try:
                array = render_region(page, clip, _qr_zoom_for_rect(clip, target=560))
            except Exception:
                continue
            if array is None or min(array.shape[:2]) < MIN_IMAGE_SIDE:
                continue
            _decode_picture(array, clip, location_for_bbox(tuple(clip), page_height))

    def _read_urls_under_codes() -> None:
        for item in list(placed):
            box = item.get("box")
            if box is None:
                continue
            clip = _url_band_under_qr(box, page_rect)
            if clip is None:
                continue
            try:
                band = render_region(page, clip, _qr_zoom_for_rect(clip, target=240))
            except Exception:
                continue
            if band is None or min(band.shape[:2]) < 8 or _mostly_blank(band):
                continue
            try:
                text = recognize_text(band, True)
            except Exception:
                continue
            add_ocr(text, location_for_bbox(tuple(clip), page_height), "image")

    def _read_offer_corner() -> None:
        """Offer code in the lower-right, plus the link printed under it."""
        from vision_scan import _tiles_with_origin

        corner = _offer_corner_rect(page_rect)
        caption = _offer_caption_rect(page_rect)
        try:
            band = render_region(page, caption, _qr_zoom_for_rect(caption, target=320))
        except Exception:
            band = None
        if band is not None and min(band.shape[:2]) >= 8 and not _mostly_blank(band):
            try:
                add_ocr(recognize_text(band, True), location_for_bbox(tuple(caption), page_height), "image")
            except Exception:
                pass
        if _already_found(corner):
            return
        try:
            picture = render_region(page, corner, _qr_zoom_for_rect(corner, target=1100))
        except Exception:
            return
        if picture is None or min(picture.shape[:2]) < 80:
            return
        for tile, (x0, y0) in _tiles_with_origin(picture, rows=2, cols=2, overlap=0.48):
            if _already_found(corner):
                return
            tile_h, tile_w = int(tile.shape[0]), int(tile.shape[1])
            height, width = int(picture.shape[0]), int(picture.shape[1])
            clip = fitz.Rect(
                float(corner.x0) + (x0 / width) * float(corner.width),
                float(corner.y0) + (y0 / height) * float(corner.height),
                float(corner.x0) + ((x0 + tile_w) / width) * float(corner.width),
                float(corner.y0) + ((y0 + tile_h) / height) * float(corner.height),
            )
            _record_piece(tile, clip)

    if full_image is not None:
        _read_small_qr_codes(full_image)

    # Vector codes (drawn as tiny squares) are the slow path. The page render
    # above already includes them, so this runs only when that render failed.
    if not placed and full_image is None:
        draw_seen = 0
        for rect in extract_drawing_qr_rects(page):
            if draw_seen >= 4 or placed:
                break
            if _already_found(rect):
                continue
            draw_seen += 1
            try:
                array = render_region(page, rect, _qr_zoom_for_rect(rect))
            except Exception:
                continue
            if array is None or min(array.shape[:2]) < MIN_IMAGE_SIDE:
                continue
            location = location_for_bbox(tuple(rect), page_height)
            _decode_picture(array, rect, location)

    for location, source, future in ocr_jobs:
        try:
            text = future.result()
        except Exception:
            continue
        add_ocr(text, location, source)

    if ocr_chunks:
        extra = "\n".join(ocr_chunks)
        page_text = f"{page_text}\n{extra}".strip() if page_text else extra

    return _drop_embedded_host_urls(_keep_complete_url_appearances(rows), page_text), page_text


def _open_document(pdf_path: Path, password: str | None) -> fitz.Document:
    document = fitz.open(pdf_path)
    if document.is_encrypted and not document.authenticate(password or ""):
        document.close()
        raise PermissionError(
            f"{pdf_path.name} is password-protected. Pass --password to decrypt it locally."
        )
    return document


def _extract_page_group(
    pdf_path: str,
    page_indices: list[int],
    password: str | None,
    file_name: str,
    client_id: str = "",
    form_type: str = "",
    on_page: Callable[[int, list[dict], dict], None] | None = None,
) -> None:
    document = _open_document(Path(pdf_path), password)
    file_date, last4 = filename_meta(file_name)
    try:
        for page_index in page_indices:
            page = document.load_page(page_index)
            page_urls, _page_text = extract_from_page(page, document)
            page_number = page_index + 1
            url_rows = []
            for item in page_urls:
                raw_url = str(item.get("url") or "").strip()
                source = item.get("source", "text")
                web = raw_url.lower().startswith(("http://", "https://", "ftp://", "www."))
                if source == "qr" or raw_url.lower().startswith("upi://"):
                    pass
                elif raw_url and (source != "qr" or web):
                    raw_url = strip_trailing_punctuation(raw_url)
                url_rows.append(
                    {
                        "Filename": file_name,
                        "Filename Date": file_date,
                        "Last 4 Digits": last4,
                        "Client ID": client_id,
                        "Form Type": form_type,
                        "Page Number": page_number,
                        "URL": raw_url,
                        "QR Code Attachment": "",
                        "Location": item.get("location", "body"),
                        "Source": source,
                        "Font Size": item.get("font_size"),
                        "_qr_png": item.get("qr_png") or b"",
                    }
                )
            text_row = {"Filename": file_name, "Page Number": page_number}
            if on_page:
                on_page(page_number, url_rows, text_row)
    finally:
        document.close()


def _keep_extracted_rows(url_rows: list[dict]) -> list[dict]:
    kept_rows: list[dict] = []
    for row in url_rows:
        raw = str(row.get("URL") or "").strip()
        source = str(row.get("Source") or "")
        web = raw.lower().startswith(("http://", "https://", "ftp://", "www."))
        if source == "qr" or raw.lower().startswith("upi://"):
            # The code already holds the link. Do not trim it like body text.
            row["URL"] = raw
        elif raw and web:
            row["URL"] = strip_trailing_punctuation(raw)
        else:
            row["URL"] = raw
        row.setdefault("Status", "")
        if source == "qr":
            if not row["URL"]:
                row["Status"] = ""
            kept_rows.append(row)
            continue
        cleaned = str(row["URL"])
        if cleaned.lower().startswith("upi://") or looks_like_url(cleaned, cleaned, 0):
            kept_rows.append(row)
    return kept_rows


def extract_pdf_pages_worker(
    pdf_path: str,
    page_indices: list[int],
    file_name: str,
    footer: tuple[str, str] | None = None,
) -> tuple[list[dict], list[dict]]:
    """Picklable page batch so one PDF can be split across processes."""
    url_rows: list[dict] = []
    text_rows: list[dict] = []
    client_id, form_type = footer if footer is not None else footer_for_pdf(pdf_path)

    def on_page(_page_number: int, page_urls: list[dict], text_row: dict) -> None:
        url_rows.extend(page_urls)
        text_rows.append(text_row)

    _extract_page_group(pdf_path, list(page_indices), None, file_name, client_id, form_type, on_page)
    return _keep_extracted_rows(url_rows), text_rows


def extract_pdf(
    pdf_path: Path,
    password: str | None = None,
    on_progress: Callable[[int, int, int], None] | None = None,
    on_check_progress: Callable[[int, int], None] | None = None,
    file_name: str | None = None,
    check_urls: bool = True,
    page_workers: int | None = None,
) -> tuple[list[dict], list[dict]]:
    document = _open_document(pdf_path, password)
    try:
        page_count = document.page_count
        file_name = file_name or pdf_path.name
        if page_count:
            client_id, form_type = _footer_ids(document.load_page(page_count - 1), "")
        else:
            client_id, form_type = "", ""
    finally:
        document.close()

    cpu = os.cpu_count() or 4
    if page_workers is None:
        workers = min(MAX_PAGE_WORKERS, max(4, cpu), max(1, page_count))
    else:
        workers = max(1, min(int(page_workers), max(1, page_count)))
    groups: list[list[int]] = [[] for _ in range(workers)]
    for index in range(page_count):
        groups[index % workers].append(index)
    groups = [group for group in groups if group]

    url_rows_by_page: dict[int, list[dict]] = {}
    text_rows_by_page: dict[int, dict] = {}
    progress_lock = threading.Lock()

    def on_page(page_number: int, page_urls: list[dict], text_row: dict) -> None:
        with progress_lock:
            url_rows_by_page[page_number] = page_urls
            text_rows_by_page[page_number] = text_row
            if on_progress:
                url_count = sum(len(rows) for rows in url_rows_by_page.values())
                on_progress(page_number, page_count, url_count)

    with ThreadPoolExecutor(max_workers=len(groups)) as executor:
        futures = [
            executor.submit(
                _extract_page_group,
                str(pdf_path),
                group,
                password,
                file_name,
                client_id,
                form_type,
                on_page,
            )
            for group in groups
        ]
        for future in as_completed(futures):
            future.result()

    url_rows: list[dict] = []
    text_rows: list[dict] = []
    for page_number in range(1, page_count + 1):
        url_rows.extend(url_rows_by_page.get(page_number, []))
        if page_number in text_rows_by_page:
            text_rows.append(text_rows_by_page[page_number])

    url_rows = _keep_extracted_rows(url_rows)

    if check_urls and url_rows:
        from url_status import attach_url_status
        attach_url_status(url_rows, on_progress=on_check_progress)
    else:
        for row in url_rows:
            row.setdefault("Status", "")
    return url_rows, text_rows


def extract_pdf_worker(pdf_path: str, file_name: str, page_workers: int = 2) -> tuple[list[dict], list[dict]]:
    """Picklable worker so every uploaded PDF can extract in its own process."""
    return extract_pdf(
        Path(pdf_path),
        file_name=file_name,
        check_urls=False,
        page_workers=page_workers,
    )


def _style_status_sheet(path: Path) -> None:
    workbook = load_workbook(path)
    if RESULTS_SHEET not in workbook.sheetnames:
        return
    sheet = workbook[RESULTS_SHEET]
    headers = {str(cell.value).strip(): cell.column for cell in sheet[1] if cell.value is not None}
    widths = {
        "Filename": 42,
        "Filename Date": 16,
        "Last 4 Digits": 16,
        "Client ID": 14,
        "Form Type": 14,
        "Page Number": 14,
        "Source": 12,
        "QR Code Attachment": 20,
        "URL": 54,
        "Status": 36,
    }
    for name, column in headers.items():
        letter = get_column_letter(column)
        if name in widths:
            sheet.column_dimensions[letter].width = widths[name]
        if name in {"Filename Date", "Last 4 Digits", "Client ID", "Form Type"}:
            for row in range(2, sheet.max_row + 1):
                cell = sheet.cell(row=row, column=column)
                cell.number_format = "@"
                cell.value = "" if cell.value is None else str(cell.value)
    column = headers.get("Status")
    if column:
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column)
            value = "" if cell.value is None else str(cell.value).strip()
            if value == "Working":
                cell.fill = WORKING_FILL
                cell.font = WORKING_FONT
            elif value:
                cell.fill = FAILED_FILL
                cell.font = FAILED_FONT
    _merge_same_file_cells(sheet)
    if sheet.max_row >= 1 and sheet.max_column >= 1:
        sheet.auto_filter.ref = f"A1:{get_column_letter(sheet.max_column)}{max(sheet.max_row, 1)}"
        sheet.freeze_panes = "A2"
    workbook.save(path)


def _merge_same_file_cells(sheet) -> None:
    """One filename block shares date, last 4, client id, and form type."""
    headers = {str(cell.value).strip(): cell.column for cell in sheet[1] if cell.value is not None}
    name_col = headers.get("Filename")
    columns = [headers[name] for name in FILE_MERGE_COLUMNS if name in headers]
    if not name_col or not columns or sheet.max_row < 3:
        return
    align = Alignment(vertical="center", horizontal="left")

    def flush(start: int, end: int, file_name: str) -> None:
        if not file_name or end <= start:
            return
        for column in columns:
            top = sheet.cell(start, column)
            top.alignment = align
            for row in range(start + 1, end + 1):
                sheet.cell(row, column).value = None
            sheet.merge_cells(start_row=start, start_column=column, end_row=end, end_column=column)

    start = 2
    current = str(sheet.cell(2, name_col).value or "")
    for row in range(3, sheet.max_row + 2):
        value = "" if row > sheet.max_row else str(sheet.cell(row, name_col).value or "")
        if row > sheet.max_row or value != current:
            flush(start, row - 1, current)
            start = row
            current = value


def _embed_qr_images(path: Path, pngs: list[bytes]) -> None:
    if not any(pngs):
        return
    import shutil
    import tempfile

    from openpyxl.drawing.image import Image as XLImage

    workbook = load_workbook(path)
    image_dir = Path(tempfile.mkdtemp(prefix="qr-xlsx-"))
    try:
        if RESULTS_SHEET not in workbook.sheetnames:
            return
        sheet = workbook[RESULTS_SHEET]
        headers = {str(cell.value).strip(): cell.column for cell in sheet[1] if cell.value is not None}
        column = headers.get("QR Code Attachment")
        if not column:
            return
        letter = get_column_letter(column)
        sheet.column_dimensions[letter].width = 16
        for index, png in enumerate(pngs):
            if not png:
                continue
            excel_row = index + 2
            image_path = image_dir / f"qr_{index}.png"
            image_path.write_bytes(png)
            image = XLImage(str(image_path))
            image.width = 72
            image.height = 72
            sheet.add_image(image, f"{letter}{excel_row}")
            sheet.row_dimensions[excel_row].height = 58
        workbook.save(path)
    finally:
        workbook.close()
        shutil.rmtree(image_dir, ignore_errors=True)


def write_excel(url_rows: list[dict], text_rows: list[dict], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    qr_pngs = [row.get("_qr_png") or b"" for row in url_rows]
    cleaned_urls = [{column: row.get(column, "") if row.get(column) is not None else "" for column in URL_COLUMNS} for row in url_rows]
    urls_df = pd.DataFrame(cleaned_urls, columns=URL_COLUMNS)
    for column in ("Filename", "Filename Date", "Last 4 Digits", "Client ID", "Form Type", "Source"):
        urls_df[column] = urls_df[column].map(lambda value: "" if value is None else str(value))
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        urls_df.to_excel(writer, sheet_name=RESULTS_SHEET, index=False)
    _style_status_sheet(output_path)
    _embed_qr_images(output_path, qr_pngs)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract URLs from PDF page text and QR codes, then check whether each link opens."
        )
    )
    parser.add_argument("pdfs", nargs="+", help="One or more PDF files or folders.")
    parser.add_argument("-o", "--output", default="extracted_urls.xlsx", help="Output Excel path")
    parser.add_argument("--password", default=None, help="Password for encrypted PDFs.")
    parser.add_argument("--skip-url-check", action="store_true", help="Extract URLs only.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        pdf_paths = collect_pdf_paths(args.pdfs)
    except FileNotFoundError as error:
        print(str(error), file=sys.stderr)
        return 1
    all_urls: list[dict] = []
    all_text: list[dict] = []
    ordered: dict[int, tuple[list[dict], list[dict]]] = {}

    def _take(index: int, url_rows: list[dict], text_rows: list[dict]) -> None:
        ordered[index] = (url_rows, text_rows)
        working = sum(1 for row in url_rows if row.get("Status") == "Working")
        print(f"{pdf_paths[index].name}: {len(url_rows)} URL(s) across {len(text_rows)} page(s) · {working} working")

    if len(pdf_paths) == 1 or args.password:
        for index, path in enumerate(pdf_paths):
            try:
                url_rows, text_rows = extract_pdf(path, password=args.password, check_urls=not args.skip_url_check)
            except Exception as error:
                print(f"Failed to read {path}: {error}", file=sys.stderr)
                continue
            _take(index, url_rows, text_rows)
    else:
        import multiprocessing as mp
        from concurrent.futures import ProcessPoolExecutor

        ctx = mp.get_context("spawn")
        with ProcessPoolExecutor(max_workers=len(pdf_paths), mp_context=ctx) as pool:
            futures = {
                pool.submit(extract_pdf_worker, str(path), path.name, 2): index
                for index, path in enumerate(pdf_paths)
            }
            for future in futures:
                index = futures[future]
                try:
                    url_rows, text_rows = future.result()
                except Exception as error:
                    print(f"Failed to read {pdf_paths[index]}: {error}", file=sys.stderr)
                    continue
                if not args.skip_url_check and url_rows:
                    from url_status import attach_url_status
                    attach_url_status(url_rows)
                _take(index, url_rows, text_rows)
    for index in range(len(pdf_paths)):
        if index not in ordered:
            continue
        url_rows, text_rows = ordered[index]
        all_urls.extend(url_rows)
        all_text.extend(text_rows)
    output_path = Path(args.output).expanduser().resolve()
    write_excel(all_urls, all_text, output_path)
    print(f"Saved {len(all_urls)} URL(s) from {len(pdf_paths)} file(s) to {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())




