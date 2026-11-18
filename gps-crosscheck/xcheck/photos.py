"""Photo checks: download from Drive, read the printed stamp, and look for tampering.

Everything here is free and offline (RapidOCR + OpenCV). Each photo returns a PhotoResult
with what was read and a list of problems found.
"""
from __future__ import annotations

import datetime as dt
import io
import logging
import math
import os
import re
from dataclasses import dataclass, field

import cv2
import numpy as np
import requests
from PIL import Image

log = logging.getLogger(__name__)

DOWNLOAD = "https://drive.usercontent.google.com/download?id={id}&export=download"
WHATSAPP_NAME = re.compile(r"WhatsApp Image (\d{4})-(\d{2})-(\d{2}) at (\d{1,2})\.(\d{2})\.(\d{2})", re.I)
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


@dataclass
class PhotoResult:
    drive_id: str
    role: str                                  # start | equipment | end
    ok_download: bool = False
    file_name: str = ""
    whatsapp_time: dt.datetime | None = None   # from the file name, if present
    stamp_time: dt.datetime | None = None
    stamp_text: str = ""
    lat: float | None = None
    lon: float | None = None
    brightness: float | None = None            # 0..255 of the upper part
    sky_bright: bool = False
    ocr_conf: float = 0.0                      # how sure the OCR is about the stamp line
    dhash: str = ""
    problems: list[dict] = field(default_factory=list)

    def add(self, code: str, severity: str, message: str):
        self.problems.append({"code": code, "severity": severity, "message": message})


# ----------------------------------------------------------------------------- download

def download(drive_id: str, cache_dir: str) -> tuple[bytes | None, str]:
    os.makedirs(cache_dir, exist_ok=True)
    path = os.path.join(cache_dir, drive_id)
    if os.path.exists(path) and os.path.exists(path + ".name"):
        return open(path, "rb").read(), open(path + ".name").read()
    try:
        r = requests.get(DOWNLOAD.format(id=drive_id), timeout=60)
    except requests.RequestException:
        return None, ""
    if r.status_code != 200 or not r.headers.get("content-type", "").startswith("image"):
        return None, ""
    m = re.search(r'filename="?([^";]+)', r.headers.get("content-disposition", ""))
    name = m.group(1) if m else ""
    open(path, "wb").write(r.content)
    open(path + ".name", "w").write(name)
    return r.content, name


def whatsapp_time(name: str) -> dt.datetime | None:
    m = WHATSAPP_NAME.search(name or "")
    if not m:
        return None
    y, mo, d, h, mi, s = map(int, m.groups())
    try:
        return dt.datetime(y, mo, d, h, mi, s)
    except ValueError:
        return None


# ----------------------------------------------------------------------------- OCR
# RapidOCR (free, offline, built for text in photos). Loaded once per worker process.
_ENGINE = None


def _engine():
    global _ENGINE
    if _ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        _ENGINE = RapidOCR()
    return _ENGINE


def _ocr(bgr: np.ndarray) -> list[tuple[str, float]]:
    res, _ = _engine()(bgr)
    return [(r[1], float(r[2])) for r in (res or [])]


def _normalise(t: str) -> str:
    """RapidOCR drops spaces ('17Nov202611:14:05pm'): put them back where it matters."""
    t = re.sub(r"(20\d{2})[.,]?(\d{1,2}\s*[:.]\s*\d{2})", r"\1 \2", t)   # year glued to time
    t = re.sub(r"(?<=[A-Za-z])(?=\d)|(?<=\d)(?=[A-Za-z]{3})", " ", t)  # letters glued to digits
    return t


def read_stamp(bgr: np.ndarray, expected: dt.date, near: dt.time | None = None):
    """Read the stamp from the bottom part of the photo, then the top if needed.
    Returns (datetime|None, ampm_inferred, lat, lon, text, confidence)."""
    h = bgr.shape[0]
    texts, conf_used = [], 0.0
    dated = undated = None
    lat = lon = None
    for region in (bgr[int(h * 0.50):], bgr[: int(h * 0.28)]):
        lines = _ocr(region)
        for line, conf in lines:
            texts.append(line)
        for line, conf in lines:
            t = _normalise(line)
            when, inferred = parse_datetime(t, expected, near)
            if when is not None:
                if when.date() != dt.date.min and dated is None:
                    dated, conf_used = (when, inferred), conf
                elif when.date() == dt.date.min and undated is None:
                    undated = (when, inferred, conf)
        if dated is None:
            # date and time may sit on different lines of the same stamp
            joined = _normalise(" ".join(l for l, c in lines if c > 0.5))
            when, inferred = parse_datetime(joined, expected, near)
            if when is not None and when.date() != dt.date.min:
                dated, conf_used = (when, inferred), min((c for l, c in lines), default=0)
        la, lo = parse_coords(" ".join(_normalise(l) for l, c in lines))
        if la is not None and lat is None:
            lat, lon = la, lo
        if dated:
            break
    if dated:
        return dated[0], dated[1], lat, lon, _clean("\n".join(texts)), conf_used
    if undated:
        return undated[0], undated[1], lat, lon, _clean("\n".join(texts)), undated[2]
    return None, False, lat, lon, _clean("\n".join(texts)), 0.0


def _clean(t: str) -> str:
    lines = [ln.strip() for ln in t.splitlines() if re.search(r"\d", ln)]
    return " | ".join(dict.fromkeys(lines))[:400]


_DATE_NUM = re.compile(r"(?<!\d)(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})(?!\d)")
_DATE_TXT = re.compile(r"(?<!\d)(\d{1,2})(?:st|nd|rd|th)?[\s\-,]*([A-Za-z]{3,9})[\s\-,.]*(\d{4}|\d{2})(?!\d)")
_DATE_YMD = re.compile(r"(?<!\d)(20\d{2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})(?!\d)")
_MON_YEAR = re.compile(r"([A-Za-z]{3,9})[\s\-,.]*(\d{4})")
_DATE_TXT2 = re.compile(r"([A-Za-z]{3,9})[\s\-.]*(\d{1,2})(?:st|nd|rd|th)?\s*,?[\s\-]*(20\d{2})")
_TIME = re.compile(r"(?<![\d+:.])(\d{1,2})\s*[:.]\s*(\d{2})(?:\s*[:.]\s*(\d{2}))?(?![\d.])")
_AMPM = re.compile(r"^\W{0,2}\d?\s*([ap])\s*\.?\s*[r]?\s*[mn]", re.I)


def _dates(t: str, expected: dt.date) -> list[dt.date]:
    out = []
    for a, b, y in _DATE_NUM.findall(_DATE_YMD.sub(" ", t)):
        if len(y) == 4 and y[-2:] == str(expected.year)[-2:]:
            y = str(expected.year)                     # "2926" misread for 2026
        y = int(y) + (2000 if len(y) == 2 else 0)
        a, b = int(a), int(b)
        for d, m in ((a, b), (b, a)):                 # dd/mm or mm/dd
            try:
                out.append(dt.date(y, m, d))
            except ValueError:
                pass
    for y, m, d in _DATE_YMD.findall(t):
        try:
            out.append(dt.date(int(y), int(m), int(d)))
        except ValueError:
            pass
    t_no_ymd = _DATE_YMD.sub(" ", t)
    for d, mon, y in _DATE_TXT.findall(t_no_ymd):
        m = MONTHS.get(mon[:3].lower())
        if m:
            y = int(y) + (2000 if len(y) == 2 else 0)
            try:
                out.append(dt.date(y, m, int(d)))
            except ValueError:
                pass
    for mon, d, y in _DATE_TXT2.findall(t):
        m = MONTHS.get(mon[:3].lower())
        if m:
            try:
                out.append(dt.date(int(y), m, int(d)))
            except ValueError:
                pass
    # Day smudged but "Nov 2026" readable and matching the duty month: accept the duty date
    if not out:
        for mon, y in _MON_YEAR.findall(t):
            if MONTHS.get(mon[:3].lower()) == expected.month and int(y) == expected.year:
                out.append(expected)
                break
    # OCR sometimes misreads the year (2028 for 2026): keep day+month, fix the year
    fixed = []
    for d in out:
        if abs(d.year - expected.year) <= 3 and d.year != expected.year:
            try:
                alt = d.replace(year=expected.year)
                if abs((alt - expected).days) <= 2:
                    fixed.append(alt)
            except ValueError:
                pass
    return out + fixed


def parse_datetime(t: str, expected: dt.date, near: dt.time | None = None):
    """Return (datetime, ampm_inferred). Date part is date.min when no date was readable."""
    dates = [d for d in _dates(t, expected) if abs((d - expected).days) <= 60]
    times = []
    for m in _TIME.finditer(t):
        h, mi, s = int(m.group(1)), int(m.group(2)), m.group(3)
        if t[max(0, m.start() - 4):m.start()].upper().endswith(("GMT+", "GMT-", "UTC+", "+")):
            continue                                   # time-zone offset, not a time
        ap = _AMPM.search(t[m.end():m.end() + 7])
        inferred = False
        if ap:
            if ap.group(1).lower() == "p" and h < 12:
                h += 12
            if ap.group(1).lower() == "a" and h == 12:
                h = 0
        elif near is not None and 1 <= h <= 12 and not re.search(r"\b\d{1,2}\s*:\s*\d{2}\s*IST", t[m.start():m.end() + 5]):
            # no AM/PM readable: pick the 12-hour reading closest to the duty time
            def gap(hh):
                a = hh * 60 + mi
                b = near.hour * 60 + near.minute
                return min(abs(a - b), 1440 - abs(a - b))
            alt = (h + 12) % 24
            if gap(alt) < gap(h):
                h, inferred = alt, True
        if h < 24 and mi < 60:
            times.append((dt.time(h, mi, int(s) if s and int(s) < 60 else 0), inferred))
    if not times:
        return None, False
    tm, inferred = times[0]
    if not dates:
        return dt.datetime.combine(dt.date.min, tm), inferred
    d = min(dates, key=lambda x: abs((x - expected).days))
    return dt.datetime.combine(d, tm), inferred


_FLOAT = re.compile(r"(?<![\d.])(-?\d{1,3}\.\d{3,})")


def parse_coords(t: str) -> tuple[float | None, float | None]:
    nums = [float(x) for x in _FLOAT.findall(t.replace(",", " ").replace("°", " "))]
    for a, b in zip(nums, nums[1:]):
        if 0 <= abs(a) <= 90 and 0 <= abs(b) <= 180 and (abs(a) > 0.0 or abs(b) > 0.0):
            return a, b
    m = re.search(r"Lat[^\d-]*(-?\d+(?:\.\d+)?)[^\d-]+Long?[^\d-]*(-?\d+(?:\.\d+)?)", t, re.I)
    if m:
        return float(m.group(1)), float(m.group(2))
    return None, None


# ----------------------------------------------------------------------------- image tests

def brightness(bgr: np.ndarray) -> tuple[float, bool]:
    """Mean brightness of the upper 55% and whether a bright sky is visible at the top."""
    h = bgr.shape[0]
    upper = cv2.cvtColor(bgr[: int(h * 0.55)], cv2.COLOR_BGR2HSV)
    v = float(upper[..., 2].mean())
    top = cv2.cvtColor(bgr[: int(h * 0.12)], cv2.COLOR_BGR2HSV)
    # daylight sky: bright, and either bluish or whitish-grey, over most of the top strip
    hue, sat, val = top[..., 0], top[..., 1], top[..., 2]
    # only a clearly BLUE sky counts as daylight (white walls, fog and floodlit haze do not)
    sky = (val > 150) & (hue > 95) & (hue < 125) & (sat > 50)
    return v, bool(sky.mean() > 0.5)


def dhash(bgr: np.ndarray, size: int = 16) -> str:
    h = bgr.shape[0]
    g = cv2.cvtColor(bgr[: int(h * 0.65)], cv2.COLOR_BGR2GRAY)   # ignore the stamp area
    g = cv2.resize(g, (size + 1, size), interpolation=cv2.INTER_AREA)
    bits = (g[:, 1:] > g[:, :-1]).flatten()
    return "".join("1" if b else "0" for b in bits)


def same_place(raw_a: bytes, raw_b: bytes) -> int:
    """Number of matching image features between two photos (ignoring the stamp area).
    Two photos of the same spot from the same angle give a high count."""
    def feats(raw):
        g = cv2.imdecode(np.frombuffer(raw, np.uint8), cv2.IMREAD_GRAYSCALE)
        g = g[: int(g.shape[0] * 0.65)]
        g = cv2.resize(g, (600, int(600 * g.shape[0] / g.shape[1])))
        orb = cv2.ORB_create(1500)
        return orb.detectAndCompute(g, None)
    (ka, da), (kb, db) = feats(raw_a), feats(raw_b)
    if da is None or db is None or len(ka) < 20 or len(kb) < 20:
        return 0
    matches = cv2.BFMatcher(cv2.NORM_HAMMING).knnMatch(da, db, k=2)
    good = [m for m, n in (p for p in matches if len(p) == 2) if m.distance < 0.75 * n.distance]
    if len(good) < 12:
        return len(good)
    src = np.float32([ka[m.queryIdx].pt for m in good]).reshape(-1, 1, 2)
    dst = np.float32([kb[m.trainIdx].pt for m in good]).reshape(-1, 1, 2)
    _, mask = cv2.findHomography(src, dst, cv2.RANSAC, 6.0)
    return int(mask.sum()) if mask is not None else 0


def hamming(a: str, b: str) -> int:
    return sum(x != y for x, y in zip(a, b)) if a and b and len(a) == len(b) else 999


def km_between(lat1, lon1, lat2, lon2) -> float:
    r = 6371.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


# ----------------------------------------------------------------------------- main entry

def analyse(drive_id: str, role: str, expected: dt.date, cfg: dict, cache_dir: str,
            near: dt.time | None = None) -> PhotoResult:
    res = PhotoResult(drive_id=drive_id, role=role)
    raw, name = download(drive_id, cache_dir)
    if raw is None:
        res.add("PHOTO_NOT_OPENING", "high", "Photo link does not open (not shared, deleted or not an image)")
        return res
    res.ok_download = True
    res.file_name = name
    res.whatsapp_time = whatsapp_time(name)
    bgr = cv2.imdecode(np.frombuffer(raw, np.uint8), cv2.IMREAD_COLOR)
    if bgr is None:
        res.ok_download = False
        res.add("PHOTO_NOT_OPENING", "high", "File is not a readable image")
        return res

    when, inferred, lat, lon, text, conf = read_stamp(bgr, expected, near)
    res.ocr_conf = conf
    res.stamp_text = text
    res.lat, res.lon = lat, lon
    res.brightness, res.sky_bright = brightness(bgr)
    res.dhash = dhash(bgr)

    if when is None:
        if re.search(r"\d{1,2}\s*:\s*\d{2}", text) or _dates(text, expected):
            res.add("STAMP_UNREADABLE", "check", "Stamp present but could not be read reliably - check by eye")
        else:
            res.add("NO_TIME_STAMP", "high", "No date/time stamp on the photo")
    else:
        if when.date() == dt.date.min:
            when = dt.datetime.combine(expected, when.time())
            res.add("STAMP_NO_DATE", "medium", "Stamp shows a time but no readable date")
        res.stamp_time = when
        if inferred:
            res.add("AMPM_INFERRED", "low", "AM/PM not readable on the stamp; taken from the duty time")
    if lat is None:
        res.add("NO_COORDINATES", "low", "No GPS coordinates on the photo")
    else:
        box = cfg["valid_area"]
        if not (box["lat_min"] <= lat <= box["lat_max"] and box["lon_min"] <= lon <= box["lon_max"]):
            res.add("INVALID_COORDINATES", "high",
                    f"Coordinates {lat:.5f}, {lon:.5f} are outside the division (no GPS fix or faked)")
    return res
