"""Read the control team's Google Sheet (public Excel export).

Each tab has fixed columns on the left (Sr DEN, AEN, PWI, beat...) and then one
block of columns per date, newest block first. We find the block for the date
we want by its header ("Date :- 18.11.2026" / "Date-17/18.11.2026") and read
columns by their header text, so a column added or moved in the sheet does not
break the reader.
"""
from __future__ import annotations

import datetime as dt
import logging
import re
from dataclasses import dataclass, field

import warnings

import openpyxl
import openpyxl.worksheet._reader as _xl_reader
import requests

log = logging.getLogger(__name__)
warnings.filterwarnings("ignore", module="openpyxl")

# Some GPS-number cells are formatted as dates in the sheet. openpyxl would turn
# those big numbers into "#VALUE!"; keep the raw number instead.
_orig_from_excel = _xl_reader.from_excel


def _safe_from_excel(value, *a, **k):
    try:
        return _orig_from_excel(value, *a, **k)
    except (OverflowError, ValueError):
        return value


_xl_reader.from_excel = _safe_from_excel

EXPORT_URL = "https://docs.google.com/spreadsheets/d/{id}/export?format=xlsx"

RILAPP_ID = re.compile(r"(?<!\d)(1[0-9]\d{4})(?!\d)")          # 6-digit, e.g. 120084 / 110559
NBTRACK_SIM = re.compile(r"(?<!\d)(5754\d{9})(?!\d)")          # 13-digit SIM no, e.g. 5754130017736
NBTRACK_IMEI = re.compile(r"(?<!\d)(86\d{13})(?!\d)")           # 15-digit IMEI, e.g. 861896070008650
NBTRACK_TAIL = re.compile(r"(?<!\d)(\d{3,6})\s*\(\s*NB\s*\)", re.I)  # "7122(NB)" = last digits of an NB device
BARE_TAIL = re.compile(r"(?<![\d.])(\d{4,5})(?![\d.])")         # "18028" = last 5 digits of an NB SIM
DRIVE_ID = re.compile(r"(?:file/d/|[?&]id=)([\w-]{20,})")


@dataclass
class Device:
    vendor: str | None      # "rilapp" | "nbtrack" | None
    ref: str | None         # rilapp device id, or nbtrack SIM / IMEI / tail digits
    raw: str                # what the sheet says, as written
    replaced: bool = False
    ref_kind: str = ""      # for nbtrack: "sim" | "imei" | "tail"
    no_gps: bool = False    # the sheet says the batch has no GPS


@dataclass
class SheetRow:
    kind: str                       # "keyman" | "cwp"
    duty_date: dt.date              # keyman: the day; cwp: the evening the night starts
    row_no: int                     # row number in the sheet (for the dashboard link)
    sr_den: str = ""
    aen: str = ""
    pwi: str = ""
    beat: str = ""                  # keyman beat / cwp batch + location
    names: str = ""
    device: Device | None = None
    roster: str = ""
    roster_start: dt.time | None = None
    roster_end: dt.time | None = None
    start_time: dt.time | None = None
    end_time: dt.time | None = None
    km_target: float | None = None
    km_done: float | None = None
    km_raw: str = ""
    max_speed: float | None = None
    avg_speed: float | None = None
    speed_raw: str = ""
    call1: dt.time | None = None
    call2: dt.time | None = None
    start_photos: list[str] = field(default_factory=list)
    equipment_photos: list[str] = field(default_factory=list)
    end_photos: list[str] = field(default_factory=list)
    gps_checked: str = ""
    analysis: str = ""
    not_applicable: bool = False    # the row is "---" for the whole day (no duty)


# ----------------------------------------------------------------------------- helpers

def _s(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def parse_time(v) -> dt.time | None:
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.time().replace(microsecond=0)
    if isinstance(v, dt.time):
        return v.replace(microsecond=0)
    if isinstance(v, (int, float)) and 0 <= v < 1:          # Excel fraction of a day
        secs = round(v * 86400)
        return dt.time(secs // 3600 % 24, secs // 60 % 60)
    m = re.search(r"(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?\s*([AaPp][Mm])?", str(v))
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    ap = (m.group(4) or "").lower()
    if ap == "pm" and h < 12:
        h += 12
    if ap == "am" and h == 12:
        h = 0
    if h > 23 or mi > 59:
        return None
    return dt.time(h, mi)


def parse_float(v) -> float | None:
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"\d+(?:\.\d+)?", _s(v))
    return float(m.group()) if m else None


def parse_roster(text: str) -> tuple[dt.time | None, dt.time | None]:
    times = re.findall(r"\d{1,2}[:.]\d{2}", text or "")
    if len(times) >= 2:
        return parse_time(times[0]), parse_time(times[1])
    return None, None


def parse_device(text: str) -> Device:
    raw = _s(text)
    up = raw.upper()
    # If the cell says "X REPLACED BY Y", the device in use is the one after REPLACED BY.
    part = up.split("REPLACED BY", 1)[1] if "REPLACED BY" in up else up
    part = re.sub(r"(\d)\.0\b", r"\1", part)
    found: list[tuple[int, Device]] = []
    for rx, vendor, kind in ((NBTRACK_SIM, "nbtrack", "sim"), (NBTRACK_IMEI, "nbtrack", "imei"),
                             (NBTRACK_TAIL, "nbtrack", "tail")):
        for m in rx.finditer(part):
            found.append((m.start(), Device(vendor, m.group(1), raw, ref_kind=kind)))
    masked = NBTRACK_TAIL.sub(" ", NBTRACK_IMEI.sub(" ", NBTRACK_SIM.sub(" ", part)))
    for m in RILAPP_ID.finditer(masked):
        found.append((m.start(), Device("rilapp", m.group(1), raw)))
    if not found and "OLD" not in part:
        for m in BARE_TAIL.finditer(masked):
            found.append((m.start(), Device("nbtrack", m.group(1), raw, ref_kind="tail")))
    if found:
        dev = min(found, key=lambda x: x[0])[1]
        dev.replaced = "REPLACED BY" in up
        return dev
    return Device(None, None, raw,
                  no_gps=bool(re.search(r"NO\s*GPS|SUBMITTED|OLD", up)))


def drive_ids(v) -> list[str]:
    return list(dict.fromkeys(DRIVE_ID.findall(_s(v))))


# ----------------------------------------------------------------------------- workbook

def download(sheet_id: str, dest: str) -> str:
    log.info("Downloading sheet export")
    with requests.get(EXPORT_URL.format(id=sheet_id), stream=True, timeout=600) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(1 << 20):
                f.write(chunk)
    return dest


def _block_dates(cell_text: str, kind: str) -> dt.date | None:
    nums = [int(n) for n in re.findall(r"\d+", cell_text)]
    if len(nums) < 3:
        return None
    d, m, y = nums[-3], nums[-2], nums[-1]
    if y < 100:
        y += 2000
    try:
        day = dt.date(y, m, d)
    except ValueError:
        return None
    # CWP header "17/18.11.2026" names the night by both days: return the evening's date.
    if kind == "cwp" and len(nums) >= 4:
        day -= dt.timedelta(days=1)
    return day


HEADER_KEYS = {
    # internal name: words that identify the header (lower case, all must appear)
    "names": [["name"]],
    "phone": [["ph"]],
    "device": [["gps", "number"]],
    "roster": [["roster"]],
    "start_time": [["start", "time"], ["strating", "time"]],
    "start_photo": [["starting", "picture"], ["start", "picture"]],
    "equipment": [["equipment"]],
    "call1": [["1st", "call"]],
    "call2": [["2nd", "call"]],
    "end_time": [["end", "time"]],
    "end_photo": [["end", "picture"]],
    "km_target": [["distance", "to be"]],
    "km_done": [["distance", "traveled"], ["distance", "travelled"]],
    "gps_checked": [["gps", "checked"]],
    "analysis": [["overall", "analysis"]],
    "speed": [["speed"]],
}


def _classify(header: str) -> str | None:
    h = re.sub(r"\s+", " ", header.lower())
    for key, alternatives in HEADER_KEYS.items():
        for words in alternatives:
            if all(w in h for w in words):
                if key == "km_done" and "to be" in h:
                    continue
                if key == "start_time" and "picture" in h:
                    continue
                return key
    return None


def read_tab(path: str, tab: str, kind: str, wanted: dt.date) -> list[SheetRow]:
    """Return all rows of one tab for one date block."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    ws = wb[tab]
    rows = ws.iter_rows(values_only=True)
    top = [list(next(rows)) for _ in range(6)]

    # 1. find the date row and the block start column for the wanted date
    date_row = block_col = None
    for ri, r in enumerate(top):
        for ci, v in enumerate(r):
            t = _s(v)
            if "date" in t.lower() and re.search(r"\d{4}", t):
                if _block_dates(t, kind) == wanted:
                    date_row, block_col = ri, ci
        if block_col is not None:
            break
    if block_col is None:
        wb.close()
        raise LookupError(f"No block for {wanted:%d.%m.%Y} in tab '{tab}'")

    # 2. header row = first row below the date row that has "starting" somewhere
    hdr_idx = next(i for i in range(date_row + 1, 6)
                   if any("start" in _s(v).lower() for v in top[i]))
    header = top[hdr_idx]

    # block = from block_col until the next "date" cell in the date row (or end)
    next_dates = [ci for ci, v in enumerate(top[date_row])
                  if ci > block_col and "date" in _s(v).lower()]
    block_end = next_dates[0] if next_dates else len(header)
    # Keyman blocks start one column after the name column in some layouts:
    # look back 1 column if the date cell sits on a merged range start.
    cols: dict[str, int] = {}
    for ci in range(max(0, block_col - 1), block_end):
        key = _classify(_s(header[ci]))
        if key and key not in cols:
            cols[key] = ci
    # fixed columns on the left (before the first block)
    first_block = min(ci for ci, v in enumerate(top[date_row]) if "date" in _s(v).lower())
    left: dict[str, int] = {}
    for ci in range(0, first_block):
        h = _s(header[ci]).lower()
        if "sr" in h and "den" in h or h.startswith("sr.dn") or h == "sr.dn":
            left["sr_den"] = ci
        elif h == "aen":
            left["aen"] = ci
        elif h == "pwi":
            left["pwi"] = ci
        elif "beat" in h or "batch" in h:
            left["beat"] = ci
        elif "location" in h:
            left["location"] = ci
        else:
            k = _classify(h)
            if k and k not in cols:
                cols[k] = ci          # e.g. CWP keeps GPS number and roster on the left

    out: list[SheetRow] = []
    carry = {"sr_den": "", "aen": "", "pwi": ""}
    all_rows = top[hdr_idx + 1:] + [list(r) for r in rows]
    for offset, r in enumerate(all_rows):
        row_no = hdr_idx + 2 + offset
        get = lambda k, src=cols: r[src[k]] if k in src and src[k] < len(r) else None   # noqa: E731
        for k in carry:
            if k in left and _s(r[left[k]]):
                carry[k] = _s(r[left[k]])
        beat = _s(r[left["beat"]]) if "beat" in left else ""
        if "location" in left:
            beat = (beat + " " + _s(r[left["location"]])).strip()
        dev_text = _s(get("device"))
        if not beat and not dev_text:
            continue
        roster = _s(get("roster"))
        rs, re_ = parse_roster(roster)
        sp_raw = _s(get("speed"))
        sp_nums = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", sp_raw)]
        row = SheetRow(
            kind=kind, duty_date=wanted, row_no=row_no,
            sr_den=carry["sr_den"], aen=carry["aen"], pwi=carry["pwi"], beat=beat,
            names=_s(get("names")), device=parse_device(dev_text),
            roster=roster, roster_start=rs, roster_end=re_,
            start_time=parse_time(get("start_time")), end_time=parse_time(get("end_time")),
            km_target=parse_float(get("km_target")),
            km_done=parse_float(get("km_done")), km_raw=_s(get("km_done")),
            speed_raw=sp_raw,
            call1=parse_time(get("call1")), call2=parse_time(get("call2")),
            start_photos=drive_ids(get("start_photo")),
            equipment_photos=drive_ids(get("equipment")),
            end_photos=drive_ids(get("end_photo")),
            gps_checked=_s(get("gps_checked")), analysis=_s(get("analysis")),
        )
        if len(sp_nums) >= 2:
            row.max_speed, row.avg_speed = sp_nums[0], sp_nums[1]
        elif len(sp_nums) == 1:
            row.avg_speed = sp_nums[0]
        block_vals = [_s(get(k)) for k in ("start_time", "end_time", "km_done")]
        row.not_applicable = all(v in ("", "---", "-") for v in block_vals) and not row.start_photos
        out.append(row)
    wb.close()
    log.info("Sheet %s %s: %d rows", kind, wanted, len(out))
    return out
