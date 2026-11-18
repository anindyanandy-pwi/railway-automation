"""rilapp (ip3.rilapp.com/railways) - the 6-digit-ID GPS portal.

Plain web pages, so a normal HTTP session is enough:
  login.php                            -> sign in
  patrollingReportGPS.php              -> per device: start, end, km, avg speed
  vehicleMapGoogleFilterNewAPINew.php  -> list of devices, and the GPS points of one device
  exceptionReport.php                  -> portal's own exceptions (not done, off device, overspeed...)
  keymanGroupSummary.php               -> keymen "on device after 5 AM" / "off device before 3 PM"
"""
from __future__ import annotations

import datetime as dt
import json
import logging
import re
import time

import requests
from bs4 import BeautifulSoup

from .track import GpsSummary, Point, Track

log = logging.getLogger(__name__)
ID_IN_NAME = re.compile(r"#\s*(\d{6})")


def _dt(text: str) -> dt.datetime | None:
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?", text or "")
    if not m:
        return None
    d, mo, y, h, mi, s = m.groups()
    return dt.datetime(int(y), int(mo), int(d), int(h), int(mi), int(s or 0))


def _f(text: str) -> float | None:
    m = re.search(r"\d+(?:\.\d+)?", text or "")
    return float(m.group()) if m else None


class Rilapp:
    def __init__(self, base_url: str, user: str, password: str):
        self.base = base_url.rstrip("/") + "/"
        self.s = requests.Session()
        self.s.headers["User-Agent"] = "Mozilla/5.0 (GPS cross-check)"
        self.user, self.password = user, password
        self._devices: dict[str, tuple[str, str]] | None = None   # id -> (internal value, name)

    # ------------------------------------------------------------------ basics
    def login(self):
        r = self.s.post(self.base + "login.php", data={"username": self.user, "password": self.password},
                        timeout=60, allow_redirects=True)
        r.raise_for_status()
        if "logout.php" not in r.text and "dashboard" not in r.url:
            raise RuntimeError("rilapp login failed (check RILAPP_USER / RILAPP_PASSWORD)")
        log.info("rilapp: signed in")

    def _get(self, page: str, params: dict | None = None, tries: int = 3) -> str:
        for i in range(tries):
            try:
                r = self.s.get(self.base + page, params=params, timeout=180)
                r.raise_for_status()
                if "login.php" in r.url:
                    self.login()
                    continue
                return r.text
            except requests.RequestException as e:
                log.warning("rilapp %s failed (%s), retry %d", page, type(e).__name__, i + 1)
                time.sleep(5 * (i + 1))
        raise RuntimeError(f"rilapp page {page} not reachable")

    @staticmethod
    def _q(start: dt.datetime, end: dt.datetime) -> dict:
        return {"fdate": start.strftime("%d/%m/%Y"), "ftime": start.strftime("%H:%M"),
                "tdate": end.strftime("%d/%m/%Y"), "ttime": end.strftime("%H:%M"), "Submit": "Update"}

    # ------------------------------------------------------------------ devices
    def devices(self) -> dict[str, tuple[str, str]]:
        if self._devices is None:
            html = self._get("vehicleMapGoogleFilterNewAPINew.php")
            soup = BeautifulSoup(html, "html.parser")
            out = {}
            for opt in soup.select("select[name=deviceid] option"):
                m = ID_IN_NAME.search(opt.text)
                if m and opt.get("value"):
                    out[m.group(1)] = (opt["value"], opt.text.split("#")[0].strip())
            self._devices = out
            log.info("rilapp: %d devices listed", len(out))
        return self._devices

    # ------------------------------------------------------------------ patrolling report
    def patrolling(self, start: dt.datetime, end: dt.datetime, category: str) -> dict[str, dict]:
        """category: '-KM' keymen, '-PM' patrolmen. Returns {device id: row}."""
        q = self._q(start, end) | {"category": category}
        soup = BeautifulSoup(self._get("patrollingReportGPS.php", q), "html.parser")
        out = {}
        table = soup.find("table")
        for tr in table.find_all("tr") if table else []:
            c = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
            if len(c) < 9:
                continue
            m = ID_IN_NAME.search(c[1])
            if not m:
                continue
            out[m.group(1)] = {
                "section": c[0], "name": c[1].split("#")[0].strip(),
                "start": _dt(c[2]), "start_addr": c[3], "end": _dt(c[4]), "end_addr": c[5],
                "km": _f(c[6]), "overspeed": c[7], "avg_speed": _f(c[8]),
                "low_batt": c[9] if len(c) > 9 else "",
            }
        log.info("rilapp: patrolling report %s %s-%s: %d devices", category, start, end, len(out))
        return out

    # ------------------------------------------------------------------ GPS points
    def history(self, device_id: str, start: dt.datetime, end: dt.datetime) -> Track | None:
        dev = self.devices().get(device_id)
        if not dev:
            return None
        q = self._q(start, end) | {"category": "all", "deviceid": dev[0]}
        html = self._get("vehicleMapGoogleFilterNewAPINew.php", q)
        i = html.find('{"points":')
        if i < 0:
            return Track([])
        # find the matching closing brace of this JSON object
        depth, j = 0, i
        for j in range(i, len(html)):
            ch = html[j]
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    break
        try:
            data = json.loads(html[i:j + 1])
        except json.JSONDecodeError:
            return Track([])
        pts = []
        for p in data.get("points", []):
            t = _dt(p.get("time", ""))
            if t and p.get("latitude") is not None:
                pts.append(Point(t, float(p["latitude"]), float(p["longitude"]), float(p.get("speed") or 0)))
        pts.sort(key=lambda p: p.t)
        return Track(pts)

    # ------------------------------------------------------------------ portal exceptions
    def exceptions(self, start: dt.datetime, end: dt.datetime, category: str) -> dict[str, list[str]]:
        """{device id: [exception names]} from exceptionReport.php (and keyman summary for -KM)."""
        out: dict[str, list[str]] = {}

        def add(dev_id, name):
            out.setdefault(dev_id, [])
            if name not in out[dev_id]:
                out[dev_id].append(name)

        q = self._q(start, end) | {"category": category, "mindistance": "3"}
        soup = BeautifulSoup(self._get("exceptionReport.php", q), "html.parser")
        table = soup.find("table")
        if table:
            heads = [th.get_text(" ", strip=True) for th in table.find("tr").find_all(["th", "td"])]
            for tr in table.find_all("tr")[1:]:
                cells = tr.find_all("td")
                for head, td in zip(heads, cells):
                    h = head.lower()
                    label = ("Patrolling Not Done" if "not done" in h else
                             "Off Device" if "off device" in h else
                             "Start Low Battery" if "low batt" in h else
                             "Overspeed" if "overspeed" in h else None)
                    if label:
                        for m in ID_IN_NAME.finditer(td.get_text(" ", strip=True)):
                            add(m.group(1), label)

        if category == "-KM":
            names = {name: dev_id for dev_id, (_v, name) in self.devices().items()}
            q = self._q(start, end)
            q.pop("tdate")
            soup = BeautifulSoup(self._get("keymanGroupSummary.php", q), "html.parser")
            table = soup.find("table")
            if table:
                heads = [th.get_text(" ", strip=True) for th in table.find("tr").find_all(["th", "td"])]
                for tr in table.find_all("tr")[1:]:
                    for head, td in zip(heads, tr.find_all("td")):
                        h = head.lower()
                        label = ("On device after 5 AM (late start)" if "after" in h else
                                 "Off device before 3 PM (early finish)" if "before" in h else
                                 "Less than 4 km travelled" if "less than 4" in h else
                                 "Less than 2 hours working" if "less than 2" in h else None)
                        if not label:
                            continue
                        for token in td.stripped_strings:
                            m = ID_IN_NAME.search(token)
                            if m:
                                add(m.group(1), label)
                                continue
                            base = re.sub(r"-\d{1,2}$", "", token.strip())   # name + "-HH"
                            if base in names:
                                add(names[base], label)
        log.info("rilapp: portal exceptions %s: %d devices", category, len(out))
        return out


def summarise(report_row: dict | None, track: Track | None, device_id: str) -> GpsSummary:
    s = GpsSummary(vendor="rilapp", device_ref=device_id)
    if report_row:
        s.found = True
        s.device_name = report_row["name"]
        s.start, s.end = report_row["start"], report_row["end"]
        s.km, s.avg_speed = report_row["km"], report_row["avg_speed"]
        s.source = "report"
    if track and track.points:
        s.found = True
        s.track = track
        s.max_speed = track.max_speed()
        if not report_row:
            s.start, s.end, s.km = track.first, track.last, track.km()
            s.source = "history"
        else:
            s.source = "report+history"
    return s
