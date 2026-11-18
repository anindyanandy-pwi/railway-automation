"""nbtrack (rail.nbtrack.in, NU Technology) - the SIM/IMEI GPS portal.

Uses the same REST API as the portal's own web page (api.nbtrack.in/api/v1):
  POST /auth/login                {username, password}                  -> token
  GET  /tracker/list-all                                                -> all trackers (name, IMEI, SIM...)
  POST /history/tracking/search   {imei, deviceId, gpsTimeFrom, gpsTimeTo} -> GPS points
  GET  /user/assigned-to-tracker                                        -> section users
  POST /reports/generate          {reportCode: GPS_EXCEPTION_REPORT, parameters}
"""
from __future__ import annotations

import datetime as dt
import logging
import re
import time
from zoneinfo import ZoneInfo

import requests

from .track import GpsSummary, Point, Track

log = logging.getLogger(__name__)
IST = ZoneInfo("Asia/Kolkata")
UTC = dt.timezone.utc

EXC_LABELS = {
    "lateStart": "Late Start",
    "earlyFinish": "Early Finish",
    "switchedOff": "Switched Off",
    "beatNotCoveredOrLessDistance": "Beat Not Covered / Less Distance",
    "unattendedOrNotStarted": "Unattended / Not Started",
    "highSpeedGreaterThan5Kmph": "High Speed > 5 km/h",
    "lowBatteryLessThan70": "Low Battery < 70%",
}


def _iso(t: dt.datetime) -> str:
    return t.replace(tzinfo=IST).astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def _parse_time(v) -> dt.datetime | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return dt.datetime.fromtimestamp(v / 1000 if v > 1e11 else v, UTC).astimezone(IST).replace(tzinfo=None)
    s = str(v).replace("Z", "+00:00")
    try:
        t = dt.datetime.fromisoformat(s)
    except ValueError:
        return None
    if t.tzinfo is None:
        t = t.replace(tzinfo=UTC)      # API times are UTC unless they say otherwise
    return t.astimezone(IST).replace(tzinfo=None)


def _items(data):
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for k in ("data", "content", "items", "trackers", "results"):
            if isinstance(data.get(k), list):
                return data[k]
            if isinstance(data.get(k), dict):
                inner = _items(data[k])
                if inner:
                    return inner
    return []


def _walk_strings(obj):
    if isinstance(obj, dict):
        for v in obj.values():
            yield from _walk_strings(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from _walk_strings(v)
    elif obj is not None:
        yield str(obj)


class Nbtrack:
    def __init__(self, api_url: str, user: str, password: str):
        self.api = api_url.rstrip("/")
        self.s = requests.Session()
        self.s.headers.update({"Content-Type": "application/json", "User-Agent": "Mozilla/5.0 (GPS cross-check)"})
        self.user, self.password = user, password
        self.trackers: list[dict] = []

    def login(self):
        r = self.s.post(f"{self.api}/auth/login", json={"username": self.user, "password": self.password}, timeout=60)
        if r.status_code >= 400:
            raise RuntimeError(f"nbtrack login failed (HTTP {r.status_code}) - check NBTRACK_USER / NBTRACK_PASSWORD")
        tok = r.json().get("token")
        if not tok:
            raise RuntimeError("nbtrack login returned no token")
        self.s.headers["Authorization"] = f"Bearer {tok}"
        log.info("nbtrack: signed in")

    def _call(self, method: str, path: str, tries: int = 3, **kw):
        for i in range(tries):
            try:
                r = self.s.request(method, f"{self.api}{path}", timeout=120, **kw)
                if r.status_code == 401:
                    self.login()
                    continue
                if r.status_code >= 500:
                    raise requests.HTTPError(f"HTTP {r.status_code}")
                r.raise_for_status()
                return r
            except requests.RequestException as e:
                log.warning("nbtrack %s failed (%s), retry %d", path, e, i + 1)
                time.sleep(5 * (i + 1))
        raise RuntimeError(f"nbtrack {path} not reachable")

    # ------------------------------------------------------------------ trackers
    def load_trackers(self):
        data = self._call("GET", "/tracker/list-all").json()
        rows = _items(data)
        out = []
        for t in rows:
            if not isinstance(t, dict):
                continue
            strings = list(_walk_strings(t))
            sims = [s for s in strings if re.fullmatch(r"5754\d{9}", s)]
            imeis = [s for s in strings if re.fullmatch(r"\d{15}", s)]
            name = next((str(v) for k, v in t.items() if "name" in k.lower() and isinstance(v, str)), "")
            out.append({"id": t.get("id") or t.get("deviceId") or t.get("trackerId"),
                        "imei": t.get("imei") or (imeis[0] if imeis else None),
                        "sim": sims[0] if sims else None, "name": name, "raw": t})
        self.trackers = out
        log.info("nbtrack: %d trackers listed", len(out))
        if out:
            log.info("nbtrack: tracker fields seen: %s", sorted(out[0]["raw"].keys()))

    def find(self, ref: str, kind: str) -> list[dict]:
        """Find a tracker by SIM, IMEI or the last digits the sheet sometimes uses."""
        if kind == "sim":
            return [t for t in self.trackers if t["sim"] == ref]
        if kind == "imei":
            return [t for t in self.trackers if t["imei"] == ref]
        return [t for t in self.trackers
                if (t["sim"] or "").endswith(ref) or (t["imei"] or "").endswith(ref)]

    # ------------------------------------------------------------------ GPS points
    def history(self, tracker: dict, start: dt.datetime, end: dt.datetime) -> Track:
        body = {"imei": tracker["imei"], "gpsTimeFrom": _iso(start), "gpsTimeTo": _iso(end)}
        if tracker.get("id"):
            body["deviceId"] = tracker["id"]
        data = self._call("POST", "/history/tracking/search", json=body).json()
        pts = []
        for p in _items(data):
            t = _parse_time(p.get("gpsTime") or p.get("time") or p.get("timestamp"))
            lat = p.get("latitude", p.get("lat"))
            lon = p.get("longitude", p.get("lng", p.get("lon")))
            if t is None or lat is None or lon is None:
                continue
            pts.append(Point(t, float(lat), float(lon), float(p.get("speed") or 0)))
        pts.sort(key=lambda p: p.t)
        return Track(pts)

    # ------------------------------------------------------------------ portal exceptions
    def exceptions(self, start: dt.datetime, end: dt.datetime) -> dict[str, list[str]]:
        """{tracker id or name: [exception names]} from the portal's GPS Exception Report."""
        try:
            users = _items(self._call("GET", "/user/assigned-to-tracker").json())
            user_ids = [u.get("id") for u in users if isinstance(u, dict) and u.get("id")]
            cats = set()
            for t in self.trackers:
                for k, v in t["raw"].items():
                    if "categor" in k.lower():
                        if isinstance(v, dict) and v.get("id"):
                            cats.add(v["id"])
                        elif isinstance(v, (str, int)) and "id" in k.lower():
                            cats.add(v)
            body = {"reportCode": "GPS_EXCEPTION_REPORT",
                    "parameters": {"fromDateTime": _iso(start), "toDateTime": _iso(end),
                                   "userIds": user_ids, "categoryIds": sorted(cats, key=str)}}
            r = self._call("POST", "/reports/generate", json=body)
            data = r.json()
            if r.status_code == 202:                       # big report: generated in background
                job_id = data.get("jobId")
                for _ in range(60):
                    time.sleep(10)
                    job = self._call("GET", f"/reports/jobs/{job_id}").json()
                    status = str(job.get("status", "")).upper()
                    if status in ("COMPLETED", "DONE", "SUCCESS"):
                        data = job.get("result") or job.get("data") or job
                        break
                    if status in ("FAILED", "ERROR"):
                        raise RuntimeError("report job failed")
        except Exception as e:  # noqa: BLE001
            log.warning("nbtrack: portal exception report not available (%s)", e)
            return {}
        out: dict[str, list[str]] = {}

        def walk(o):
            if isinstance(o, dict):
                det = o.get("details")
                if isinstance(det, dict):
                    for key, label in EXC_LABELS.items():
                        for item in det.get(key) or []:
                            if isinstance(item, dict):
                                k = str(item.get("trackerId") or item.get("trackerName"))
                                out.setdefault(k, [])
                                if label not in out[k]:
                                    out[k].append(label)
                                if item.get("trackerName"):
                                    out.setdefault("name:" + item["trackerName"], out[k])
                for v in o.values():
                    walk(v)
            elif isinstance(o, list):
                for v in o:
                    walk(v)
        walk(data)
        log.info("nbtrack: portal exceptions for %d devices", len({id(v) for v in out.values()}))
        return out


def summarise(tracker: dict | None, track: Track | None, ref: str) -> GpsSummary:
    s = GpsSummary(vendor="nbtrack", device_ref=ref)
    if not tracker:
        return s
    s.found = True
    s.device_name = tracker["name"]
    if track and track.points:
        s.track = track
        s.start, s.end = track.first, track.last
        s.km = track.km()
        s.max_speed = track.max_speed()
        hours = (s.end - s.start).total_seconds() / 3600 if s.end and s.start else 0
        s.avg_speed = round(s.km / hours, 2) if hours > 0 else None
        s.source = "history"
    return s
