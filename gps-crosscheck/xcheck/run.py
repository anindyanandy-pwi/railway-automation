"""Morning run: python -m xcheck.run [--date YYYY-MM-DD]

Run on day D it checks:
  * Cold Weather Patrolling, night D-1 -> D   (full)
  * Keyman, day D                             (start only)
  * Keyman, day D-1                           (whole day - completes the previous morning's check)

Logs show counts only - no names or phone numbers (the GitHub repo is public).
Results go to the results store (Apps Script web app) and nowhere else.
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import datetime as dt
import json
import logging
import os
import sys
import tempfile
import time
from zoneinfo import ZoneInfo

import requests
import yaml

from . import checks, photos, sheet
from .nbtrack import Nbtrack
from .nbtrack import summarise as nb_summary
from .rilapp import Rilapp
from .rilapp import summarise as ril_summary

log = logging.getLogger("xcheck")
CACHE = os.path.join(tempfile.gettempdir(), "xcheck-photos")


# ----------------------------------------------------------------------------- settings / windows

def load_settings(cfg: dict, results_url: str | None, key: str | None) -> dict:
    st = dict(cfg["defaults"])
    if results_url:
        try:
            r = requests.get(results_url, params={"action": "settings", "key": key}, timeout=60)
            remote = r.json().get("settings", {})
            for k, v in remote.items():
                if k in st and v not in ("", None):
                    st[k] = type(st[k])(v)
            log.info("Using dashboard settings")
        except Exception as e:  # noqa: BLE001
            log.warning("Dashboard settings not readable (%s) - using defaults", type(e).__name__)
    return st


def window(cfg: dict, name: str, day: dt.date) -> tuple[dt.datetime, dt.datetime]:
    w = cfg["windows"][name]
    a = dt.datetime.combine(day + dt.timedelta(days=w["from_day"]), dt.time.fromisoformat(w["from"]))
    b = dt.datetime.combine(day + dt.timedelta(days=w["to_day"]), dt.time.fromisoformat(w["to"]))
    return a, b


# ----------------------------------------------------------------------------- photos (parallel)

def _photo_job(args):
    drive_id, role, expected, near, cfg = args
    return photos.analyse(drive_id, role, expected, cfg, CACHE, near)


def analyse_photos(jobs: list[tuple], workers: int) -> dict[tuple[str, str], photos.PhotoResult]:
    out = {}
    with cf.ProcessPoolExecutor(max_workers=workers) as ex:
        for res in ex.map(_photo_job, jobs, chunksize=4):
            out[(res.drive_id, res.role)] = res
    return out


def _raw(drive_id: str) -> bytes:
    data, _ = photos.download(drive_id, CACHE)
    return data or b""


# ----------------------------------------------------------------------------- output

def _t(x):
    if isinstance(x, dt.datetime):
        return x.strftime("%d.%m.%Y %H:%M")
    if isinstance(x, dt.time):
        return x.strftime("%H:%M")
    return x


def to_record(run_date: dt.date, r: checks.RowResult, sheet_id: str, gid: dict) -> dict:
    row, g = r.row, r.gps
    link = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit#gid={gid.get(row.kind, 0)}&range=A{row.row_no}"
    ph = {p.role: p for p in r.photos}
    return {
        "run_date": str(run_date),
        "check": {"cwp": "CWP night", "keyman_start": "Keyman start", "keyman_full": "Keyman full day"}[r.mode],
        "duty_date": str(row.duty_date),
        "status": r.status,
        "sr_den": row.sr_den, "aen": row.aen, "pwi": row.pwi, "beat": row.beat.replace("\n", " "),
        "names": row.names.replace("\n", ", "),
        "gps_in_sheet": (row.device.raw if row.device else "").replace("\n", " "),
        "vendor": row.device.vendor if row.device else "",
        "portal_device": g.device_name if g else "",
        "roster": row.roster,
        "sheet_start": _t(row.start_time), "sheet_end": _t(row.end_time),
        "sheet_km": row.km_raw, "sheet_speed": row.speed_raw,
        "gps_start": _t(g.start) if g else "", "gps_end": _t(g.end) if g else "",
        "gps_km": g.km if g else "", "gps_max_speed": g.max_speed if g else "", "gps_avg_speed": g.avg_speed if g else "",
        "start_photo_time": _t(ph["start"].stamp_time) if "start" in ph else "",
        "end_photo_time": _t(ph["end"].stamp_time) if "end" in ph else "",
        "start_photo": f"https://drive.google.com/file/d/{ph['start'].drive_id}/view" if "start" in ph else "",
        "end_photo": f"https://drive.google.com/file/d/{ph['end'].drive_id}/view" if "end" in ph else "",
        "actual_start": _t(r.actual_start), "actual_start_from": r.actual_start_source,
        "actual_end": _t(r.actual_end),
        "findings": [{"code": f.code, "severity": f.severity, "area": f.area, "message": f.message} for f in r.findings],
        "portal_exceptions": [{"exception": v.exception, "verdict": v.verdict, "reason": v.reason} for v in r.verdicts],
        "photo_hashes": [{"drive_id": p.drive_id, "role": p.role, "dhash": p.dhash} for p in r.photos if p.dhash],
        "sheet_link": link,
    }


# ----------------------------------------------------------------------------- main

def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default=os.environ.get("DIVISION_CONFIG", "config/howrah.yaml"))
    ap.add_argument("--date", help="run date D (default: today in division time)")
    ap.add_argument("--only", choices=["cwp", "keyman"], help="check one tab only")
    ap.add_argument("--limit", type=int, default=0, help="first N rows per check (testing)")
    ap.add_argument("--no-portals", action="store_true", help="skip GPS portals (testing photos/sheet only)")
    ap.add_argument("--sheet-file", help="use an already downloaded sheet export (testing)")
    ap.add_argument("--out", help="also write results JSON to this file (local testing only)")
    ap.add_argument("--workers", type=int, default=max(2, os.cpu_count() or 2))
    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    t0 = time.time()

    cfg = yaml.safe_load(open(a.config, encoding="utf-8"))
    tz = ZoneInfo(cfg["timezone"])
    run_date = dt.date.fromisoformat(a.date) if a.date else dt.datetime.now(tz).date()
    results_url, key = os.environ.get("RESULTS_URL"), os.environ.get("RESULTS_KEY")
    st = load_settings(cfg, results_url, key)
    log.info("Run date %s, division %s", run_date, cfg["division"])

    # ---- sheet
    path = a.sheet_file or sheet.download(cfg["sheet"]["id"], os.path.join(tempfile.gettempdir(), "sheet.xlsx"))
    plan = []    # (mode, rows, window)
    yday = run_date - dt.timedelta(days=1)
    if a.only in (None, "cwp"):
        plan.append(("cwp", sheet.read_tab(path, cfg["sheet"]["cwp_tab"], "cwp", yday), window(cfg, "cwp_night", run_date)))
    if a.only in (None, "keyman"):
        plan.append(("keyman_start", sheet.read_tab(path, cfg["sheet"]["keyman_tab"], "keyman", run_date),
                     window(cfg, "keyman_start", run_date)))
        plan.append(("keyman_full", sheet.read_tab(path, cfg["sheet"]["keyman_tab"], "keyman", yday),
                     window(cfg, "keyman_full", run_date)))
    if a.limit:
        plan = [(m, rows[: a.limit], w) for m, rows, w in plan]

    # ---- portals
    ril = nb = None
    if not a.no_portals:
        pc = cfg["portals"]
        ril = Rilapp(pc["rilapp"]["base_url"], os.environ[pc["rilapp"]["user_env"]],
                     os.environ[pc["rilapp"]["password_env"]])
        nb = Nbtrack(pc["nbtrack"]["api_url"], os.environ[pc["nbtrack"]["user_env"]],
                     os.environ[pc["nbtrack"]["password_env"]])
        for name, portal in (("rilapp", ril), ("nbtrack", nb)):
            try:
                portal.login()
                if portal is nb:
                    nb.load_trackers()
                else:
                    ril.devices()
            except Exception as e:  # noqa: BLE001
                log.error("%s not available: %s", name, e)
                if portal is ril:
                    ril = None
                else:
                    nb = None

    # ---- photos (all at once, in parallel)
    jobs, seen = [], set()
    for mode, rows, _w in plan:
        for r in rows:
            clock = checks.Clock(r)
            for role, ids, near in (("start", r.start_photos, r.roster_start),
                                    ("equipment", r.equipment_photos if mode == "cwp" else [], r.roster_start),
                                    ("end", r.end_photos if mode == "cwp" else [], r.roster_end)):
                for fid in ids[:1]:
                    exp = r.duty_date
                    if mode == "cwp" and role == "end":
                        exp = r.duty_date + dt.timedelta(days=1)
                    if (fid, role) not in seen:
                        seen.add((fid, role))
                        jobs.append((fid, role, exp, near, cfg))
            del clock
    log.info("Reading %d photos with %d workers", len(jobs), a.workers)
    ph = analyse_photos(jobs, a.workers)
    log.info("Photos done in %.0f s", time.time() - t0)

    # ---- per check
    gid = {"cwp": 1182515923, "keyman": 31130341}
    results: list[checks.RowResult] = []
    for mode, rows, (w_from, w_to) in plan:
        cat = "-PM" if mode == "cwp" else "-KM"
        ril_rep = ril_exc = {}
        nb_exc = {}
        if ril:
            try:
                ril_rep = ril.patrolling(w_from, w_to, cat)
                ril_exc = ril.exceptions(w_from, w_to, cat)
            except Exception as e:  # noqa: BLE001
                log.error("rilapp reports failed for %s: %s", mode, e)
        if nb:
            nb_exc = nb.exceptions(w_from, w_to)
        n_gps = 0
        for r in rows:
            dev = r.device
            gps, exc = None, []
            if dev and dev.vendor == "rilapp" and ril:
                track = None
                try:
                    track = ril.history(dev.ref, w_from, w_to)
                except Exception as e:  # noqa: BLE001
                    log.warning("rilapp history failed (%s)", type(e).__name__)
                gps = ril_summary(ril_rep.get(dev.ref), track, dev.ref)
                if dev.ref in ril.devices():
                    gps.found = True
                    gps.device_name = gps.device_name or ril.devices()[dev.ref][1]
                exc = ril_exc.get(dev.ref, [])
            elif dev and dev.vendor == "nbtrack" and nb:
                hits = nb.find(dev.ref, dev.ref_kind)
                tracker = hits[0] if len(hits) == 1 else None
                track = None
                if tracker:
                    try:
                        track = nb.history(tracker, w_from, w_to)
                    except Exception as e:  # noqa: BLE001
                        log.warning("nbtrack history failed (%s)", type(e).__name__)
                gps = nb_summary(tracker, track, dev.ref)
                if tracker:
                    exc = nb_exc.get(str(tracker["id"]), []) or nb_exc.get("name:" + tracker["name"], [])
            if gps and (gps.start or (gps.track and gps.track.points)):
                n_gps += 1
            row_photos = []
            for role, ids in (("start", r.start_photos), ("equipment", r.equipment_photos), ("end", r.end_photos)):
                if ids and (ids[0], role) in ph:
                    row_photos.append(ph[(ids[0], role)])
            results.append(checks.evaluate(r, gps, row_photos, mode, st, cfg, exc))
        log.info("%s: %d rows, %d with GPS data", mode, len(rows), n_gps)

    # ---- cross-row photo checks
    hist = []
    if results_url:
        try:
            hist = requests.get(results_url, params={"action": "hashes", "key": key, "days": 30},
                                timeout=120).json().get("hashes", [])
        except Exception:  # noqa: BLE001
            log.warning("Old photo fingerprints not available")
    checks.cross_row(results, hist, st, photos.same_place, _raw)

    records = [to_record(run_date, r, cfg["sheet"]["id"], gid) for r in results]
    counts = {}
    for rec in records:
        counts[rec["status"]] = counts.get(rec["status"], 0) + 1
    log.info("Result: %s (%.0f s)", counts, time.time() - t0)

    payload = {"run_date": str(run_date), "division": cfg["division"],
               "generated_at": dt.datetime.now(tz).strftime("%d.%m.%Y %H:%M"),
               "settings": st, "records": records}
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=1, default=str)
    if results_url and not a.out:
        r = requests.post(results_url, params={"key": key}, data=json.dumps(payload, default=str),
                          headers={"Content-Type": "application/json"}, timeout=600)
        ok = r.ok and r.json().get("ok")
        log.info("Results store: %s", "saved" if ok else f"FAILED ({r.status_code})")
        if not ok:
            sys.exit(1)


if __name__ == "__main__":
    main()
