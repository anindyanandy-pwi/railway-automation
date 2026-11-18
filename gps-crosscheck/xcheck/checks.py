"""The cross-check rules.

For one sheet row (one keyman beat or one CWP batch for one duty) we compare
three sources - the sheet, the GPS portal and the photos - and:
  1. list every disagreement as a finding (with the evidence), and
  2. give each exception the portal itself raised a verdict:
     CONFIRMED (it is real), FALSE (the evidence shows it is wrong) or CANNOT VERIFY.

Severity: high (red) / medium (amber) / low (info) / check (needs a human look).
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from .photos import PhotoResult, km_between
from .sheet import SheetRow
from .track import GpsSummary

SEV_ORDER = {"high": 3, "medium": 2, "check": 1, "low": 0}


@dataclass
class Finding:
    code: str
    severity: str
    message: str
    area: str            # device | sheet | duty | photo | portal


@dataclass
class Verdict:
    exception: str       # as the portal names it
    verdict: str         # CONFIRMED | FALSE | CANNOT VERIFY
    reason: str


@dataclass
class RowResult:
    mode: str
    row: SheetRow
    gps: GpsSummary | None
    photos: list[PhotoResult]
    findings: list[Finding] = field(default_factory=list)
    verdicts: list[Verdict] = field(default_factory=list)
    actual_start: dt.datetime | None = None
    actual_start_source: str = ""
    actual_end: dt.datetime | None = None

    def add(self, code, severity, message, area):
        self.findings.append(Finding(code, severity, message, area))

    @property
    def status(self) -> str:
        worst = max((SEV_ORDER[f.severity] for f in self.findings), default=-1)
        return {3: "RED", 2: "AMBER", 1: "CHECK", 0: "OK"}.get(worst, "OK")


def hm(t: dt.datetime | dt.time | None) -> str:
    return t.strftime("%H:%M") if t else "-"


def mins(a: dt.datetime, b: dt.datetime) -> int:
    return round((a - b).total_seconds() / 60)


class Clock:
    """Places sheet times (just HH:MM) on the right calendar day for a duty."""

    def __init__(self, row: SheetRow):
        self.day = row.duty_date
        self.night = row.kind == "cwp" or (row.roster_start and row.roster_end and row.roster_end < row.roster_start)

    def at(self, t: dt.time | None) -> dt.datetime | None:
        if t is None:
            return None
        d = self.day
        if self.night and t < dt.time(12, 0):
            d += dt.timedelta(days=1)
        return dt.datetime.combine(d, t)


# ============================================================================= main

def evaluate(row: SheetRow, gps: GpsSummary | None, photos: list[PhotoResult], mode: str,
             st: dict, cfg: dict, portal_exc: list[str]) -> RowResult:
    """mode: 'cwp' (full night), 'keyman_start' (start only), 'keyman_full' (whole day)."""
    res = RowResult(mode=mode, row=row, gps=gps, photos=photos)
    clock = Clock(row)
    start_only = mode == "keyman_start"
    r_start, r_end = clock.at(row.roster_start), clock.at(row.roster_end)
    tol_s = dt.timedelta(minutes=st["start_tolerance_min"])
    tol_e = dt.timedelta(minutes=st["end_tolerance_min"])
    has_track = bool(gps and gps.track and gps.track.points)
    has_gps = bool(gps and (gps.start or has_track))

    # ------------------------------------------------------------------ device
    dev = row.device
    if dev is None or dev.vendor is None:
        res.add("NO_GPS", "medium",
                f"No usable GPS for this {'batch' if row.kind == 'cwp' else 'beat'} (sheet says: {dev.raw if dev else '-'})",
                "device")
    elif gps is None or not gps.found:
        res.add("DEVICE_NOT_ON_PORTAL", "high",
                f"GPS {dev.ref} is not found on the {dev.vendor} portal" +
                (" (only the last digits are in the sheet)" if dev.ref_kind == "tail" else ""), "device")
    elif not has_gps:
        res.add("NO_GPS_DATA", "high", f"No GPS data from {gps.device_name or dev.ref} during the duty", "device")
    if gps and gps.found and gps.device_name and row.pwi:
        sec = gps.device_name.split("-")[0].split("/")[0].strip().upper()
        pwi = row.pwi.replace("/", "").replace(" ", "").upper()
        if sec and pwi and sec.replace(" ", "") not in pwi and pwi not in sec.replace(" ", ""):
            res.add("DEVICE_OTHER_SECTION", "medium",
                    f"GPS {dev.ref} is registered on the portal as '{gps.device_name}', not under {row.pwi}", "device")

    # ------------------------------------------------------------------ photos (individual)
    starts = [p for p in photos if p.role == "start"]
    ends = [p for p in photos if p.role == "end"]
    equip = [p for p in photos if p.role == "equipment"]
    if not row.not_applicable:
        if not row.start_photos:
            res.add("NO_START_PHOTO", "high", "Starting photo not uploaded", "photo")
        if row.kind == "cwp" and not start_only:
            if not row.end_photos:
                res.add("NO_END_PHOTO", "medium", "End photo not uploaded", "photo")
            if not row.equipment_photos:
                res.add("NO_EQUIPMENT_PHOTO", "low", "Equipment photo not uploaded", "photo")

    for p in photos:
        tag = {"start": "Start photo", "end": "End photo", "equipment": "Equipment photo"}[p.role]
        for prob in p.problems:
            if prob["code"] == "NO_COORDINATES" and p.role != "start":
                continue
            res.add(prob["code"], prob["severity"], f"{tag}: {prob['message']}", "photo")
        if not p.stamp_time:
            continue
        exp_day = row.duty_date
        if row.kind == "cwp" and (p.role == "end" or p.stamp_time.time() < dt.time(12)):
            exp_day = row.duty_date + dt.timedelta(days=1)
        if p.stamp_time.date() != exp_day:
            res.add("STAMP_WRONG_DATE", "high" if p.ocr_conf >= 0.75 else "check",
                    f"{tag} stamp is dated {p.stamp_time:%d.%m.%Y}, duty is {exp_day:%d.%m.%Y} (old photo?)", "photo")
        if p.whatsapp_time:
            ahead = mins(p.stamp_time, p.whatsapp_time)
            if ahead > st["photo_ahead_tolerance_min"]:
                res.add("STAMP_AFTER_WHATSAPP", "high",
                        f"{tag} stamp says {hm(p.stamp_time)} but the photo reached WhatsApp at "
                        f"{hm(p.whatsapp_time)} - stamp time has been altered", "photo")
            elif -ahead > st["photo_max_delay_min"]:
                res.add("STAMP_LONG_BEFORE_WHATSAPP", "medium",
                        f"{tag} stamp {hm(p.stamp_time)} but sent on WhatsApp only at {hm(p.whatsapp_time)} "
                        f"({-ahead} min later) - old photo or altered stamp?", "photo")
        nf = cfg["night_hours"]
        night = _in_night(p.stamp_time.time(), nf["from"], nf["to"])
        if night and p.sky_bright:
            res.add("DAYLIGHT_IN_NIGHT_PHOTO", "high",
                    f"{tag} shows daylight sky but the stamp says {hm(p.stamp_time)} - stamp time altered?", "photo")
        if has_track and p.lat is not None and not any(x["code"] == "INVALID_COORDINATES" for x in p.problems):
            pos = gps.track.position_at(p.stamp_time)
            if pos:
                d = km_between(p.lat, p.lon, pos.lat, pos.lon)
                if d > st["photo_vs_gps_max_km"]:
                    res.add("PHOTO_FAR_FROM_GPS", "high",
                            f"{tag} location is {d:.1f} km from where the GPS was at {hm(p.stamp_time)}", "photo")

    credible = [p for p in starts if p.stamp_time and _credible(p, res, row)]
    photo_start = min((p.stamp_time for p in credible), default=None)
    credible_end = [p for p in ends if p.stamp_time and _credible(p, res, row)]
    photo_end = max((p.stamp_time for p in credible_end), default=None)

    # ------------------------------------------------------------------ sheet vs GPS / photo
    s_start, s_end = clock.at(row.start_time), clock.at(row.end_time)
    if has_gps:
        if s_start and gps.start and abs(s_start - gps.start) > tol_s:
            res.add("SHEET_START_DIFF", "medium",
                    f"Sheet start {hm(s_start)} but GPS start {hm(gps.start)} ({mins(s_start, gps.start):+d} min)", "sheet")
        if not start_only:
            if s_end and gps.end and abs(s_end - gps.end) > tol_e:
                res.add("SHEET_END_DIFF", "medium",
                        f"Sheet end {hm(s_end)} but GPS end {hm(gps.end)} ({mins(s_end, gps.end):+d} min)", "sheet")
            if row.km_done is not None and gps.km:
                diff = abs(row.km_done - gps.km)
                if diff > 0.5 and diff > gps.km * st["sheet_km_tolerance_percent"] / 100:
                    res.add("SHEET_KM_DIFF", "medium",
                            f"Sheet km {row.km_done:g} but GPS km {gps.km:g}", "sheet")
            if "NOT GENERATED" in row.km_raw.upper() or "NOT GENERATED" in row.speed_raw.upper():
                res.add("SHEET_SAYS_NO_REPORT", "medium",
                        f"Sheet says 'report not generated' but the portal has data "
                        f"({hm(gps.start)}-{hm(gps.end)}, {gps.km or 0:g} km)", "sheet")
            if row.max_speed is not None and gps.max_speed is not None and abs(row.max_speed - gps.max_speed) > 1:
                res.add("SHEET_SPEED_DIFF", "low",
                        f"Sheet max speed {row.max_speed:g} but GPS max speed {gps.max_speed:g} km/h", "sheet")
    if s_start and photo_start and abs(s_start - photo_start) > tol_s:
        res.add("SHEET_VS_PHOTO_START", "medium",
                f"Sheet start {hm(s_start)} but start photo stamped {hm(photo_start)}", "sheet")

    # ------------------------------------------------------------------ actual duty
    gps_start = gps.start if has_gps else None
    if has_track and gps_start:
        moved = gps.track.first_move()
        if moved and moved - gps_start > dt.timedelta(minutes=30):
            res.add("IDLE_AFTER_START", "medium",
                    f"GPS switched on at {hm(gps_start)} but did not move until {hm(moved)}", "duty")
    if photo_start and (not gps_start or photo_start < gps_start):
        res.actual_start, res.actual_start_source = photo_start, "photo"
        if gps_start and gps_start - photo_start > tol_s:
            res.add("GPS_STARTED_LATE", "medium",
                    f"Patroller on track at {hm(photo_start)} (photo) but GPS started only at {hm(gps_start)} "
                    f"- device switched on late", "duty")
    elif gps_start:
        res.actual_start, res.actual_start_source = gps_start, "GPS"
    if r_start and res.actual_start and res.actual_start > r_start + tol_s:
        res.add("LATE_START", "high",
                f"Late start: duty from {hm(r_start)}, actually started {hm(res.actual_start)} "
                f"({res.actual_start_source})", "duty")
    if r_start and not res.actual_start and not row.not_applicable and dev and dev.vendor:
        res.add("NOT_STARTED", "high", "No evidence of duty start (no GPS data, no readable start photo)", "duty")

    if not start_only:
        ends_seen = [t for t in (gps.end if has_gps else None, photo_end) if t]
        res.actual_end = max(ends_seen) if ends_seen else None
        if r_end and res.actual_end and res.actual_end < r_end - tol_e:
            res.add("EARLY_FINISH", "high",
                    f"Early finish: duty till {hm(r_end)}, last evidence {hm(res.actual_end)}", "duty")
        if has_gps and row.km_target and gps.km is not None and gps.km < row.km_target * st["km_min_percent"] / 100:
            res.add("LESS_DISTANCE", "high",
                    f"GPS shows {gps.km:g} km against beat {row.km_target:g} km", "duty")
        if has_track:
            runs = [r for r in gps.track.high_speed_runs(st["max_walking_speed_kmph"])
                    if r[3] >= st["speed_spike_points"]]
            if runs:
                worst = max(runs, key=lambda r: r[2])
                res.add("HIGH_SPEED", "high",
                        f"Moving above {st['max_walking_speed_kmph']} km/h for {len(runs)} stretch(es), "
                        f"top {worst[2]:.0f} km/h at {hm(worst[0])} - vehicle used?", "duty")
            gaps = gps.track.gaps(st["gps_gap_switched_off_min"], r_start, r_end) if r_start and r_end else []
            gaps = [g for g in gaps if g[0] > (res.actual_start or r_start)]
            if gaps:
                g = max(gaps, key=lambda g: g[1] - g[0])
                res.add("GPS_GAP", "medium",
                        f"GPS silent {hm(g[0])}-{hm(g[1])} ({mins(g[1], g[0])} min) during duty", "duty")

    # ------------------------------------------------------------------ portal's own exceptions
    for name in portal_exc:
        res.verdicts.append(_verdict(name, res, row, gps, has_gps, has_track, photo_start, photo_end,
                                     r_start, r_end, tol_s, tol_e, st, start_only))
    for v in res.verdicts:
        if v.verdict == "FALSE":
            res.add("PORTAL_FALSE_EXCEPTION", "low", f"Portal says '{v.exception}' - not correct: {v.reason}", "portal")
    return res


def _in_night(t: dt.time, frm: str, to: str) -> bool:
    a = dt.time.fromisoformat(frm)
    b = dt.time.fromisoformat(to)
    return t >= a or t <= b


def _credible(p: PhotoResult, res: RowResult, row: SheetRow) -> bool:
    """A photo time we can rely on: readable, right day, not flagged as altered."""
    bad = {"STAMP_AFTER_WHATSAPP", "DAYLIGHT_IN_NIGHT_PHOTO", "INVALID_COORDINATES", "PHOTO_FAR_FROM_GPS",
           "STAMP_WRONG_DATE", "DUPLICATE_PHOTO"}
    own = {x["code"] for x in p.problems}
    if own & bad:
        return False
    tag = {"start": "Start photo", "end": "End photo", "equipment": "Equipment photo"}[p.role]
    return not any(f.code in bad and f.message.startswith(tag) for f in res.findings)


def _verdict(name, res, row, gps, has_gps, has_track, photo_start, photo_end,
             r_start, r_end, tol_s, tol_e, st, start_only) -> Verdict:
    n = name.lower()
    if "late" in n or "after 5" in n:
        if photo_start and r_start:
            if photo_start <= r_start + tol_s:
                return Verdict(name, "FALSE", f"start photo shows patroller on track at {hm(photo_start)}"
                               + (f"; GPS was switched on at {hm(gps.start)}" if has_gps else ""))
            return Verdict(name, "CONFIRMED", f"start photo also late ({hm(photo_start)})")
        if has_gps and r_start:
            g = gps.start
            if g <= r_start + tol_s:
                return Verdict(name, "FALSE", f"GPS itself started at {hm(g)}, within the allowed time")
            return Verdict(name, "CONFIRMED", f"GPS started {hm(g)}; no readable start photo to prove otherwise")
        return Verdict(name, "CANNOT VERIFY", "no GPS data and no readable start photo")
    if "early" in n or "before 3" in n:
        if start_only:
            return Verdict(name, "CANNOT VERIFY", "duty still running at the time of the morning check")
        if photo_end and r_end and photo_end >= r_end - tol_e:
            return Verdict(name, "FALSE", f"end photo stamped {hm(photo_end)}")
        if has_gps and r_end:
            if gps.end >= r_end - tol_e:
                return Verdict(name, "FALSE", f"GPS active till {hm(gps.end)}; duty ends {hm(r_end)}")
            return Verdict(name, "CONFIRMED", f"GPS stopped at {hm(gps.end)}; duty ends {hm(r_end)}")
        return Verdict(name, "CANNOT VERIFY", "no GPS end time")
    if "not started" in n or "not done" in n or "unattended" in n:
        if has_gps:
            return Verdict(name, "FALSE", f"GPS data exists ({hm(gps.start)}-{hm(gps.end)}, {gps.km or 0:g} km)")
        if photo_start:
            return Verdict(name, "FALSE", f"start photo at {hm(photo_start)} shows duty was done; GPS not working / not carried")
        return Verdict(name, "CONFIRMED", "no GPS data and no readable start photo")
    if "switched off" in n or "off device" in n:
        if not has_track:
            return Verdict(name, "CANNOT VERIFY", "no GPS points available")
        gaps = gps.track.gaps(st["gps_gap_switched_off_min"], r_start, r_end) if r_start and r_end else []
        if gaps:
            g = max(gaps, key=lambda g: g[1] - g[0])
            return Verdict(name, "CONFIRMED", f"no GPS signal {hm(g[0])}-{hm(g[1])}")
        return Verdict(name, "FALSE", "GPS points are continuous through the duty")
    if "distance" in n or "beat not covered" in n or "4 km" in n:
        if start_only:
            return Verdict(name, "CANNOT VERIFY", "duty still running at the time of the morning check")
        if not has_gps or gps.km is None:
            return Verdict(name, "CANNOT VERIFY", "no GPS distance")
        need = (row.km_target or 0) * st["km_min_percent"] / 100
        if "4 km" in n:
            need = 4
        if not need:
            return Verdict(name, "CANNOT VERIFY", "beat length not in the sheet")
        if gps.km >= need:
            return Verdict(name, "FALSE", f"GPS shows {gps.km:g} km (needed {need:g})")
        return Verdict(name, "CONFIRMED", f"GPS shows {gps.km:g} km (needed {need:g})")
    if "speed" in n:
        if not has_track:
            return Verdict(name, "CANNOT VERIFY", "no GPS points available")
        runs = gps.track.high_speed_runs(st["max_walking_speed_kmph"])
        real = [r for r in runs if r[3] >= st["speed_spike_points"]]
        if real:
            w = max(real, key=lambda r: r[2])
            return Verdict(name, "CONFIRMED", f"{len(real)} sustained stretch(es), top {w[2]:.0f} km/h at {hm(w[0])}")
        if runs:
            return Verdict(name, "FALSE", f"only {len(runs)} single-point jump(s) - GPS jitter, not real movement")
        return Verdict(name, "FALSE", f"no point above {st['max_walking_speed_kmph']} km/h")
    if "2 hour" in n:
        if has_gps and gps.start and gps.end:
            h = (gps.end - gps.start).total_seconds() / 3600
            return Verdict(name, "CONFIRMED" if h < 2 else "FALSE", f"GPS active {h:.1f} h")
        return Verdict(name, "CANNOT VERIFY", "no GPS times")
    return Verdict(name, "CANNOT VERIFY", "no rule for this exception type")


# ============================================================================= cross-row checks

def cross_row(results: list[RowResult], history_hashes: list[dict], st: dict, same_place_fn, raw_fn):
    """Photos reused between batches/days, and different batches photographed at the same spot."""
    from .photos import hamming
    seen: list[tuple[str, RowResult, PhotoResult]] = []
    for r in results:
        for p in r.photos:
            if not p.dhash:
                continue
            key = f"{r.row.pwi} {r.row.beat[:30]}"
            for k2, r2, p2 in seen:
                if r2 is r or p2.drive_id == p.drive_id:
                    continue
                if hamming(p.dhash, p2.dhash) <= 12:
                    r.add("DUPLICATE_PHOTO", "high",
                          f"{p.role.title()} photo is the same picture as {k2}'s {p2.role} photo", "photo")
            for h in history_hashes:
                if h.get("drive_id") != p.drive_id and hamming(p.dhash, h.get("dhash", "")) <= 12 \
                        and h.get("date") != str(r.row.duty_date):
                    r.add("DUPLICATE_PHOTO", "high",
                          f"{p.role.title()} photo was already used on {h.get('date')} ({h.get('batch', '')})", "photo")
                    break
            seen.append((key, r, p))
    # same spot, same angle, different batch, stamps within 15 minutes (same section only, to keep it fast)
    by_sec: dict[str, list[tuple[RowResult, PhotoResult]]] = {}
    for r in results:
        for p in r.photos:
            if p.role == "start" and p.stamp_time:
                by_sec.setdefault(r.row.pwi, []).append((r, p))
    for items in by_sec.values():
        for i, (ra, pa) in enumerate(items):
            for rb, pb in items[i + 1:]:
                if ra is rb or abs((pa.stamp_time - pb.stamp_time).total_seconds()) > 900:
                    continue
                score = same_place_fn(raw_fn(pa.drive_id), raw_fn(pb.drive_id))
                if score >= 60:
                    for x, y, py in ((ra, rb, pb), (rb, ra, pa)):
                        x.add("SAME_SPOT_OTHER_BATCH", "medium",
                              f"Start photo taken at the same spot, same angle as '{y.row.beat[:40]}' "
                              f"({hm(py.stamp_time)}) - one person photographing both?", "photo")
