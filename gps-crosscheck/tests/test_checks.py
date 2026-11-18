"""Rule tests with real cases from 17/18 Nov 2026. Run: python -m pytest tests -q"""
import datetime as dt

import yaml

from xcheck import checks
from xcheck.photos import PhotoResult
from xcheck.sheet import Device, SheetRow
from xcheck.track import GpsSummary, Point, Track

CFG = yaml.safe_load(open("config/howrah.yaml"))
ST = CFG["defaults"]
D = dt.datetime


def cwp_row(**kw):
    base = dict(kind="cwp", duty_date=dt.date(2026, 11, 17), row_no=9, pwi="BMG", beat="3 CWP BATCH BMG/P-3",
                device=Device("rilapp", "120084", "120084"), roster="23:00-07:00",
                roster_start=dt.time(23), roster_end=dt.time(7), start_time=dt.time(23, 15),
                end_time=dt.time(6, 56), km_target=8, km_done=8.4, km_raw="8.4",
                start_photos=["a"], end_photos=["b"], equipment_photos=["c"])
    base.update(kw)
    return SheetRow(**base)


def walk(start, end, step_min=1, speed=2.5, lat=22.60, lon=88.33, jump_at=None):
    pts, t, i = [], start, 0
    while t <= end:
        s = speed
        if jump_at and jump_at[0] <= t <= jump_at[1]:
            s = 18
        pts.append(Point(t, lat + i * 0.0003, lon, s))
        t += dt.timedelta(minutes=step_min)
        i += 1
    return Track(pts)


def photo(role, when, wa=None, **kw):
    p = PhotoResult(drive_id=role, role=role, ok_download=True, stamp_time=when, whatsapp_time=wa, ocr_conf=0.9)
    for k, v in kw.items():
        setattr(p, k, v)
    return p


def codes(res):
    return {f.code for f in res.findings}


def test_bmg_p3_sheet_end_does_not_match_gps():
    # Real case: sheet says end 06:56, rilapp says GPS stopped 06:21
    trk = walk(D(2026, 11, 17, 23, 5), D(2026, 11, 18, 6, 21))
    gps = GpsSummary("rilapp", "120084", "BMG-PM-Batch-5-2/0-4/0", True,
                     D(2026, 11, 17, 23, 5, 56), D(2026, 11, 18, 6, 21, 18), 8.7, 2.7, 2.5, trk, "report+history")
    res = checks.evaluate(cwp_row(), gps,
                          [photo("start", D(2026, 11, 17, 23, 5)), photo("end", D(2026, 11, 18, 6, 56, 13))],
                          "cwp", ST, CFG, ["Off Device"])
    assert "SHEET_END_DIFF" in codes(res)
    assert "EARLY_FINISH" not in codes(res)          # end photo 06:56 is within 15 min of 07:00
    assert res.actual_end == D(2026, 11, 18, 6, 56, 13)


def test_photo_proves_patroller_on_time_portal_late_start_false():
    trk = walk(D(2026, 11, 17, 23, 40), D(2026, 11, 18, 6, 58))
    gps = GpsSummary("nbtrack", "5754130017736", "HWH/P-1", True, trk.first, trk.last, trk.km(), 2.5, 2.5, trk, "history")
    res = checks.evaluate(cwp_row(device=Device("nbtrack", "5754130017736", "5754130017736", ref_kind="sim"),
                                  start_time=dt.time(23, 2)), gps,
                          [photo("start", D(2026, 11, 17, 22, 50), D(2026, 11, 17, 23, 7, 55)),
                           photo("end", D(2026, 11, 18, 7, 0))],
                          "cwp", ST, CFG, ["Late Start"])
    assert res.verdicts[0].verdict == "FALSE"
    assert "GPS_STARTED_LATE" in codes(res)
    assert "LATE_START" not in codes(res)


def test_altered_stamp_is_not_trusted_and_late_start_confirmed():
    # stamp 05:14 but WhatsApp received it 05:04 -> altered; GPS started 05:40
    trk = walk(D(2026, 11, 17, 5, 40), D(2026, 11, 17, 13, 50))
    gps = GpsSummary("rilapp", "120162", "HWH-KM-K-08", True, trk.first, trk.last, trk.km(), 2.5, 2.5, trk, "history")
    row = SheetRow(kind="keyman", duty_date=dt.date(2026, 11, 17), row_no=10, pwi="HWH", beat="PF-20 TO PF-23",
                   device=Device("rilapp", "120162", "120162"), roster="05:00-14:00",
                   roster_start=dt.time(5), roster_end=dt.time(14), start_time=dt.time(5, 14), start_photos=["a"])
    res = checks.evaluate(row, gps, [photo("start", D(2026, 11, 17, 5, 14), D(2026, 11, 17, 5, 4))],
                          "keyman_full", ST, CFG, ["On device after 5 AM (late start)"])
    assert "STAMP_AFTER_WHATSAPP" in codes(res)
    assert "LATE_START" in codes(res)
    assert res.verdicts[0].verdict == "CONFIRMED"


def test_speed_spike_is_jitter_sustained_is_vehicle():
    trk = walk(D(2026, 11, 17, 23, 0), D(2026, 11, 18, 7, 0))
    trk.points[100].speed = 22                     # one-point jump
    gps = GpsSummary("nbtrack", "x", "n", True, trk.first, trk.last, trk.km(), 2.5, 22, trk, "history")
    res = checks.evaluate(cwp_row(), gps, [], "cwp", ST, CFG, ["High Speed > 5 km/h"])
    assert res.verdicts[0].verdict == "FALSE"
    trk2 = walk(D(2026, 11, 17, 23, 0), D(2026, 11, 18, 7, 0),
                jump_at=(D(2026, 11, 18, 1, 0), D(2026, 11, 18, 1, 20)))
    gps2 = GpsSummary("nbtrack", "x", "n", True, trk2.first, trk2.last, trk2.km(), 2.5, 18, trk2, "history")
    res2 = checks.evaluate(cwp_row(), gps2, [], "cwp", ST, CFG, ["High Speed > 5 km/h"])
    assert res2.verdicts[0].verdict == "CONFIRMED"
    assert "HIGH_SPEED" in codes(res2)


def test_no_gps_but_photo_means_not_started_is_false():
    gps = GpsSummary("rilapp", "120341", "HWH-KM-K-02", True)
    row = SheetRow(kind="keyman", duty_date=dt.date(2026, 11, 18), row_no=5, pwi="HWH", beat="PF-12 TO PF-15",
                   device=Device("rilapp", "120341", "120341"), roster="05:00-14:00",
                   roster_start=dt.time(5), roster_end=dt.time(14), start_time=dt.time(5, 10), start_photos=["a"])
    res = checks.evaluate(row, gps, [photo("start", D(2026, 11, 18, 5, 8))], "keyman_start", ST, CFG,
                          ["Patrolling Not Done"])
    assert res.verdicts[0].verdict == "FALSE"
    assert "NO_GPS_DATA" in codes(res)


def test_wrong_photo_date_flagged():
    res = checks.evaluate(cwp_row(), None, [photo("start", D(2026, 11, 24, 23, 5))], "cwp", ST, CFG, [])
    assert "STAMP_WRONG_DATE" in codes(res)
