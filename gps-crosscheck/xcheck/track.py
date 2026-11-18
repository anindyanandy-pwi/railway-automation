"""Turn raw GPS points (from either portal) into duty facts: start, end, km, speeds, gaps."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from .photos import km_between


@dataclass
class Point:
    t: dt.datetime          # local time (IST)
    lat: float
    lon: float
    speed: float            # km/h as reported by the device


@dataclass
class Track:
    points: list[Point] = field(default_factory=list)

    def window(self, start: dt.datetime, end: dt.datetime) -> "Track":
        return Track([p for p in self.points if start <= p.t <= end])

    # ------------------------------------------------------------------ facts
    @property
    def first(self) -> dt.datetime | None:
        return self.points[0].t if self.points else None

    @property
    def last(self) -> dt.datetime | None:
        return self.points[-1].t if self.points else None

    def first_move(self, metres: float = 60) -> dt.datetime | None:
        """First time the device is more than `metres` from where it was switched on."""
        if not self.points:
            return None
        p0 = self.points[0]
        for p in self.points:
            if km_between(p0.lat, p0.lon, p.lat, p.lon) * 1000 > metres:
                return p.t
        return None

    def km(self, max_kmph: float = 25) -> float:
        """Distance walked. Jumps faster than max_kmph between two points are GPS errors and skipped."""
        total = 0.0
        for a, b in zip(self.points, self.points[1:]):
            d = km_between(a.lat, a.lon, b.lat, b.lon)
            hours = max((b.t - a.t).total_seconds() / 3600, 1 / 3600)
            if d / hours <= max_kmph:
                total += d
        return round(total, 2)

    def max_speed(self) -> float | None:
        return max((p.speed for p in self.points), default=None)

    def high_speed_runs(self, limit: float) -> list[tuple[dt.datetime, dt.datetime, float, int]]:
        """Stretches of consecutive points above `limit` km/h: (from, to, top speed, points)."""
        runs, cur = [], []
        for p in self.points:
            if p.speed > limit:
                cur.append(p)
            elif cur:
                runs.append(cur)
                cur = []
        if cur:
            runs.append(cur)
        return [(r[0].t, r[-1].t, max(x.speed for x in r), len(r)) for r in runs]

    def gaps(self, minutes: float, start: dt.datetime | None = None,
             end: dt.datetime | None = None) -> list[tuple[dt.datetime, dt.datetime]]:
        """Periods with no GPS point for longer than `minutes` (device off / no signal)."""
        out = []
        ts = [p.t for p in self.points]
        if start and (not ts or ts[0] - start > dt.timedelta(minutes=minutes)):
            out.append((start, ts[0] if ts else end))
        for a, b in zip(ts, ts[1:]):
            if b - a > dt.timedelta(minutes=minutes):
                out.append((a, b))
        if end and ts and end - ts[-1] > dt.timedelta(minutes=minutes):
            out.append((ts[-1], end))
        return out

    def position_at(self, t: dt.datetime, within_min: float = 10) -> Point | None:
        best = min(self.points, key=lambda p: abs((p.t - t).total_seconds()), default=None)
        if best and abs((best.t - t).total_seconds()) <= within_min * 60:
            return best
        return None


@dataclass
class GpsSummary:
    """What a portal says about one device for one duty window."""
    vendor: str
    device_ref: str
    device_name: str = ""
    found: bool = False                 # device known on the portal
    start: dt.datetime | None = None
    end: dt.datetime | None = None
    km: float | None = None
    avg_speed: float | None = None
    max_speed: float | None = None
    track: Track | None = None
    source: str = ""                    # "report" | "history" | "report+history"
    portal_exceptions: list[str] = field(default_factory=list)   # names as the portal gives them
    error: str = ""
