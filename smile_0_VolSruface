"""
fx_vol_surface.py

Builds an FX implied-vol surface from market quotes per maturity:
    ATM vol, 25-delta RR, 25-delta BF, 10-delta RR, 10-delta BF

Quote conventions (all configurable / documented)
-------------------------------------------------
RR  = vol(call) - vol(put)                      (call = call on the base ccy, e.g. USD)
BF  = 0.5*(vol(call) + vol(put)) - vol(ATM)     ("smile" / simple butterfly)
=> vol_25C = ATM + BF25 + RR25/2      vol_25P = ATM + BF25 - RR25/2
   vol_10C = ATM + BF10 + RR10/2      vol_10P = ATM + BF10 - RR10/2

delta_type     : "spot" (default; delta = DF_foreign * N(d1)) or "forward" (N(d1)).
                 premium_adjusted=True (default False) multiplies that delta by
                 K/F -- see strike_from_delta() for why that needs a numerical
                 solve rather than a closed form.
atm_convention : "forward" (K = F, default) or "dns" (delta-neutral straddle,
                 K = F*exp(-0.5*vol_ATM^2*Te) -- the d2=0 point).

Construction
------------
1. For each quoted maturity: 5 pillar vols (10P, 25P, ATM, 25C, 10C) -> 5 strikes
   (each pillar's strike is solved from its own vol; closed form, no iteration).
2. Smile in strike: interpolator through the pillars' (ln(K/F), vol) points --
   smile_method="cubic" (natural spline, default), "pchip" (shape-preserving,
   no overshoot), or "linear" (straight lines pillar-to-pillar; simplest and
   most transparent, but the kink in slope at every single pillar -- not just
   at the wings -- makes it the most prone of the three to the butterfly
   arbitrage check flagging negative density; always check_arbitrage() when
   using it). Outside the outermost pillars, wing_extrapolation="linear"
   (default) continues with the boundary slope (never lowering vol going
   outward; for smile_method="linear" this is seamless -- same slope inside
   and outside the join, since the interior is already straight lines);
   "flat" holds the boundary vol instead, which leaves a kink that
   check_arbitrage() reports as negative density.
3. Across maturities: each delta-pillar vol is interpolated LINEARLY IN TOTAL
   VARIANCE (vol^2 * Te) in expiry, flat-vol extrapolation beyond the first/last
   tenor. Zero rates (rd, rf) come from separate RateCurve objects (own tenor
   grids, independent of the vol quote tenors), linearly interpolated in rate
   vs. tenor and evaluated at Td (the discounting/forward horizon, not Te),
   flat-extrapolated beyond their own first/last tenor. The smile for the
   target expiry is then rebuilt from the interpolated pillars with that
   expiry's forward.
4. check_arbitrage(): butterfly (density >= 0) per tenor and calendar
   (total variance non-decreasing across tenors) checks.

Time: ACT/365. Te = valuation -> expiry (vol time); Td = spot date -> delivery
(forward and foreign discounting), matching how the pricer uses te / td.
Delivery defaults to expiry + 2 business days (weekends, plus any dates passed
in `holidays`) unless given per tenor.

Requires: numpy, scipy, matplotlib (plot only).
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date, timedelta
from statistics import NormalDist
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
from scipy.interpolate import CubicSpline, PchipInterpolator
from scipy.optimize import newton, brentq

_N = NormalDist()

PILLAR_NAMES = ["10P", "25P", "ATM", "25C", "10C"]
PILLAR_DELTAS = [-0.10, -0.25, 0.0, 0.25, 0.10]   # signed; ATM handled by convention

# 5-delta-step extended grid: 5..45-delta put, ATM, 45..5-delta call, in
# ascending-strike order. See FxVolSurface(pillar_step=5) / _expand_pillars().
EXT5_PUT_NAMES = ["05P", "10P", "15P", "20P", "25P", "30P", "35P", "40P", "45P"]
EXT5_CALL_NAMES = ["45C", "40C", "35C", "30C", "25C", "20C", "15C", "10C", "05C"]
EXT5_PILLAR_NAMES = EXT5_PUT_NAMES + ["ATM"] + EXT5_CALL_NAMES
EXT5_PILLAR_DELTAS = ([-d / 100 for d in (5, 10, 15, 20, 25, 30, 35, 40, 45)] + [0.0]
                     + [d / 100 for d in (45, 40, 35, 30, 25, 20, 15, 10, 5)])


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def yearfrac(d0: date, d1: date) -> float:
    """ACT/365."""
    return (d1 - d0).days / 365.0


def add_business_days(d: date, n: int, holidays: Optional[Sequence[date]] = None) -> date:
    """Shift `d` by `n` business days, skipping weekends and any date in
    `holidays` (a list/set of public holiday dates; None or empty = weekends
    only, the old behaviour)."""
    holiday_set = set(holidays) if holidays else set()
    step = 1 if n >= 0 else -1
    moved = 0
    while moved < abs(n):
        d += timedelta(days=step)
        if d.weekday() < 5 and d not in holiday_set:
            moved += 1
    return d


def strike_from_delta(delta: float, vol: float, fwd: float, te: float,
                      df_foreign: float, delta_type: str = "spot",
                      premium_adjusted: bool = False) -> float:
    """Strike whose delta equals `delta` under Black with `vol`.

    Unadjusted (premium_adjusted=False, default) -- closed form:
        call (delta>0): delta = disc * N(d1)
        put  (delta<0): delta = -disc * N(-d1)
        Invert N to get d1 directly, then K = F*exp(-d1*vol*sqrt(Te) + 0.5*vol^2*Te).

    Premium-adjusted (premium_adjusted=True) -- delta itself carries a K/F factor,
    and uses d2 (not d1):
        call (delta>0): delta = disc * N(d2) * K/F
        put  (delta<0): delta = -disc * N(-d2) * K/F
    K now appears twice -- inside d2(K), and again as the multiplicative K/F
    outside it -- so there's no closed-form inverse. As a function of K this
    is a hump that is 0 at both K->0 and K->infinity, so for a given delta
    magnitude there are generally TWO strikes solving the equation, not one.
    Market convention picks the strike near the unadjusted-delta strike, so we
    solve numerically (Newton's method) starting from that unadjusted strike as
    the initial guess, which converges to the nearby (correct) root rather than
    the far one.

    disc = df_foreign for delta_type="spot", 1.0 for delta_type="forward".
    d1(K) = (ln(F/K) + 0.5*vol^2*Te) / (vol*sqrt(Te)),  d2 = d1 - vol*sqrt(Te)
    """
    disc = df_foreign if delta_type == "spot" else 1.0
    sqrt_te = math.sqrt(te)

    if delta > 0:
        d1_target = _N.inv_cdf(delta / disc)
    else:
        d1_target = -_N.inv_cdf(-delta / disc)
    k_unadjusted = fwd * math.exp(-d1_target * vol * sqrt_te + 0.5 * vol * vol * te)

    if not premium_adjusted:
        return k_unadjusted

    def f(k: float) -> float:
        d1 = (math.log(fwd / k) + 0.5 * vol * vol * te) / (vol * sqrt_te)
        d2 = d1 - vol * sqrt_te
        pa_delta = disc * _N.cdf(d2) * k / fwd if delta > 0 else -disc * _N.cdf(-d2) * k / fwd
        return pa_delta - delta

    root = newton(f, x0=k_unadjusted, tol=1e-12, maxiter=100)
    if abs(f(root)) > 1e-8:
        raise ValueError(f"strike_from_delta(premium_adjusted=True) failed to converge for delta={delta}")
    return root


def black_delta(k: float, vol: float, fwd: float, te: float, disc: float,
                is_call: bool, premium_adjusted: bool = False) -> float:
    """Signed Black delta (call > 0, put < 0) of a single option struck at
    `k`, using this module's delta convention: `disc` should already be
    df_foreign for spot delta or 1.0 for forward delta (see strike_from_delta)."""
    d1 = (math.log(fwd / k) + 0.5 * vol * vol * te) / (vol * math.sqrt(te))
    if premium_adjusted:
        d = d1 - vol * math.sqrt(te)
        return disc * _N.cdf(d) * k / fwd if is_call else -disc * _N.cdf(-d) * k / fwd
    return disc * _N.cdf(d1) if is_call else -disc * _N.cdf(-d1)


def solve_pillar_on_line(target_delta: float, k1: float, vol1: float, k2: float, vol2: float,
                         fwd: float, te: float, disc: float, is_call: bool,
                         premium_adjusted: bool = False) -> Tuple[float, float]:
    """Find (K, vol) on the line K(vol) = k1 + (k2-k1)/(vol2-vol1)*(vol-vol1) --
    i.e. strike assumed linear in vol between two known anchor pillars
    (k1, vol1) and (k2, vol2) -- whose Black delta equals `target_delta`.

    If target_delta lies between the anchors' own deltas, this is a bracketed
    interior solve (Brent's method). If it lies outside both (extrapolation
    past whichever anchor is closer to it, e.g. solving for 5-delta from the
    10/25-delta line), there's no bracket, so Newton's method is used instead,
    started from the nearer anchor's vol.

    Returns (K, vol) for the new pillar.
    """
    if vol1 == vol2:
        raise ValueError("solve_pillar_on_line: anchor vols are equal -- strike cannot be "
                         "expressed as a function of vol between these two pillars")
    slope = (k2 - k1) / (vol2 - vol1)

    def k_of(vol: float) -> float:
        return k1 + slope * (vol - vol1)

    def f(vol: float) -> float:
        return black_delta(k_of(vol), vol, fwd, te, disc, is_call, premium_adjusted) - target_delta

    f1, f2 = f(vol1), f(vol2)
    if f1 == 0:
        vol = vol1
    elif f2 == 0:
        vol = vol2
    elif f1 * f2 < 0:
        vol = brentq(f, vol1, vol2, xtol=1e-14, maxiter=200)
    else:
        x0 = vol1 if abs(f1) < abs(f2) else vol2
        vol = newton(f, x0=x0, tol=1e-14, maxiter=100)
        if abs(f(vol)) > 1e-8:
            raise ValueError(f"solve_pillar_on_line: failed to converge for target_delta={target_delta}")
    return k_of(vol), vol


def black_call(fwd: float, k: np.ndarray, vol: np.ndarray, te: float, df_dom: float) -> np.ndarray:
    """Undiscounted-forward Black call, discounted by df_dom (vectorised)."""
    sd = vol * math.sqrt(te)
    d1 = (np.log(fwd / k) + 0.5 * sd * sd) / sd
    d2 = d1 - sd
    cdf = np.vectorize(_N.cdf)
    return df_dom * (fwd * cdf(d1) - k * cdf(d2))


# ---------------------------------------------------------------------------
# Market quote container
# ---------------------------------------------------------------------------

class RateCurve:
    """Zero rate curve: (tenor in years, continuously-compounded decimal rate)
    pairs, linearly interpolated in rate vs. tenor. Flat extrapolation beyond
    the first/last tenor (numpy's np.interp default -- the boundary rate is
    held constant past either end, not extended by the local slope)."""

    def __init__(self, points: Sequence[Tuple[float, float]]):
        pts = sorted(points, key=lambda p: p[0])
        t = np.array([p[0] for p in pts], dtype=float)
        r = np.array([p[1] for p in pts], dtype=float)
        if np.any(np.diff(t) <= 0):
            raise ValueError("RateCurve tenors must be strictly increasing (duplicate tenor found)")
        self.tenors, self.rates = t, r

    def rate(self, t: float) -> float:
        return float(np.interp(t, self.tenors, self.rates))


@dataclass
class TenorQuote:
    expiry: date
    atm: float
    rr25: float
    bf25: float
    rr10: float
    bf10: float
    delivery: Optional[date] = None  # default: expiry + 2 business days

    def pillar_vols(self, vol_in_percent: bool = True) -> np.ndarray:
        """[10P, 25P, ATM, 25C, 10C] vols as decimals."""
        s = 0.01 if vol_in_percent else 1.0
        atm, rr25, bf25, rr10, bf10 = (x * s for x in (self.atm, self.rr25, self.bf25, self.rr10, self.bf10))
        return np.array([
            atm + bf10 - 0.5 * rr10,   # 10P
            atm + bf25 - 0.5 * rr25,   # 25P
            atm,                       # ATM
            atm + bf25 + 0.5 * rr25,   # 25C
            atm + bf10 + 0.5 * rr10,   # 10C
        ])


# ---------------------------------------------------------------------------
# One-expiry smile
# ---------------------------------------------------------------------------

class Smile:
    def __init__(self, te: float, td: float, fwd: float, df_dom: float, df_for: float,
                 strikes: np.ndarray, vols: np.ndarray, method: str = "cubic",
                 wing_extrapolation: str = "linear", names: Optional[Sequence[str]] = None):
        if np.any(np.diff(strikes) <= 0):
            raise ValueError(
                f"Pillar strikes are not strictly increasing ({np.round(strikes, 4).tolist()}); "
                "check the RR/BF quotes (and, if using pillar_step, the extended-pillar solve) "
                "for this expiry."
            )
        self.te, self.td, self.fwd = te, td, fwd
        self.df_dom, self.df_for = df_dom, df_for
        self.strikes, self.vols = strikes, vols
        self.names = list(names) if names is not None else list(PILLAR_NAMES)
        self.atm_index = self.names.index("ATM")
        self.atm_vol = float(vols[self.atm_index])
        self.k = np.log(strikes / fwd)
        if method == "cubic":
            self._f = CubicSpline(self.k, vols, bc_type="natural")
        elif method == "pchip":
            self._f = PchipInterpolator(self.k, vols)
        elif method == "linear":
            self._f = lambda x: np.interp(x, self.k, vols)
        else:
            raise ValueError("smile_method must be 'cubic', 'pchip', or 'linear'")

        if wing_extrapolation not in ("linear", "flat"):
            raise ValueError("wing_extrapolation must be 'linear' or 'flat'")
        # Outside 10P..10C: continue with the boundary slope (C1, no kink) but never let
        # vol fall going outward; otherwise flat. A flat wing on a sloped boundary creates
        # a concave kink = negative density.
        if wing_extrapolation == "linear":
            if method == "linear":
                # Already piecewise-linear inside, so re-using the last/first
                # segment's own slope continues the exact same straight line
                # with no kink at the join at all (not just a value match).
                hi_slope = (vols[-1] - vols[-2]) / (self.k[-1] - self.k[-2])
                lo_slope = (vols[1] - vols[0]) / (self.k[1] - self.k[0])
            else:
                hi_slope = float(self._f(self.k[-1], 1))
                lo_slope = float(self._f(self.k[0], 1))
            self._slope_hi = max(hi_slope, 0.0)
            self._slope_lo = min(lo_slope, 0.0)
        else:
            self._slope_hi = self._slope_lo = 0.0

    def vol(self, strike) -> np.ndarray:
        """Vol at strike(s); spline inside the pillar range, wing extrapolation outside."""
        k = np.log(np.asarray(strike, dtype=float) / self.fwd)
        v = self._f(np.clip(k, self.k[0], self.k[-1]))
        return (v + self._slope_hi * np.maximum(k - self.k[-1], 0.0)
                  + self._slope_lo * np.minimum(k - self.k[0], 0.0))


# ---------------------------------------------------------------------------
# Surface
# ---------------------------------------------------------------------------

class FxVolSurface:
    def __init__(self, val_date: date, spot_date: date, spot: float,
                 quotes: Sequence[TenorQuote], domestic_curve: RateCurve, foreign_curve: RateCurve,
                 vol_in_percent: bool = True,
                 delta_type: str = "spot", atm_convention: str = "forward",
                 smile_method: str = "cubic", wing_extrapolation: str = "linear",
                 premium_adjusted: bool = True, holidays: Optional[Sequence[date]] = None,
                 pillar_step: Optional[int] = None):
        if delta_type not in ("spot", "forward"):
            raise ValueError("delta_type must be 'spot' or 'forward'")
        if atm_convention not in ("forward", "dns"):
            raise ValueError("atm_convention must be 'forward' or 'dns'")
        if pillar_step not in (None, 5):
            raise ValueError("pillar_step must be None (5-pillar smile) or 5 (5..45-delta extended smile)")

        self.val_date, self.spot_date, self.spot = val_date, spot_date, spot
        self.delta_type, self.atm_convention, self.smile_method = delta_type, atm_convention, smile_method
        self.wing_extrapolation = wing_extrapolation
        self.premium_adjusted = premium_adjusted
        self.domestic_curve, self.foreign_curve = domestic_curve, foreign_curve
        self.holidays = set(holidays) if holidays else set()
        self.pillar_step = pillar_step

        self.quotes: List[TenorQuote] = sorted(quotes, key=lambda q: q.expiry)
        self.T = np.array([yearfrac(val_date, q.expiry) for q in self.quotes])
        if np.any(self.T <= 0) or np.any(np.diff(self.T) <= 0):
            raise ValueError("Quote expiries must be after the valuation date and strictly increasing")

        self.pillars = np.array([q.pillar_vols(vol_in_percent) for q in self.quotes])  # (n, 5)
        self._delivery_by_expiry: Dict[date, date] = {
            q.expiry: (q.delivery or add_business_days(q.expiry, 2, self.holidays)) for q in self.quotes
        }
        self._cache: Dict[Tuple[date, date], Smile] = {}

    # -- time interpolation of the delta pillars (linear in total variance) --
    def pillar_vols_at(self, te: float) -> np.ndarray:
        if te <= self.T[0]:
            return self.pillars[0].copy()
        if te >= self.T[-1]:
            return self.pillars[-1].copy()
        w = self.pillars ** 2 * self.T[:, None]                       # total variance
        w_te = np.array([np.interp(te, self.T, w[:, j]) for j in range(w.shape[1])])
        return np.sqrt(w_te / te)

    # -- build (and cache) the smile for any expiry --
    def build_smile(self, expiry: date, delivery: Optional[date] = None) -> Smile:
        delivery = delivery or self._delivery_by_expiry.get(expiry) or add_business_days(expiry, 2, self.holidays)
        key = (expiry, delivery)
        if key in self._cache:
            return self._cache[key]

        te = yearfrac(self.val_date, expiry)
        td = yearfrac(self.spot_date, delivery)
        if te <= 0:
            raise ValueError("expiry must be after the valuation date")

        td_from_valdate = yearfrac(self.val_date, delivery)

        t_settle = yearfrac(self.val_date, self.spot_date)
        rd_settle = self.domestic_curve.rate(t_settle)
        rf_settle = self.foreign_curve.rate(t_settle)
        df_settle_dom = math.exp(-rd_settle * t_settle)
        df_settle_for = math.exp(-rf_settle * t_settle)

        rd_delivery = self.domestic_curve.rate(td_from_valdate)
        rf_delivery = self.foreign_curve.rate(td_from_valdate)
        df_delivery_dom = math.exp(-rd_delivery * td_from_valdate)
        df_delivery_for = math.exp(-rf_delivery * td_from_valdate)

        df_spot_dom = df_delivery_dom / df_settle_dom
        df_spot_for = df_delivery_for / df_settle_for

        rd = -math.log(df_spot_dom) / yearfrac(self.spot_date, delivery)
        rf = -math.log(df_spot_for) / yearfrac(self.spot_date, delivery)

        # rd = self.domestic_curve.rate(td)
        # rf = self.foreign_curve.rate(td)
        fwd = self.spot * math.exp((rd - rf) * td)
        df_dom, df_for = math.exp(-rd * td), math.exp(-rf * td)

        vols = self.pillar_vols_at(te)
        strikes = np.empty(5)
        for i, (name, delta) in enumerate(zip(PILLAR_NAMES, PILLAR_DELTAS)):
            if name == "ATM":
                strikes[i] = fwd if self.atm_convention == "forward" else fwd * math.exp(-0.5 * vols[i] ** 2 * te)
            else:
                strikes[i] = strike_from_delta(delta, vols[i], fwd, te, df_for, self.delta_type,
                                              self.premium_adjusted)

        names = PILLAR_NAMES
        if self.pillar_step == 5:
            strikes, vols, names = self._expand_pillars(strikes, vols, fwd, te, df_for)

        smile = Smile(te, td, fwd, df_dom, df_for, strikes, vols, self.smile_method,
                     self.wing_extrapolation, names)
        self._cache[key] = smile
        return smile

    def _expand_pillars(self, base_strikes: np.ndarray, base_vols: np.ndarray, fwd: float, te: float,
                        df_for: float) -> Tuple[np.ndarray, np.ndarray, List[str]]:
        """Build the 19-point 5-delta-step grid from the 5 known pillars
        (base_strikes/base_vols, in PILLAR_NAMES order: 10P, 25P, ATM, 25C, 10C).

        For each new pillar, assume strike is linear in vol between the two
        nearest known anchors (10P&25P, 25P&ATM, ATM&25C, or 25C&10C), then
        solve for the point on that line whose Black delta matches the target.
        05P and 05C are the exception: rather than extrapolating that line
        past 10P/10C, they use FLAT vol extrapolation -- vol_05P = vol_10P and
        vol_05C = vol_10C -- with the strike then solved the same way the
        original 5 pillars are (strike_from_delta at that flat vol).

        Fallback for 30P..45P and 30C..45C: the ATM anchor's own delta
        magnitude shrinks as tenor and rate-differential grow (it is not
        always ~50%), so a target like 45-delta can exceed what the 25-delta/
        ATM line can reach without solving to a strike PAST the ATM strike --
        nonsensical for something labelled a 45-delta put. Whenever a target
        here isn't strictly bracketed by its anchors' own deltas, it is instead
        filled by straight-line interpolation of (K, vol) against nominal
        delta-grid position (25, 30, 35, 40, 45, with ATM standing in at the
        conventional 50 position), using the nearest pillars on each side that
        WERE delta-solved -- 25P/25C and ATM always qualify, so a fill is
        always possible. This trades an exact delta at that one pillar for a
        continuous, gap-free smile; it never affects 05..20-delta or ATM.

        KNOWN LIMITATION: on tenors where multiple consecutive pillars need
        filling at once (seen on this module's own 5Y sample tenor, premium
        -adjusted delta), the straight line can meet the genuinely delta-solved
        segment at a visibly different slope, since it has no knowledge of the
        smile's actual curvature there. That slope kink can be sharp enough to
        tip check_arbitrage()'s butterfly check slightly negative right at the
        handoff point -- confirmed with all three smile_method options, so it
        is not specific to one spline choice. smile_method="linear" is the
        most exposed to this generally (see its note below), since it has no
        curvature smoothing anywhere, not just at this handoff. Always run
        check_arbitrage() after building a pillar_step=5 surface on long-dated
        or highly skewed tenors; this is accepted as a known, usually small,
        trade-off of this fill method rather than something this code works
        around automatically.
        """
        disc = df_for if self.delta_type == "spot" else 1.0
        k10p, k25p, katm, k25c, k10c = base_strikes
        v10p, v25p, vatm, v25c, v10c = base_vols
        pa = self.premium_adjusted

        def solve(target: float, k1: float, v1: float, k2: float, v2: float, is_call: bool) -> Tuple[float, float]:
            return solve_pillar_on_line(target, k1, v1, k2, v2, fwd, te, disc, is_call, pa)

        def bracketed(target: float, k1: float, v1: float, k2: float, v2: float, is_call: bool) -> bool:
            f1 = black_delta(k1, v1, fwd, te, disc, is_call, pa) - target
            f2 = black_delta(k2, v2, fwd, te, disc, is_call, pa) - target
            return f1 * f2 < 0

        def fill_gaps(wing_order: List[str], nominal_x: Dict[str, int], pts: Dict[str, Tuple[float, float]],
                     atm_pt: Tuple[float, float]) -> None:
            """In-place: fill any name in wing_order missing from pts via linear
            interpolation against nominal_x, using the nearest ORIGINALLY solved
            neighbours on each side (never a neighbour that was itself filled
            here, so simultaneous gaps interpolate independently)."""
            solved = dict(pts)  # snapshot: delta-solved pillars only, before any filling
            for i, name in enumerate(wing_order):
                if name in pts:
                    continue
                left = next((wing_order[j] for j in range(i - 1, -1, -1) if wing_order[j] in solved), None)
                right_name, right_pt = "ATM", atm_pt
                for j in range(i + 1, len(wing_order)):
                    if wing_order[j] in solved:
                        right_name, right_pt = wing_order[j], solved[wing_order[j]]
                        break
                left_name, left_pt = left, solved[left]
                x_l, x_r, x_i = nominal_x[left_name], nominal_x.get(right_name, 50), nominal_x[name]
                frac = (x_i - x_l) / (x_r - x_l)
                k = left_pt[0] + frac * (right_pt[0] - left_pt[0])
                v = left_pt[1] + frac * (right_pt[1] - left_pt[1])
                pts[name] = (k, v)

        put_pts = {"10P": (k10p, v10p), "25P": (k25p, v25p)}
        put_pts["05P"] = (strike_from_delta(-0.05, v10p, fwd, te, df_for, self.delta_type, pa), v10p)
        for name, target in [("15P", -0.15), ("20P", -0.20)]:
            put_pts[name] = solve(target, k10p, v10p, k25p, v25p, is_call=False)
        for name, target in [("30P", -0.30), ("35P", -0.35), ("40P", -0.40), ("45P", -0.45)]:
            if bracketed(target, k25p, v25p, katm, vatm, is_call=False):
                put_pts[name] = solve(target, k25p, v25p, katm, vatm, is_call=False)
        fill_gaps(["25P", "30P", "35P", "40P", "45P"],
                 {"25P": 25, "30P": 30, "35P": 35, "40P": 40, "45P": 45}, put_pts, (katm, vatm))

        call_pts = {"25C": (k25c, v25c), "10C": (k10c, v10c)}
        call_pts["05C"] = (strike_from_delta(0.05, v10c, fwd, te, df_for, self.delta_type, pa), v10c)
        for name, target in [("45C", 0.45), ("40C", 0.40), ("35C", 0.35), ("30C", 0.30)]:
            if bracketed(target, katm, vatm, k25c, v25c, is_call=True):
                call_pts[name] = solve(target, katm, vatm, k25c, v25c, is_call=True)
        for name, target in [("20C", 0.20), ("15C", 0.15)]:
            call_pts[name] = solve(target, k25c, v25c, k10c, v10c, is_call=True)
        fill_gaps(["25C", "30C", "35C", "40C", "45C"],
                 {"25C": 25, "30C": 30, "35C": 35, "40C": 40, "45C": 45}, call_pts, (katm, vatm))

        names = EXT5_PILLAR_NAMES
        strikes = np.array([put_pts[n][0] for n in EXT5_PUT_NAMES] + [katm]
                          + [call_pts[n][0] for n in EXT5_CALL_NAMES])
        vols = np.array([put_pts[n][1] for n in EXT5_PUT_NAMES] + [vatm]
                       + [call_pts[n][1] for n in EXT5_CALL_NAMES])
        return strikes, vols, names

    # -- public API --
    def vol(self, expiry: date, strike, delivery: Optional[date] = None):
        """Implied vol (decimal) for `expiry` and `strike` (scalar or array)."""
        out = self.build_smile(expiry, delivery).vol(strike)
        return float(out) if np.ndim(out) == 0 else out

    def pillar_table(self) -> List[dict]:
        """Strikes and vols of the pillars (5, or up to 19 if pillar_step=5) for
        every quoted maturity. With pillar_step=5, a long tenor whose ATM point
        doesn't reach 45-delta will simply be missing those K_/vol_ keys for
        that row -- see _expand_pillars for why."""
        rows = []
        for q in self.quotes:
            s = self.build_smile(q.expiry)
            row = {"expiry": q.expiry, "Te": s.te, "forward": s.fwd}
            for name, k, v in zip(s.names, s.strikes, s.vols):
                row[f"K_{name}"], row[f"vol_{name}"] = float(k), float(v)
            rows.append(row)
        return rows

    # -- no-arbitrage diagnostics --
    def check_arbitrage(self, k_range: Tuple[float, float] = (-0.20, 0.20), n: int = 401,
                        tol: float = 1e-9) -> List[str]:
        warnings: List[str] = []
        ks = np.linspace(k_range[0], k_range[1], n)

        # Butterfly: d2C/dK2 >= 0 on a uniform strike grid
        for q in self.quotes:
            s = self.build_smile(q.expiry)
            strikes = np.linspace(s.fwd * math.exp(k_range[0]), s.fwd * math.exp(k_range[1]), n)
            c = black_call(s.fwd, strikes, s.vol(strikes), s.te, s.df_dom)
            dk = strikes[1] - strikes[0]
            density = (c[2:] - 2 * c[1:-1] + c[:-2]) / dk ** 2 / s.df_dom
            if density.min() < -tol:
                bad = strikes[1:-1][density.argmin()]
                warnings.append(f"Butterfly: negative density at {q.expiry} (min {density.min():.2e} near K={bad:.2f})")

        # Calendar: total variance at fixed log-forward-moneyness must not decrease
        smiles = [self.build_smile(q.expiry) for q in self.quotes]
        for a, b, qa, qb in zip(smiles[:-1], smiles[1:], self.quotes[:-1], self.quotes[1:]):
            wa = a.vol(a.fwd * np.exp(ks)) ** 2 * a.te
            wb = b.vol(b.fwd * np.exp(ks)) ** 2 * b.te
            if np.any(wb < wa - tol):
                k_bad = ks[np.argmin(wb - wa)]
                warnings.append(f"Calendar: total variance falls {qa.expiry} -> {qb.expiry} (worst at ln(K/F)={k_bad:+.3f})")
        return warnings


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------

def export_surface_csv(surf: FxVolSurface, prefix: str = "fx_vol_surface",
                       strikes: Optional[Sequence[float]] = None) -> Tuple[str, str]:
    """Write the surface to two CSV files (vols in percent).

    <prefix>_pillars.csv : one row per quoted tenor - expiry, Te, forward and the
                           strike / vol of each pillar (5: 10P,25P,ATM,25C,10C; or
                           19 if the surface was built with pillar_step=5).
    <prefix>_grid.csv    : matrix, one row per quoted tenor, one column per absolute
                           strike in `strikes`, cells = interpolated vol (%).
                           Default strikes: 1000 to 2000 in steps of 25.
    """
    import csv

    pillars_path, grid_path = f"{prefix}_pillars.csv", f"{prefix}_grid.csv"

    rows = surf.pillar_table()
    # Canonical full name set for this surface's pillar_step -- NOT just the
    # first tenor's, since a long tenor can be missing far pillars (capped
    # when the ATM anchor's own delta can't reach them; see _expand_pillars).
    names = EXT5_PILLAR_NAMES if surf.pillar_step == 5 else PILLAR_NAMES
    header = ["expiry", "Te", "forward"] + [f"{kind}_{n}" for n in names for kind in ("K", "vol_pct")]
    with open(pillars_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        for r in rows:
            out = [r["expiry"].isoformat(), round(r["Te"], 8), round(r["forward"], 6)]
            for n in names:
                if f"K_{n}" in r:
                    out += [round(r[f"K_{n}"], 6), round(r[f"vol_{n}"] * 100, 8)]
                else:
                    out += ["", ""]   # pillar not reachable for this tenor (capped -- see _expand_pillars)
            w.writerow(out)

    strikes = np.arange(1000.0, 2000.0 + 1e-9, 25.0) if strikes is None else np.asarray(strikes, dtype=float)
    with open(grid_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["expiry", "Te"] + [f"{k:g}" for k in strikes])
        for q in surf.quotes:
            smile = surf.build_smile(q.expiry)
            w.writerow([q.expiry.isoformat(), round(smile.te, 8)] + [round(float(v) * 100, 6) for v in smile.vol(strikes)])

    return pillars_path, grid_path


# ---------------------------------------------------------------------------
# Plot
# ---------------------------------------------------------------------------

def plot_surface(surf: FxVolSurface, path: str = "fx_vol_surface.png") -> None:
    import matplotlib.pyplot as plt

    fig = plt.figure(figsize=(14, 5.8))

    # Left: smile per quoted tenor (vol vs strike), pillar points marked
    ax = fig.add_subplot(1, 2, 1)
    cmap = plt.get_cmap("viridis")
    for i, q in enumerate(surf.quotes):
        s = surf.build_smile(q.expiry)
        strikes = np.linspace(s.strikes[0] * 0.97, s.strikes[-1] * 1.03, 200)
        color = cmap(i / max(1, len(surf.quotes) - 1))
        ax.plot(strikes, s.vol(strikes) * 100, color=color, lw=1.6, label=q.expiry.isoformat())
        ax.scatter(s.strikes, s.vols * 100, color=color, s=14, zorder=3)
    ax.set_xlabel("Strike")
    ax.set_ylabel("Implied vol (%)")
    ax.set_title("Smile by maturity (dots = 10P/25P/ATM/25C/10C pillars)")
    ax.grid(alpha=0.3)
    ax.legend(fontsize=7, ncol=2, title="expiry")

    # Right: surface, x = ln(K/F) / (ATM vol * sqrt(Te)) (comparable across maturities),
    # y = expiry (years), z = vol
    ax3 = fig.add_subplot(1, 2, 2, projection="3d")
    expiries_y = np.linspace(surf.T[0], surf.T[-1], 45)
    m = np.linspace(-1.4, 1.4, 41)
    z = np.empty((len(expiries_y), len(m)))
    for i, ty in enumerate(expiries_y):
        exp_date = surf.val_date + timedelta(days=int(round(ty * 365)))
        s = surf.build_smile(exp_date)
        k = m * s.atm_vol * math.sqrt(s.te)
        z[i] = s.vol(s.fwd * np.exp(k)) * 100
    mm, tt = np.meshgrid(m, expiries_y)
    ax3.plot_surface(mm, tt, z, cmap="viridis", linewidth=0, antialiased=True)
    ax3.set_xlabel("ln(K/F) / (ATM vol*sqrt(Te))")
    ax3.set_ylabel("Expiry (years)")
    ax3.set_zlabel("Vol (%)")
    ax3.set_title("Vol surface")
    fig.tight_layout()
    fig.savefig(path, dpi=150)


# ---------------------------------------------------------------------------
# Example -- PLACEHOLDER market data: replace with your own quotes
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    val_date = date(2026, 9, 17)
    spot_date = date(2026, 9, 21)
    spot = 1382.20

    expiries = [
        date(2026, 9, 24), date(2026, 10, 19), date(2026, 11, 19), date(2026, 12, 17),
        date(2027, 3, 18), date(2027, 6, 17), date(2027, 9, 17), date(2028, 9, 19),
        date(2029, 9, 19), date(2031, 9, 18),
    ]
    # vols in percent; RR = call - put; BF = 0.5*(call+put) - ATM
    atm  = [10.5200, 10.1710, 10.1750, 10.1900, 9.9800, 9.9240, 9.9400, 10.1980, 10.3130, 10.5380]
    rr25 = [0.4430, 0.4680, 0.5100, 0.5830, 0.6800, 0.7620, 0.8110, 0.9270, 1.0040, 1.0825]
    bf25 = [0.1755, 0.1900, 0.2030, 0.2260, 0.2570, 0.2910, 0.3350, 0.3900, 0.4225, 0.4590]
    rr10 = [0.8400, 0.8570, 0.9680, 1.0880, 1.3560, 1.4810, 1.6550, 1.8070, 2.0570, 2.2815]
    bf10 = [0.5960, 0.6400, 0.7030, 0.7500, 0.8760, 0.9890, 1.1440, 1.3230, 1.4330, 1.5580]

    # Zero rate curves -- (tenor in years, decimal rate), own tenor grids,
    # linearly interpolated (flat beyond the ends). Replace with your own.
    foreign_curve = RateCurve([   # USD
        (0.00273973, 0.03690369), (0.09863014, 0.03932408), (0.18904110, 0.03987406),
        (0.26575342, 0.04040737), (0.51506849, 0.04190994), (0.76438356, 0.04323944),
        (1.01643836, 0.04420281), (1.51506849, 0.04506122), (2.02465753, 0.04531277),
        (3.02465753, 0.04529400), (4.02465753, 0.04515969), (5.02191781, 0.04508102),
        (6.02191781, 0.04507179), (7.02191781, 0.04511237), (8.02739726, 0.04521183),
        (9.02739726, 0.04535818), (10.02739726, 0.04552743), (12.02465753, 0.04596804),
        (15.03287671, 0.04664988), (20.03561644, 0.04713717), (25.03835616, 0.04663614),
        (30.04383562, 0.04557059), (40.04383562, 0.04284475), (50.05205479, 0.03963063),
    ])
    domestic_curve = RateCurve([  # KRW
        (0.00273973, 0.02766200), (0.01095890, 0.02786911), (0.03013699, 0.02871253),
        (0.09315068, 0.03304886), (0.18356164, 0.03357098), (0.26027397, 0.03358150),
        (0.50958904, 0.03371147), (0.75890411, 0.03446875), (1.01643836, 0.03504543),
        (1.51506849, 0.03601875), (2.02465753, 0.03612013), (3.03013699, 0.03689988),
        (4.02465753, 0.03715662), (5.02191781, 0.03677475), (6.02465753, 0.03677794),
        (7.02191781, 0.03716942), (8.02739726, 0.03739190), (9.02739726, 0.03709087),
        (10.02739726, 0.03609148), (12.02465753, 0.03570299), (15.03287671, 0.03543245),
        (20.03561644, 0.03450087),
    ])

    quotes = [TenorQuote(e, a, r25, b25, r10, b10)
              for e, a, r25, b25, r10, b10 in zip(expiries, atm, rr25, bf25, rr10, bf10)]

    # Sample public holidays (2026) -- KRW (South Korea) and USD (US federal).
    # USDKRW settlement is affected by non-business days in BOTH centers, so
    # the two lists are unioned below. Source: officeholidays.com / OPM;
    # verify against your own settlement calendar before relying on this for
    # real trades, and extend into later years for tenors beyond 2026.
    KRW_HOLIDAYS_2026 = [
        date(2026, 1, 1),                                   # New Year's Day
        date(2026, 2, 16), date(2026, 2, 17), date(2026, 2, 18),  # Seollal (Lunar New Year)
        date(2026, 3, 2),                                   # Independence Movement Day (observed; Mar 1 is a Sunday)
        date(2026, 5, 1),                                   # Labor Day
        date(2026, 5, 5),                                   # Children's Day
        date(2026, 5, 25),                                  # Buddha's Birthday (observed; May 24 is a Sunday)
        date(2026, 6, 6),                                   # Memorial Day (Saturday)
        date(2026, 7, 17),                                  # Constitution Day
        date(2026, 8, 17),                                  # Liberation Day (observed; Aug 15 is a Saturday)
        date(2026, 9, 24), date(2026, 9, 25), date(2026, 9, 26),  # Chuseok (Harvest Festival)
        date(2026, 10, 5),                                  # National Foundation Day (observed; Oct 3 is a Saturday)
        date(2026, 10, 9),                                  # Hangeul Day
        date(2026, 12, 25),                                 # Christmas Day
        date(2029, 9, 21),
        date(2029, 9, 24),
    ]
    USD_HOLIDAYS_2026 = [
        date(2026, 1, 1),    # New Year's Day
        date(2026, 1, 19),   # Martin Luther King Jr. Day
        date(2026, 2, 16),   # Washington's Birthday (Presidents' Day)
        date(2026, 5, 25),   # Memorial Day
        date(2026, 6, 19),   # Juneteenth
        date(2026, 7, 3),    # Independence Day (observed; Jul 4 is a Saturday)
        date(2026, 9, 7),    # Labor Day
        date(2026, 10, 12),  # Columbus Day
        date(2026, 11, 11),  # Veterans Day
        date(2026, 11, 26),  # Thanksgiving Day
        date(2026, 12, 25),  # Christmas Day
    ]
    holidays = sorted(set(KRW_HOLIDAYS_2026) | set(USD_HOLIDAYS_2026))

    surf = FxVolSurface(val_date, spot_date, spot, quotes, domestic_curve, foreign_curve,
                        delta_type="spot", atm_convention="dns", smile_method="linear",
                        holidays=holidays)

    print(f"{'expiry':<12}{'Te':>6}{'fwd':>10}  " + "  ".join(f"K_{n:<3}/vol%" for n in PILLAR_NAMES))
    for r in surf.pillar_table():
        cells = "  ".join(f"{r['K_' + n]:8.4f}/{r['vol_' + n] * 100:5.4f}" for n in PILLAR_NAMES)
        print(f"{r['expiry'].isoformat():<12}{r['Te']:6.3f}{r['forward']:10.4f}  {cells}")

    print()
    for exp, k in [(date(2027, 9, 17), 1357.1996), (date(2027, 1, 15), 1380.0), (date(2028, 3, 20), 1300.0)]:
        print(f"vol({exp}, K={k}) = {surf.vol(exp, k):.4%}")

    warns = surf.check_arbitrage()
    print("\nArbitrage checks:", "OK" if not warns else "")
    for w in warns:
        print("  WARNING:", w)

    plot_surface(surf, "fx_vol_surface.png")
    print("\nSaved fx_vol_surface.png")

    pillars_csv, grid_csv = export_surface_csv(surf, "fx_vol_surface")
    print(f"Saved {pillars_csv}, {grid_csv}")

    # -----------------------------------------------------------------
    # Extended 5-delta-step pillar grid (05P..45P, ATM, 45C..05C)
    # -----------------------------------------------------------------
    surf_ext = FxVolSurface(val_date, spot_date, spot, quotes, domestic_curve, foreign_curve,
                            delta_type="spot", atm_convention="dns", smile_method="linear",
                            holidays=holidays, pillar_step=5)
    ext_rows = surf_ext.pillar_table()

    print("\nExtended pillar strikes")
    print(f"{'expiry':<12}" + "".join(f"{n:>10}" for n in EXT5_PILLAR_NAMES))
    for q, r in zip(quotes, ext_rows):
        cells = "".join(f"{r[f'K_{n}']:>10.4f}" if f"K_{n}" in r else f"{'--':>10}" for n in EXT5_PILLAR_NAMES)
        print(f"{q.expiry.isoformat():<12}{cells}")

    print("\nExtended pillar vols (%)")
    print(f"{'expiry':<12}" + "".join(f"{n:>10}" for n in EXT5_PILLAR_NAMES))
    for q, r in zip(quotes, ext_rows):
        cells = "".join(f"{r[f'vol_{n}']*100:>10.4f}" if f"vol_{n}" in r else f"{'--':>10}" for n in EXT5_PILLAR_NAMES)
        print(f"{q.expiry.isoformat():<12}{cells}")

    ext_warns = surf_ext.check_arbitrage()
    print("\nExtended surface arbitrage checks:", "OK" if not ext_warns else "")
    for w in ext_warns:
        print("  WARNING:", w)

    ext_pillars_csv, ext_grid_csv = export_surface_csv(surf_ext, "fx_vol_surface_ext")
    print(f"Saved {ext_pillars_csv}, {ext_grid_csv}")





