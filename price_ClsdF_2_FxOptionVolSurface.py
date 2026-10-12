"""
price_from_extended_surface.py

Prices vanilla FX options off the EXTENDED vol surface (fx_vol_surface.py,
pillar_step=5 -- 19 pillars per quoted tenor) as the primary surface, with
vol looked up by plain 2D LINEAR interpolation on tenor and strike, done
directly against each quoted tenor's own 19-pillar (strike, vol) grid.

This is a DIFFERENT, simpler lookup than FxVolSurface.vol()'s own default
behaviour. FxVolSurface.vol() interpolates the five delta-pillar vols across
tenor in TOTAL VARIANCE, then rebuilds a fresh smile (re-solving strikes) at
the target expiry and interpolates across strike with smile_method
(cubic/pchip/linear). Here, instead, each quoted tenor's smile is built
ONCE (so its 19 strikes are fixed, as actually quoted), and a target
(expiry, strike) is priced by:

    1. Bracketing `expiry` between two quoted tenor dates (flat/clamped at
       the ends -- querying before the first or after the last quoted
       tenor just uses that tenor's own grid, no extrapolation).
    2. In EACH of those two tenors' own 19-pillar grids, linearly
       interpolating vol at `strike` directly (flat beyond that tenor's own
       outermost pillar strike -- again no extrapolation).
    3. Linearly interpolating those two vols against Te (plain linear in
       time, not total variance).

Forward and discounting still come from FxVolSurface's own RateCurve-driven
day-count logic (surf.build_smile(expiry) for Td/Te/forward); only the vol
itself is replaced by this grid lookup. Pricing/Greeks then come from
fx_option_pricing.py.
"""

import math
from datetime import date

import numpy as np

from smile_VolSurface import FxVolSurface, TenorQuote, RateCurve, yearfrac
from price_ClsdF_1_FxOption import mx_fx_option_pricer, mx_fx_option_grk_calc, ccy_to_prem_type, greeks_quote_type


def extended_surface_vol(surf: FxVolSurface, expiry: date, strike: float) -> float:
    """Vol at (expiry, strike): linear interpolation on tenor AND strike,
    directly against the quoted tenors' own extended (pillar_step=5) pillar
    grids. See module docstring for the exact steps."""
    if surf.pillar_step != 5:
        raise ValueError("extended_surface_vol requires a surface built with pillar_step=5")

    te = yearfrac(surf.val_date, expiry)
    T = surf.T  # quoted tenor year-fractions, ascending

    if te <= T[0]:
        i_lo = i_hi = 0
        frac = 0.0
    elif te >= T[-1]:
        i_lo = i_hi = len(T) - 1
        frac = 0.0
    else:
        i_hi = int(np.searchsorted(T, te))
        i_lo = i_hi - 1
        frac = (te - T[i_lo]) / (T[i_hi] - T[i_lo])

    def vol_at_tenor(i: int) -> float:
        q = surf.quotes[i]
        s = surf.build_smile(q.expiry)                    # this tenor's own 19 pillars, as quoted
        return float(np.interp(strike, s.strikes, s.vols))  # linear in strike; flat beyond the outer pillars

    v_lo, v_hi = vol_at_tenor(i_lo), vol_at_tenor(i_hi)
    return v_lo + frac * (v_hi - v_lo)


def price_fx_vanilla(surf: FxVolSurface, o_type: str, expiry: date, strike: float,
                     prem_ccy: str = "%USD", grk_quo_ccy: str = "%USD" ,with_greeks: bool = True) -> dict:
    """Prices a vanilla FX option at (expiry, strike) off the extended
    surface. Forward/Td/Te/discounting come from `surf`'s own curves; vol
    comes from extended_surface_vol() above. Returns a dict of inputs plus
    price and (optionally) Greeks, all in the `prem_ccy` quoting convention
    ("USD" or "KRW", via fx_option_pricing.ccy_to_prem_type)."""
    s = surf.build_smile(expiry)  # forward / Td / Te / discount factors only
    vol = extended_surface_vol(surf, expiry, strike)

    rd = -math.log(s.df_dom) / s.td
    rf = -math.log(s.df_for) / s.td
    prem_type = ccy_to_prem_type(prem_ccy)
    grk_quo_type = greeks_quote_type(grk_quo_ccy)
    price = mx_fx_option_pricer(o_type, prem_type, surf.spot, s.fwd, strike, vol, s.te, s.td, rf, rd)
    out = {
        "expiry": expiry, "strike": strike, "o_type": o_type.upper(),
        "vol": vol, "forward": s.fwd, "Te": s.te, "Td": s.td, "rd": rd, "rf": rf,
        "price": price,
    }
    if with_greeks:
        args = (o_type, prem_type, grk_quo_type, surf.spot, s.fwd, strike, vol, s.te, s.td, rf, rd)
        out["delta"] = mx_fx_option_grk_calc(*args[:3], 0, *args[3:])
        out["fwd_delta"] = mx_fx_option_grk_calc(*args[:3], 1, *args[3:])
        out["gamma"] = mx_fx_option_grk_calc(*args[:3], 2, *args[3:])
        out["vega"] = mx_fx_option_grk_calc(*args[:3], 3, *args[3:])
        out["theta"] = mx_fx_option_grk_calc(*args[:3], 4, *args[3:])
        out["rho_domestic"] = mx_fx_option_grk_calc(*args[:3], 5, *args[3:])
        out["rho_foreign"] = mx_fx_option_grk_calc(*args[:3], 6, *args[3:])
        out["vanna"] = mx_fx_option_grk_calc(*args[:3], 7, *args[3:])
        out["volga"] = mx_fx_option_grk_calc(*args[:3], 8, *args[3:])
    return out


# ---------------------------------------------------------------------------
# Example -- same placeholder market data as fx_vol_surface.py; replace with yours
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    val_date, spot_date, spot = date(2026, 9, 17), date(2026, 9, 21), 1382.20

    expiries = [
        date(2026, 9, 24), date(2026, 10, 19), date(2026, 11, 19), date(2026, 12, 17),
        date(2027, 3, 18), date(2027, 6, 17), date(2027, 9, 17), date(2028, 9, 19),
        date(2029, 9, 19), date(2031, 9, 18),
    ]
    atm  = [10.5200, 10.1710, 10.1750, 10.1900, 9.9800, 9.9240, 9.9400, 10.1980, 10.3130, 10.5380]
    rr25 = [0.4430, 0.4680, 0.5100, 0.5830, 0.6800, 0.7620, 0.8110, 0.9270, 1.0040, 1.0825]
    bf25 = [0.1755, 0.1900, 0.2030, 0.2260, 0.2570, 0.2910, 0.3350, 0.3900, 0.4225, 0.4590]
    rr10 = [0.8400, 0.8570, 0.9680, 1.0880, 1.3560, 1.4810, 1.6550, 1.8070, 2.0570, 2.2815]
    bf10 = [0.5960, 0.6400, 0.7030, 0.7500, 0.8760, 0.9890, 1.1440, 1.3230, 1.4330, 1.5580]

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
        date(2028, 1,26),
        date(2028, 1, 27),
        date(2028, 1, 28),
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

    surf = FxVolSurface(val_date, spot_date, spot, quotes, domestic_curve, foreign_curve, delta_type="spot", atm_convention="dns", smile_method="linear",
                        holidays=holidays, pillar_step=5)

    # -----------------------------------------------------------------
    # Sanity checks: querying exactly AT a quoted tenor and AT one of its
    # own pillar strikes must reproduce that pillar's vol exactly.
    # -----------------------------------------------------------------
    print("Sanity checks (exact grid points)")
    q = quotes[6]  # 2027-09-17
    s = surf.build_smile(q.expiry)
    for name in ["25P", "ATM", "25C"]:
        i = s.names.index(name)
        v = extended_surface_vol(surf, q.expiry, s.strikes[i])
        print(f"  {q.expiry} {name}: quoted={s.vols[i]*100:.4f}%  looked-up={v*100:.4f}%  "
              f"match={abs(v - s.vols[i]) < 1e-10}")

    # Between two tenors, at a strike, should be the linear blend of each
    # tenor's own interpolated vol at that strike.
    q_a, q_b = quotes[6], quotes[7]  # 2027-09-17 and 2028-09-19
    mid_expiry = date(2028, 1, 24)   # roughly halfway between them
    v_a = extended_surface_vol(surf, q_a.expiry, 1400.0)
    v_b = extended_surface_vol(surf, q_b.expiry, 1400.0)
    te_a, te_b = yearfrac(val_date, q_a.expiry), yearfrac(val_date, q_b.expiry)
    te_mid = yearfrac(val_date, mid_expiry)
    frac = (te_mid - te_a) / (te_b - te_a)
    expected = v_a + frac * (v_b - v_a)
    v_mid = extended_surface_vol(surf, mid_expiry, 1400.0)
    print(f"\nTenor interpolation check at K=1400: expected {expected*100:.4f}%, got {v_mid*100:.4f}%, "
         f"match={abs(expected - v_mid) < 1e-10}")

    # -----------------------------------------------------------------
    # Price a few sample options
    # -----------------------------------------------------------------
    print("\nSample pricing (prem_ccy='USD')")
    samples = [
        ("C", date(2027, 9, 17), 1370.47),  # on-tenor, ~ATM
        ("P", date(2027, 9, 17), 1212.37),  # on-tenor, ~10P strike
        ("C", date(2028, 1, 24), 1400.00),  # off-tenor, off-pillar strike
        ("P", date(2030, 6, 15), 1300.00),  # off-tenor, between 3Y and 5Y
    ]
    for o_type, expiry, strike in samples:
        r = price_fx_vanilla(surf, o_type, expiry, strike, prem_ccy="%KRW", grk_quo_ccy="%KRW")
        print(f"  {o_type} {expiry} K={strike:<10.4f} vol={r['vol']*100:18.14f}%"
             f"  fwd={r['forward']:8.4f}  price={r['price']:8.6f}"
             f"  delta={r['delta']:8.4f}  fwd_delta={r['fwd_delta']:8.4f}  gamma={r['gamma']:8.4f}  vega={r['vega']:6.4f}  theta={r['theta']:8.4f}"
             f"  rho_rd={r['rho_domestic']:8.4f}  rho_rf={r['rho_foreign']:8.4f} vanna={r['vanna']:8.4f}  volga={r['volga']:8.4f}")




