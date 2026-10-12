"""
fx_option_pricing.py

Python conversion of a VBA option-pricing library (FX / quanto vanilla and
barrier options, Greeks, vol interpolation, day-count conventions, and a
double-barrier reflection-principle model).

Conversion notes (read before using in production)
----------------------------------------------------
1. The original module used `Option Compare Text`, which makes VBA string
   comparisons (`=`, `Select Case`) case-insensitive. Every string
   comparison below is normalized with `.upper()` to reproduce that.

2. `CND(x)` is reproduced *exactly* as the original polynomial
   (Zelen & Severo) approximation of the normal CDF, not `scipy`'s exact
   implementation — because the pricing functions below were built and
   presumably calibrated against that specific approximation. A
   `norm_cdf_exact` wrapper (scipy) is also provided if you want to swap
   it in; a smoke test at the bottom compares the two.

3. `FXD_BSModel` contains `oType = "=c" Or oType = "C"` in the original.
   This is almost certainly a typo for `oType = "c"` (a lowercase "c"
   alone can never equal the literal string "=c"). It is reproduced
   literally below (see `NOTE (possible bug)` comments) so behavior
   matches the legacy sheet; you may want to fix this once confirmed
   against the source system.

4. `mxFxOptionVolCalc` and `MyLinearIntp` took Excel `Range` objects.
   They're translated to take plain sequences (lists / numpy arrays)
   instead — the interpolation logic is unchanged.

5. `VanillaCalc_Click` is an Excel button handler that reads/writes
   worksheet cells directly. It has no meaningful non-Excel translation,
   so it is **not** ported; `price_vanilla_example` is provided instead
   as a parameterized illustration of the same call sequence
   (standard vs. quanto branch).

6. `MakeNormS` (Box-Muller normal RNG with VBA `Static` state) is ported
   using Python's own state-free `random.gauss`, since replicating the
   exact `Static iset/gset` caching behavior has no benefit in Python.

Structs -> dataclasses, VBA enums -> Python `Enum` (member *names* are
preserved; since the underlying integer values were never shown in the
VBA source, only relative equality inside this module matters).
"""

from __future__ import annotations

import math
import random
from dataclasses import dataclass, field
from datetime import date
from enum import Enum, auto
from typing import Optional, Sequence

# scipy is optional; only used for the cross-check / `norm_cdf_exact` helper.
try:
    from scipy.stats import norm as _scipy_norm
except ImportError:  # pragma: no cover
    _scipy_norm = None


PI = 3.14159265358979  # matches the literal constant used throughout the VBA


# ---------------------------------------------------------------------------
# Enums (VBA: enumBarrType, enumPreStyle, enumRebPayment, enumOptCategry,
# enumOptType)
# ---------------------------------------------------------------------------

class BarrierType(Enum):
    UP_AND_IN = auto()
    UP_AND_OUT = auto()
    DOWN_AND_IN = auto()
    DOWN_AND_OUT = auto()


class RebateStyle(Enum):
    """VBA enumPreStyle: eKRO / eKROUSD / eUSD / eUSDKRO / eQNT / eUSDQNT."""
    KRO = auto()        # rebate quoted as % of strike (KRW-quoted deals)
    KRO_USD = auto()    # rebate quoted as flat USD amount
    USD = auto()        # rebate quoted as % of strike (USD-quoted deals)
    USD_KRO = auto()    # rebate quoted as flat KRW amount
    QNT = auto()        # rebate quoted as % of strike (quanto)
    USD_QNT = auto()    # rebate quoted as flat amount (quanto)


class RebatePaymentTiming(Enum):
    AT_MATURITY = auto()
    AT_KNOCK_TIME = auto()


class OptionCategory(Enum):
    STANDARD = auto()
    QUANTO = auto()


# ---------------------------------------------------------------------------
# Structs (VBA: DealInfo, UnderInfo, QntoInfo, SimpBarrInfo)
# ---------------------------------------------------------------------------

@dataclass
class DealInfo:
    val_date: date = None       # Valdate
    spot_date: date = None      # SpotDate
    mat_date: date = None       # Matdate
    del_date: date = None       # Deldate
    dc: float = 365.0           # day-count divisor used for Te/Td


@dataclass
class UnderInfo:
    spot: float = 0.0
    strike: float = 0.0
    forward: float = 0.0
    vol: float = 0.0
    rd: float = 0.0   # domestic rate
    rf: float = 0.0   # foreign rate


@dataclass
class QuantoInfo:
    ccy: str = ""
    fac: float = 0.0
    df: float = 0.0            # QtoInfo.DF -- used as the effective quanto discount rate
    vol_qto_de: float = 0.0    # VolQtoDe
    vol_qto_num: float = 0.0   # VolQtoNum
    qto_corr: float = 0.0


@dataclass
class SimpleBarrierInfo:
    bar_type: BarrierType = BarrierType.UP_AND_IN
    bar_lvl: float = 0.0
    reb_type: RebateStyle = RebateStyle.KRO_USD
    rebate: float = 0.0
    reb_pmt_at: RebatePaymentTiming = RebatePaymentTiming.AT_MATURITY


# ---------------------------------------------------------------------------
# Core statistical helpers
# ---------------------------------------------------------------------------

def cnd(x: float) -> float:
    """Standard normal CDF, reproducing the VBA `CND` polynomial
    (Zelen & Severo) approximation exactly, including its symmetry trick
    for negative x."""
    a1, a2, a3, a4, a5 = (
        0.31938153,
        -0.356563782,
        1.781477937,
        -1.821255978,
        1.330274429,
    )
    l = abs(x)
    k = 1.0 / (1.0 + 0.2316419 * l)
    result = 1.0 - 1.0 / math.sqrt(2 * PI) * math.exp(-l ** 2 / 2) * (
        a1 * k + a2 * k ** 2 + a3 * k ** 3 + a4 * k ** 4 + a5 * k ** 5
    )
    if x < 0:
        result = 1.0 - result
    return result


def norm_cdf_exact(x: float) -> float:
    """Exact normal CDF (scipy), provided for comparison / optional swap-in."""
    if _scipy_norm is None:
        raise ImportError("scipy is required for norm_cdf_exact")
    return float(_scipy_norm.cdf(x))


def pnd(x: float, mu: float, sigma: float) -> float:
    """Normal PDF. Returns NaN (VBA: CVErr) if sigma <= 0."""
    if sigma <= 0:
        return float("nan")
    return (1.0 / (sigma * math.sqrt(2 * PI))) * math.exp(-((x - mu) ** 2) / (2 * sigma ** 2))


def my_pow(base: float, arg: float) -> float:
    """VBA `myPow` wrapped Excel's WorksheetFunction.Power; `**` is equivalent
    for the real-valued (bar_lvl/spot)^exponent usage in this module."""
    return base ** arg


# `max`/`min` in VBA took exactly two doubles; Python's built-ins already
# do this, so no wrapper is defined (avoid shadowing the builtins).


# ---------------------------------------------------------------------------
# Vanilla / quanto vanilla pricers
# ---------------------------------------------------------------------------

def fxd_bs_model(
    o_type: str, f0: float, x: float, sigma: float, te: float, td: float, rd: float, rf: float
) -> float:
    """VBA `FXD_BSModel`.

    NOTE (possible bug, reproduced literally): the original checks
    `oType = "=c" Or oType = "C"` for calls. Because of
    `Option Compare Text`, "C" and "c" ARE equivalent, but "=c" is a
    different string entirely, so a plain lowercase "c" will actually
    fall through to the `Else` branch unless it's exactly "C" or the
    literal "=c". Reproduced as-is below; verify against the live sheet
    before relying on it for lowercase 'c' inputs.
    """
    a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
    a2 = a1 - sigma * math.sqrt(te)

    ot = o_type.upper()
    if o_type == "=c" or ot == "C":
        return f0 * math.exp(-rf * td) * cnd(a1) - x * math.exp(-rf * td) * cnd(a2)
    elif ot == "P":
        return x * math.exp(-rf * td) * cnd(-a2) - f0 * math.exp(-rf * td) * cnd(-a1)
    else:
        raise ValueError("Error Otype")


def bs_model(o_type: str, s0: float, x: float, sigma: float, t: float, rf: float, q: float) -> float:
    """VBA `BSModel`: standard Black-Scholes with continuous dividend/carry q."""
    a1 = (math.log(s0 / x) + (rf - q + sigma ** 2 / 2) * t) / (sigma * math.sqrt(t))
    a2 = a1 - sigma * math.sqrt(t)

    ot = o_type.upper()
    if ot == "C":
        return s0 * math.exp(-q * t) * cnd(a1) - x * math.exp(-rf * t) * cnd(a2)
    elif ot == "P":
        return x * math.exp(-rf * t) * cnd(-a2) - s0 * math.exp(-q * t) * cnd(-a1)
    else:
        raise ValueError("Error Otype")


def fxd_digi_model(
    digi_type: str, o_type: str, f0: float, x: float, sigma: float, te: float, td: float, rd: float, rf: float
) -> float:
    """VBA `FXD_DigiModel`. digi_type: 'DC' (cash-or-nothing) / 'DA' (asset-or-nothing)."""
    a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
    a2 = a1 - sigma * math.sqrt(te)

    dt = digi_type.upper()
    ot = o_type.upper()
    if dt == "DC":
        if ot == "C":
            return x * math.exp(-rf * td) * cnd(a2)
        elif ot == "P":
            return x * math.exp(-rf * td) * cnd(-a2)
        else:
            raise ValueError("Error Otype")
    elif dt == "DA":
        if ot == "C":
            return f0 * math.exp(-rf * td) * cnd(a1)
        elif ot == "P":
            return f0 * math.exp(-rf * td) * cnd(-a1)
        else:
            raise ValueError("Error Otype")
    else:
        raise ValueError("Unknown DigiType")


def ccy_to_prem_type(ccy: str) -> int:
    """Maps a currency code to the 0/1 convention used by `prem_type` in
    `mx_fx_option_pricer` and `grk_quo_type`/`prem_type` in
    `mx_fx_option_grk_calc`: USD -> 0, KRW -> 1.

    Example:
        prem_type = ccy_to_prem_type("USD")   # 0
        prem_type = ccy_to_prem_type("krw")   # 1 (case-insensitive)
    """
    c = ccy.strip().upper()
    if c == "%USD":
        return 0
    elif c == "%KRW":
        return 1
    elif c == "USD/KRW":
        return 2
    elif c == "KRW/USD":
        return 3
    else:
        raise ValueError(f"Unsupported currency for prem_type mapping: {ccy!r} (expected 'USD' or 'KRW')")

def greeks_quote_type(ccy: str) -> int:
    """Maps a currency code to the 0/1 convention used by `prem_type` in
    `mx_fx_option_pricer` and `grk_quo_type`/`prem_type` in
    `mx_fx_option_grk_calc`: USD -> 0, KRW -> 1.

    Example:
        prem_type = ccy_to_prem_type("USD")   # 0
        prem_type = ccy_to_prem_type("krw")   # 1 (case-insensitive)
    """
    c = ccy.strip().upper()
    if c == "%USD":
        return 0
    elif c == "%KRW":
        return 1
    else:
        raise ValueError(f"Unsupported currency for prem_type mapping: {ccy!r} (expected 'USD' or 'KRW')")

def greeks_type(ccy: str) -> int:
    """Maps a currency code to the 0/1 convention used by `prem_type` in
    `mx_fx_option_pricer` and `grk_quo_type`/`prem_type` in
    `mx_fx_option_grk_calc`: USD -> 0, KRW -> 1.

    Example:
        prem_type = ccy_to_prem_type("USD")   # 0
        prem_type = ccy_to_prem_type("krw")   # 1 (case-insensitive)
    """
    c = ccy.strip().upper()
    if c == "DELTA":
        return 0
    elif c == "FORWARD DELTA":
        return 1
    elif c == "GAMMA":
        return 2
    elif c == "VEGA":
        return 3
    elif c == "THETA":
        return 4
    elif c == "RHO RD":
        return 5
    elif c == "RHO RF":
        return 6
    elif c == "VANNA":
        return 7
    elif c == "VOLGA":
        return 8
    else:
        raise ValueError(f"Unsupported greeks type : {ccy!r} ")



def mx_fx_option_pricer(
    o_type: str,
    prem_type: int,
    s0: float,
    f0: float,
    x: float,
    sigma: float,
    te: float,
    td: float,
    rf: float,
    rd: float,
) -> float:
    """VBA `mxFxOptionPricer`. prem_type: 0 = % of foreign notional (F_0-based),
    1 = % of domestic/strike notional (X-based)."""
    a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
    a2 = a1 - sigma * math.sqrt(te)

    ot = o_type.upper()
    if o_type == "=c" or ot == "C":
        if prem_type == 0:
            return (f0 * math.exp(-rf * td) * cnd(a1) - x * math.exp(-rf * td) * cnd(a2)) / f0 * 100
        elif prem_type == 1:
            return (f0 * math.exp(-rd * td) * cnd(a1) - x * math.exp(-rd * td) * cnd(a2)) / x * 100
        elif prem_type == 2:
            return (f0 * math.exp(-rf * td) * cnd(a1) - x * math.exp(-rf * td) * cnd(a2)) / f0 / x
        elif prem_type == 3:
            return f0 * math.exp(-rd * td) * cnd(a1) - x * math.exp(-rd * td) * cnd(a2)
        else:
            raise ValueError("Error premType")
    elif ot == "P":
        if prem_type == 0: # %USD
            return (x * math.exp(-rf * td) * cnd(-a2) - f0 * math.exp(-rf * td) * cnd(-a1)) / f0 * 100
        elif prem_type == 1: # %KRW
            return (x * math.exp(-rd * td) * cnd(-a2) - f0 * math.exp(-rd * td) * cnd(-a1)) / x * 100
        elif prem_type == 2: # USD/KRW
            return (x * math.exp(-rf * td) * cnd(-a2) - f0 * math.exp(-rf * td) * cnd(-a1)) / f0 / x
        elif prem_type == 3: # KRW/USD
            return x * math.exp(-rd * td) * cnd(-a2) - f0 * math.exp(-rd * td) * cnd(-a1)
        else:
            raise ValueError("Error premType")
    else:
        raise ValueError("Error premType")


# ---------------------------------------------------------------------------
# Greeks (VBA: mxFxOptionGrkCalc)
# ---------------------------------------------------------------------------

def mx_fx_option_grk_calc(
    o_type: str,
    prem_type: int,
    grk_quo_type: int,
    grk_type: int,
    s0: float,
    f0: float,
    x: float,
    sigma: float,
    te: float,
    td: float,
    rf: float,
    rd: float,
) -> float:
    """VBA `mxFxOptionGrkCalc`.

    grk_type: 0=Spot Delta, 1=Forward Delta, 2=Gamma, 3=Vega, 4=Theta,
              5/6 = additional rho-style sensitivities, 7/8 = vanna/volga-style
              sensitivities (as in the original, undocumented beyond the
              VBA comments that exist).
    prem_type: 0 = foreign-notional quoting, 1 = domestic/strike-notional quoting.
    grk_quo_type: 0/1 selects which of two quoting conventions to return
                  (mirrors the VBA `If grkQuoType = 0 / ElseIf grkQuoType = 1`).
    """
    ot = o_type.upper()
    a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
    a2 = a1 - sigma * math.sqrt(te)

    if grk_type == 0:  # Spot Delta
        if o_type == "=c" or ot == "C":
            a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0 or prem_type == 2:
                if grk_quo_type == 0:
                    return math.exp(-rf * td) * cnd(a2) * 100 * x / f0
                elif grk_quo_type == 1:
                    return -math.exp(-rd * td) * cnd(a2) * 100
            elif prem_type == 1 or prem_type == 3:
                if grk_quo_type == 0:
                    return math.exp(-rf * td) * cnd(a1) * 100
                elif grk_quo_type == 1:
                    return -math.exp(-rd * td) * cnd(a1) * 100 * f0 / x
            raise ValueError("Error grkType")
        elif ot == "P":
            a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0 or prem_type == 2:
                if grk_quo_type == 0:
                    return -math.exp(-rf * td) * cnd(a1) * 100 * x / f0
                elif grk_quo_type == 1:
                    return math.exp(-rd * td) * cnd(a1) * 100
            elif prem_type == 1 or prem_type == 3:
                if grk_quo_type == 0:
                    return -math.exp(-rf * td) * cnd(a2) * 100
                elif grk_quo_type == 1:
                    return math.exp(-rd * td) * cnd(a2) * 100 * f0 / x
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 1:  # Forward Delta
        if o_type == "=c" or ot == "C":
            a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0 or prem_type == 2:
                if grk_quo_type == 0:
                    return cnd(a2) * 100 * x / f0
                elif grk_quo_type == 1:
                    return -cnd(a2) * 100
            elif prem_type == 1 or prem_type == 3:
                if grk_quo_type == 0:
                    return cnd(a1) * 100
                elif grk_quo_type == 1:
                    return -cnd(a1) * 100 * f0 / x
            raise ValueError("Error grkType")
        elif ot == "P":
            a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0 or prem_type == 2:
                if grk_quo_type == 0:
                    return -cnd(a1) * 100 * x / f0
                elif grk_quo_type == 1:
                    return cnd(a1) * 100
            elif prem_type == 1 or prem_type == 3:
                if grk_quo_type == 0:
                    return - cnd(a2) * 100
                elif grk_quo_type == 1:
                    return cnd(a2) * 100 * f0 / x
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 2:  # Gamma
        if o_type == "=c" or ot == "C":
            if prem_type in (0, 1, 2, 3):
                if grk_quo_type == 0:
                    return math.exp(-rf * td) * pnd(a1, 0, 1) / sigma / math.sqrt(te)
                elif grk_quo_type == 1:
                    return math.exp(-rd * td) * pnd(a1, 0, 1) / sigma * f0 / x / math.sqrt(te)
            raise ValueError("Error grkType")
        elif ot == "P":
            if prem_type in (0, 1, 2, 3):
                if grk_quo_type == 0:
                    return math.exp(-rf * td) * pnd(a2, 0, 1) / sigma * x / f0 / math.sqrt(te)
                elif grk_quo_type == 1:
                    return math.exp(-rd * td) * pnd(a2, 0, 1) / sigma / math.sqrt(te)
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 3:  # Vega
        if o_type == "=c" or ot == "C":
            if prem_type == 0:
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a1, 0, 1)
            elif prem_type == 1:
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a1, 0, 1) * f0 / x
            elif prem_type == 2:
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a1, 0, 1) / x / 100
            elif prem_type == 3:
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a1, 0, 1) * f0 / 100
            raise ValueError("Error grkType")
        elif ot == "P":
            if prem_type == 0:
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a2, 0, 1) * x / f0
            elif prem_type == 1:
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a2, 0, 1)
            elif prem_type == 2:
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a1, 0, 1) / x / 100
            elif prem_type == 3:
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a2, 0, 1) * x / 100
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 4:  # Theta
        if o_type == "=c" or ot == "C":
            if prem_type == 0:
                val = (
                    s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                    + rd * x * math.exp(-rd * td) * cnd(a2)
                    - rf * s0 * math.exp(-rf * td) * cnd(a1)
                )
                return (-val / 365) / s0 * 100
            elif prem_type == 2:
                val = (
                    s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                    + rd * x * math.exp(-rd * td) * cnd(a2)
                    - rf * s0 * math.exp(-rf * td) * cnd(a1)
                )
                return (-val / 365) / s0 / x
            elif prem_type == 1:
                val = (
                    s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                    + rd * x * math.exp(-rd * td) * cnd(a2)
                    - rf * s0 * math.exp(-rf * td) * cnd(a1)
                )
                return (-val / 365) / x * 100
            elif prem_type == 3:
                val = (
                    s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                    + rd * x * math.exp(-rd * td) * cnd(a2)
                    - rf * s0 * math.exp(-rf * td) * cnd(a1)
                )
                return -val / 365
            raise ValueError("Error grkType")
        elif ot == "P":
            if prem_type == 0:
                val = (
                    s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                    - rd * x * math.exp(-rd * td) * cnd(-a2)
                    + rf * s0 * math.exp(-rf * td) * cnd(-a1)
                )
                return (-val / 365) / s0 * 100
            elif prem_type == 2:
                val = (
                        s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                        - rd * x * math.exp(-rd * td) * cnd(-a2)
                        + rf * s0 * math.exp(-rf * td) * cnd(-a1)
                )
                return (-val / 365) / s0 * 100 / x / 100
            elif prem_type == 1:
                val = (
                    s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                    - rd * x * math.exp(-rd * td) * cnd(-a2)
                    + rf * s0 * math.exp(-rf * td) * cnd(-a1)
                )
                return (-val / 365) / x * 100
            elif prem_type == 3:
                val = (
                    s0 * math.exp(-rf * td) * pnd(a1, 0, 1) * sigma * 0.5 / math.sqrt(te)
                    - rd * x * math.exp(-rd * td) * cnd(-a2)
                    + rf * s0 * math.exp(-rf * td) * cnd(-a1)
                )
                return -val / 365
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 5: # Rho rd
        if o_type == "=c" or ot == "C":
            a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0:
                return td * math.exp(-rf * td) * cnd(a2) * x / f0
            elif prem_type == 2:
                return td * math.exp(-rf * td) * cnd(a2) / f0 / 100
            elif prem_type == 1:
                return td * math.exp(-rd * td) * cnd(a2)
            elif prem_type == 3:
                return td * math.exp(-rd * td) * cnd(a2) * x / 100
            raise ValueError("Error grkType")
        elif ot == "P":
            a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0:
                return -td * math.exp(-rf * td) * cnd(a1) * x / f0
            elif prem_type == 2:
                return -td * math.exp(-rf * td) * cnd(a1) / f0 / 100
            elif prem_type == 1:
                return -td * math.exp(-rd * td) * cnd(a1)
            elif prem_type == 3:
                return -td * math.exp(-rd * td) * cnd(a1) * x / 100
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 6: # Rho rf
        if o_type == "=c" or ot == "C":
            a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0:
                return -td * math.exp(-rf * td) * cnd(a1)
            elif prem_type == 2:
                return -td * math.exp(-rf * td) * cnd(a1) / x / 100
            elif prem_type == 1:
                return -td * math.exp(-rd * td) * cnd(a1) * f0 / x
            elif prem_type == 3:
                return -td * math.exp(-rd * td) * cnd(a1) * f0 / 100
            raise ValueError("Error grkType")
        elif ot == "P":
            a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
            a2 = a1 - sigma * math.sqrt(te)
            if prem_type == 0:
                return td * math.exp(-rf * td) * cnd(a2)
            elif prem_type == 2:
                return td * math.exp(-rf * td) * cnd(a2) / x / 100
            elif prem_type == 1:
                return td * math.exp(-rd * td) * cnd(a2) * f0 / x
            elif prem_type == 3:
                return td * math.exp(-rd * td) * cnd(a2) * f0 / 100
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 7: # Vanna; underlying, vol
        if o_type == "=c" or ot == "C":
            if prem_type == 0 or prem_type == 2:
                a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                if grk_quo_type == 0:
                    return math.exp(-rd * td) * pnd(a1, 0, 1) * a2 / sigma * x / s0
                elif grk_quo_type == 1:
                    return -math.exp(-rd * td) * pnd(a1, 0, 1) * a2 / sigma
            elif prem_type == 1 or prem_type == 3:
                a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                if grk_quo_type == 0:
                    return -math.exp(-rf * td) * pnd(a1, 0, 1) * a2 / sigma
                elif grk_quo_type == 1:
                    return math.exp(-rf * td) * pnd(a1, 0, 1) * a2 / sigma * s0 / x
            raise ValueError("Error grkType")
        elif ot == "P":
            if prem_type == 0 or prem_type == 2:
                a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                if grk_quo_type == 0:
                    return math.exp(-rd * td) * pnd(a1, 0, 1) * a2 / sigma * x / s0
                elif grk_quo_type == 1:
                    return -math.exp(-rd * td) * pnd(a1, 0, 1) * a2 / sigma
            elif prem_type == 1 or prem_type == 3:
                a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                if grk_quo_type == 0:
                    return -math.exp(-rf * td) * pnd(a1, 0, 1) * a2 / sigma
                elif grk_quo_type == 1:
                    return math.exp(-rf * td) * pnd(a1, 0, 1) * a2 / sigma * s0 / x
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    elif grk_type == 8: # Volga; vol convexity
        if o_type == "=c" or ot == "C":
            if prem_type == 0:
                a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                # original returns the same expression for both grk_quo_type 0 and 1
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a1, 0, 1) * a1 * a2 / sigma * x / s0 / 100
            elif prem_type == 2:
                a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                # original returns the same expression for both grk_quo_type 0 and 1
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a1, 0, 1) * a1 * a2 / sigma / s0 / 100 / 100
            elif prem_type == 1:
                a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a1, 0, 1) * a1 * a2 / sigma * s0 / x / 100
            elif prem_type == 3:
                a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a1, 0, 1) * a1 * a2 / sigma * s0 / 100 / 100
            raise ValueError("Error grkType")
        elif ot == "P":
            if prem_type == 0:
                a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a1, 0, 1) * a1 * a2 / sigma * x / s0 / 100
            elif prem_type == 2:
                a1 = (math.log(x / f0) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                return math.sqrt(te) * math.exp(-rd * td) * pnd(a1, 0, 1) * a1 * a2 / sigma / s0 / 100 / 100
            elif prem_type == 1:
                a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a1, 0, 1) * a1 * a2 / sigma * s0 / x / 100
            elif prem_type == 3:
                a1 = (math.log(f0 / x) + (sigma ** 2 / 2) * te) / (sigma * math.sqrt(te))
                a2 = a1 - sigma * math.sqrt(te)
                return math.sqrt(te) * math.exp(-rf * td) * pnd(a1, 0, 1) * a1 * a2 / sigma * s0 / 100 / 100
            raise ValueError("Error grkType")
        else:
            raise ValueError("Error grkType")

    else:
        raise ValueError("Unknown grkType")


# ---------------------------------------------------------------------------
# Vol surface interpolation (VBA: mxFxOptionVolCalc, MyLinearIntp2)
# ---------------------------------------------------------------------------

def my_linear_intp2(x1: float, y1: float, x2: float, y2: float, x: float) -> float:
    """VBA `MyLinearIntp2`: two-point linear interpolation."""
    slope = (y2 - y1) / (x2 - x1)
    return y1 + slope * (x - x1)


def mx_fx_option_vol_calc(
    tenor_dates: Sequence[date],
    smile_table: Sequence[Sequence[dict]],
    target_expiry: date,
    target_strike: float,
    calc_date: date,
) -> float:
    """VBA `mxFxOptionVolCalc`.

    The original read a jagged Excel range holding, for each tenor row,
    3 stacked rows (ATM vol / strike / premium) across an odd number of
    strike columns centered on the ATM column. That's reproduced here
    with an explicit, easier-to-audit input shape:

        tenor_dates: one date per tenor bucket, ascending.
        smile_table: one entry per tenor bucket, each a list of
            {"strike": float, "vol": float} dicts sorted ascending by
            strike (this replaces the ATM-centered vAtmVol/vStrk/vPrem
            grid reconstruction in the VBA -- the reconstruction itself
            carried no pricing information, only re-shaping).
        target_expiry / target_strike: point to interpolate to.
        calc_date: valuation date, for the time interpolation.

    Interpolates vol across strike within each of the two bracketing
    tenors, then interpolates across time (in the VBA: Act/365 year
    fractions from calc_date) between those two results.
    """
    if len(tenor_dates) != len(smile_table):
        raise ValueError("tenor_dates and smile_table must be the same length")

    n = len(tenor_dates)
    # bracket the expiry between two tenor buckets (VBA: Row_Idx(0), Row_Idx(1))
    row0 = None
    for i in range(n - 1):
        if tenor_dates[i] <= target_expiry < tenor_dates[i + 1]:
            row0 = i
            break
    if row0 is None:
        raise ValueError("target_expiry is outside the supplied tenor_dates range")
    row1 = row0 + 1

    def interp_strike(smile: Sequence[dict]) -> float:
        strikes = [pt["strike"] for pt in smile]
        vols = [pt["vol"] for pt in smile]
        for j in range(len(strikes) - 1):
            if strikes[j] <= target_strike < strikes[j + 1]:
                return my_linear_intp2(strikes[j], vols[j], strikes[j + 1], vols[j + 1], target_strike)
        raise ValueError("target_strike is outside the supplied smile strikes")

    vol0 = interp_strike(smile_table[row0])
    vol1 = interp_strike(smile_table[row1])

    t0 = (tenor_dates[row0] - calc_date).days / 365
    t1 = (tenor_dates[row1] - calc_date).days / 365
    t_target = (target_expiry - calc_date).days / 365

    return my_linear_intp2(t0, vol0, t1, vol1, t_target)


# ---------------------------------------------------------------------------
# Vanilla forward-price helper functions used inside the barrier formulas
# ---------------------------------------------------------------------------

def func_calc_d_one(f: float, sigma: float, te: float, z: float) -> float:
    return (math.log(f / z) + 0.5 * sigma * sigma * te) / (sigma * math.sqrt(te))


def func_calc_d_two(f: float, sigma: float, te: float, z: float) -> float:
    return func_calc_d_one(f, sigma, te, z) - sigma * math.sqrt(te)


# ---------------------------------------------------------------------------
# European barrier / rebate options (VBA: SimBarrEuroFunc, eurepeanbarrier)
# ---------------------------------------------------------------------------

def _normalize_rebate(reb_type: RebateStyle, rebate: float, strike: float) -> float:
    if reb_type in (RebateStyle.KRO, RebateStyle.USD, RebateStyle.QNT):
        return rebate / 100 * strike
    return rebate  # KRO_USD / USD_KRO / USD_QNT: already a flat amount


def sim_barr_euro_func(
    bar_type: BarrierType,
    o_type: str,
    f0: float,
    x: float,
    bar_lvl: float,
    reb_type: RebateStyle,
    rebate: float,
    sigma: float,
    te: float,
    td: float,
    rd: float,
    rf: float,
) -> float:
    """VBA `SimBarrEuroFunc`: European (knock-in/out observed only at expiry
    in this closed form) barrier option with rebate."""
    ot = o_type.upper()
    rebate = _normalize_rebate(reb_type, rebate, x)

    if bar_type == BarrierType.UP_AND_IN:
        if ot == "C":
            value = f0 * cnd(func_calc_d_one(f0, sigma, te, max(x, bar_lvl))) - x * cnd(
                func_calc_d_two(f0, sigma, te, max(x, bar_lvl))
            )
            value = math.exp(-rd * td) * value
            value += math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        elif ot == "P":
            if x < bar_lvl:
                return math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f0, sigma, te, bar_lvl))
            value = -math.exp(-rd * td) * (
                f0 * (cnd(-func_calc_d_one(f0, sigma, te, x)) - cnd(-func_calc_d_one(f0, sigma, te, bar_lvl)))
                - x * (cnd(-func_calc_d_two(f0, sigma, te, x)) - cnd(-func_calc_d_two(f0, sigma, te, bar_lvl)))
            )
            value += math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        raise ValueError("Error Otype")

    elif bar_type == BarrierType.UP_AND_OUT:
        if ot == "C":
            if x >= bar_lvl:
                return math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f0, sigma, te, bar_lvl))
            value = math.exp(-rd * td) * (
                f0 * (cnd(func_calc_d_one(f0, sigma, te, x)) - cnd(func_calc_d_one(f0, sigma, te, bar_lvl)))
                - x * (cnd(func_calc_d_two(f0, sigma, te, x)) - cnd(func_calc_d_two(f0, sigma, te, bar_lvl)))
            )
            value += math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        elif ot == "P":
            value = f0 * cnd(-func_calc_d_one(f0, sigma, te, min(x, bar_lvl))) - x * cnd(
                -func_calc_d_two(f0, sigma, te, min(x, bar_lvl))
            )
            value = -math.exp(-rd * td) * value
            value += math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        raise ValueError("Error Otype")

    elif bar_type == BarrierType.DOWN_AND_IN:
        if ot == "C":
            if x >= bar_lvl:
                return math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f0, sigma, te, bar_lvl))
            value = math.exp(-rd * td) * (
                f0 * (cnd(func_calc_d_one(f0, sigma, te, x)) - cnd(func_calc_d_one(f0, sigma, te, bar_lvl)))
                - x * (cnd(func_calc_d_two(f0, sigma, te, x)) - cnd(func_calc_d_two(f0, sigma, te, bar_lvl)))
            )
            value += math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        elif ot == "P":
            value = f0 * cnd(-func_calc_d_one(f0, sigma, te, min(x, bar_lvl))) - x * cnd(
                -func_calc_d_two(f0, sigma, te, min(x, bar_lvl))
            )
            value = -math.exp(-rd * td) * value
            value += math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        raise ValueError("Error Otype")

    elif bar_type == BarrierType.DOWN_AND_OUT:
        if ot == "C":
            value = f0 * cnd(func_calc_d_one(f0, sigma, te, max(x, bar_lvl))) - x * cnd(
                func_calc_d_two(f0, sigma, te, max(x, bar_lvl))
            )
            value = math.exp(-rd * td) * value
            value += math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        elif ot == "P":
            if x < bar_lvl:
                return math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f0, sigma, te, bar_lvl))
            value = -math.exp(-rd * td) * (
                f0 * (cnd(-func_calc_d_one(f0, sigma, te, x)) - cnd(-func_calc_d_one(f0, sigma, te, bar_lvl)))
                - x * (cnd(-func_calc_d_two(f0, sigma, te, x)) - cnd(-func_calc_d_two(f0, sigma, te, bar_lvl)))
            )
            value += math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f0, sigma, te, bar_lvl))
            return value
        raise ValueError("Error Otype")

    raise ValueError("Unknown BarrierType")


def europeanbarrier(
    f: float,
    k: float,
    sigma: float,
    rd: float,
    rf: float,
    te: float,
    td: float,
    win_type: str,
    bar_type: str,
    bar_lvl: float,
    rebate: float,
    reb_type: str,
    pre_type: str,
) -> float:
    """VBA `eurepeanbarrier` (name/typo preserved as a comment; the Python
    function itself is named without the typo). String-coded barrier type:
    'DOC'/'UIC' (down-out-call/up-in-call bucket), 'DOP'/'UIP',
    'UOC'/'DIC', 'UOP'/'DIP' -- exactly as grouped in the original
    `Select Case`. `pre_type` must be "KRO/USD" as in the source (the only
    branch implemented there); `reb_type` is "%KRO" or "KRO/USD"."""
    wt = win_type.upper()
    bt = bar_type.upper()
    pt = pre_type.upper()
    rt = reb_type.upper()

    if wt != "SINGLE":
        raise ValueError("Only WinType='Single' is implemented (matches source)")
    if pt != "KRO/USD":
        raise ValueError("Only PreType='KRO/USD' is implemented (matches source)")

    if rt == "%KRO":
        rebate = rebate / 100 * k
    elif rt == "KRO/USD":
        pass  # already a flat amount
    else:
        raise ValueError("Unknown RebType")

    if bt in ("DOC", "UIC"):
        d1 = func_calc_d_one(f, sigma, te, max(k, bar_lvl))
        d2 = func_calc_d_two(f, sigma, te, max(k, bar_lvl))
        d3 = func_calc_d_two(f, sigma, te, bar_lvl)
        value = math.exp(-rd * td) * (f * cnd(d1) - k * cnd(d2))
        value += math.exp(-rd * td) * rebate * cnd(-d3)
        return value

    elif bt in ("DOP", "UIP"):
        if k <= bar_lvl:
            return math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f, sigma, te, bar_lvl))
        value = -math.exp(-rd * td) * (
            f * (cnd(-func_calc_d_one(f, sigma, te, k)) - cnd(-func_calc_d_one(f, sigma, te, bar_lvl)))
            - k * (cnd(-func_calc_d_two(f, sigma, te, k)) - cnd(-func_calc_d_two(f, sigma, te, bar_lvl)))
        )
        value += math.exp(-rd * td) * rebate * cnd(-func_calc_d_two(f, sigma, te, bar_lvl))
        return value

    elif bt in ("UOC", "DIC"):
        if k >= bar_lvl:
            return math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f, sigma, te, bar_lvl))
        value = math.exp(-rd * td) * (
            f * (cnd(func_calc_d_one(f, sigma, te, k)) - cnd(func_calc_d_one(f, sigma, te, bar_lvl)))
            - k * (cnd(func_calc_d_two(f, sigma, te, k)) - cnd(func_calc_d_two(f, sigma, te, bar_lvl)))
        )
        value += math.exp(-rd * td) * rebate * cnd(func_calc_d_two(f, sigma, te, bar_lvl))
        return value

    elif bt in ("UOP", "DIP"):
        d1 = func_calc_d_one(f, sigma, te, min(k, bar_lvl))
        d2 = func_calc_d_two(f, sigma, te, min(k, bar_lvl))
        d3 = func_calc_d_two(f, sigma, te, bar_lvl)
        value = -math.exp(-rd * td) * (f * cnd(-d1) - k * cnd(-d2))
        value += math.exp(-rd * td) * rebate * cnd(d3)
        return value

    raise ValueError("Unknown BarType")


# ---------------------------------------------------------------------------
# American (continuously-monitored) barrier options: non-quanto and quanto
# (VBA: SimBarrAmeFunc, SimBarrAmeQntFunc)
# ---------------------------------------------------------------------------
#
# Both functions build a common set of 6 "var_i" reflection-principle terms
# and then pick a combination of them by (barrier type, strike-vs-barrier
# ordering, call/put). The VBA computes the terms twice inside a
# `For k = 0 To 1` loop (k=0 uses epsil=+1 for the call branch, k=1 uses
# epsil=-1 for the put branch) and only keeps the iteration matching the
# requested option type -- so functionally this is equivalent to just
# computing epsil from o_type once, which is what the Python version below
# does (it reproduces the same numeric result without the redundant loop).

def _barrier_terms(
    forward: float,
    spot_or_none: Optional[float],
    strike: float,
    bar_lvl: float,
    vol: float,
    te: float,
    td: float,
    o_type: str,
    rd_like: float,
    rf_like: float,
    use_forward_as_spot: bool,
    quanto_df: Optional[float] = None,
) -> dict:
    """Shared building blocks for SimBarrAmeFunc / SimBarrAmeQntFunc.

    `use_forward_as_spot=True` reproduces `SimBarrAmeQntFunc`, which prices
    off UnderInfo.Forward/Spot/Strike with the quanto discount factor
    (`QtoInfo.DF`) standing in for both rd and rf in the var_i terms.
    `use_forward_as_spot=False` reproduces `SimBarrAmeFunc`, which uses
    plain Spot/Forward/rd/rf.
    """
    spot = spot_or_none
    ot = o_type.upper()
    epsil = 1.0 if ot == "C" else -1.0

    if quanto_df is not None:
        # SimBarrAmeQntFunc: mu uses rd/rf plus the quanto convexity
        # adjustment (VolQtoNum/VolQtoDe) folded in by the caller via rd_like.
        mu = rd_like  # already includes the quanto adjustment when called from sim_barr_ame_qnt_func
        mu_hat = mu * td / te - 0.5 * vol * vol
        mu_apo = math.sqrt(mu_hat * mu_hat + 2 * quanto_df * vol * vol)
    else:
        mu = rd_like - rf_like
        mu_hat = mu * td / te - 0.5 * vol * vol
        mu_apo = math.sqrt(mu_hat * mu_hat + 2 * rd_like * vol * vol)

    eta = mu_hat / (vol * vol)
    eta_apo = mu_hat / (vol * vol) + 1

    x_ = math.log(spot / strike) / (vol * math.sqrt(te)) + eta_apo * vol * math.sqrt(te)
    x_apo = math.log(bar_lvl * bar_lvl / (spot * strike)) / (vol * math.sqrt(te)) + eta_apo * vol * math.sqrt(te)
    y_ = math.log(spot / bar_lvl) / (vol * math.sqrt(te)) + eta_apo * vol * math.sqrt(te)
    y_apo = math.log(bar_lvl / spot) / (vol * math.sqrt(te)) + eta_apo * vol * math.sqrt(te)
    z_ = math.log(bar_lvl / spot) / (vol * math.sqrt(te)) + mu_apo / vol * math.sqrt(te)

    epsil_apo = 1.0 if spot >= bar_lvl else -1.0

    if quanto_df is not None:
        df_fwd = math.exp(-quanto_df * td)
        df_strike = df_fwd
    else:
        df_fwd = math.exp(-rf_like * td)
        df_strike = math.exp(-rd_like * td)

    var = {}
    var[1] = epsil * spot * df_fwd * cnd(epsil * x_) - epsil * strike * df_strike * cnd(
        epsil * x_ - epsil * vol * math.sqrt(te)
    )
    var[2] = epsil * spot * df_fwd * cnd(epsil * y_) - epsil * strike * df_strike * cnd(
        epsil * y_ - epsil * vol * math.sqrt(te)
    )
    var[3] = epsil * spot * df_fwd * my_pow(bar_lvl / spot, 2 * eta_apo) * cnd(epsil_apo * x_apo) - epsil * strike * df_strike * my_pow(
        bar_lvl / spot, 2 * eta
    ) * cnd(epsil_apo * x_apo - epsil_apo * vol * math.sqrt(te))
    var[4] = epsil * spot * df_fwd * my_pow(bar_lvl / spot, 2 * eta_apo) * cnd(epsil_apo * y_apo) - epsil * strike * df_strike * my_pow(
        bar_lvl / spot, 2 * eta
    ) * cnd(epsil_apo * y_apo - epsil_apo * vol * math.sqrt(te))
    var[5] = 0.0  # filled in by caller (needs rebate, which differs by function)
    var[6] = 0.0  # filled in by caller

    return {
        "var": var,
        "mu": mu,
        "mu_hat": mu_hat,
        "mu_apo": mu_apo,
        "eta": eta,
        "eta_apo": eta_apo,
        "epsil": epsil,
        "epsil_apo": epsil_apo,
        "y": y_,
        "z": z_,
        "df_strike": df_strike,
    }


def sim_barr_ame_func(
    bar_type: BarrierType,
    o_type: str,
    forward: float,
    spot: float,
    strike: float,
    bar_lvl: float,
    reb_type: RebateStyle,
    rebate: float,
    reb_pmt: RebatePaymentTiming,
    vol: float,
    te: float,
    td: float,
    rd: float,
    rf: float,
) -> float:
    """VBA `SimBarrAmeFunc`: continuously-monitored American barrier option
    (non-quanto)."""
    ot = o_type.upper()
    rebate = _normalize_rebate(reb_type, rebate, strike)

    terms = _barrier_terms(forward, spot, strike, bar_lvl, vol, te, td, ot, rd, rf, use_forward_as_spot=False)
    var = terms["var"]
    eta, eta_apo = terms["eta"], terms["eta_apo"]
    mu_hat, mu_apo = terms["mu_hat"], terms["mu_apo"]
    epsil_apo = terms["epsil_apo"]
    y_, z_ = terms["y"], terms["z"]

    y_apo = math.log(bar_lvl / spot) / (vol * math.sqrt(te)) + eta_apo * vol * math.sqrt(te)
    var[5] = rebate * math.exp(-rd * td) * (
        cnd(epsil_apo * y_ - epsil_apo * vol * math.sqrt(te))
        - my_pow(bar_lvl / spot, 2 * eta) * cnd(epsil_apo * y_apo - epsil_apo * vol * math.sqrt(te))
    )
    var[6] = rebate * (
        my_pow(bar_lvl / spot, (mu_hat + mu_apo) / (vol * vol)) * cnd(epsil_apo * z_)
        + my_pow(bar_lvl / spot, (mu_hat - mu_apo) / (vol * vol))
        * cnd(epsil_apo * z_ - 2 * epsil_apo * (mu_apo / vol) * math.sqrt(te))
    )

    # Up/Down-and-Out "in the barrier already breached" rebate override,
    # and the immediate/at-maturity rebate cases -- these use mu_hat (not
    # mu_apo) in z, exactly as in the VBA `Select Case SimpBarrInfo.bartype`
    # block that runs before the main pricing Select Case.
    if bar_type in (BarrierType.UP_AND_OUT, BarrierType.DOWN_AND_OUT) and reb_pmt == RebatePaymentTiming.AT_MATURITY:
        z_mat = math.log(bar_lvl / spot) / (vol * math.sqrt(te)) + mu_hat / vol * math.sqrt(te)
        var[6] = rebate * math.exp(-rd * td) * (
            my_pow(bar_lvl / spot, (2 * mu_hat) / (vol * vol)) * cnd(epsil_apo * z_mat)
            + cnd(epsil_apo * z_mat - 2 * epsil_apo * (mu_hat / vol) * math.sqrt(te))
        )

    if bar_type == BarrierType.UP_AND_IN:
        if spot > bar_lvl:
            return fxd_bs_model(ot, forward, strike, vol, te, td, rd, rf) if False else (
                bs_model(ot, forward, strike, vol, te, rd, rf)  # spot already knocked in: plain vanilla
            )
        if strike > bar_lvl:
            return var[1] + var[5] if ot == "C" else var[1] - var[2] + var[4] + var[5]
        else:
            return var[2] - var[3] + var[4] + var[5] if ot == "C" else var[3] + var[5]

    elif bar_type == BarrierType.UP_AND_OUT:
        if spot > bar_lvl:
            return math.exp(-rd * td) * rebate if reb_pmt == RebatePaymentTiming.AT_MATURITY else rebate
        if strike > bar_lvl:
            return var[6] if ot == "C" else var[2] - var[4] + var[6]
        else:
            return (var[1] - var[2] + var[3] - var[4] + var[6]) if ot == "C" else (var[1] - var[3] + var[6])

    elif bar_type == BarrierType.DOWN_AND_IN:
        if spot <= bar_lvl:
            return bs_model(ot, forward, strike, vol, te, rd, rf)
        if strike > bar_lvl:
            return var[3] + var[5] if ot == "C" else var[2] - var[3] + var[4] + var[5]
        else:
            return (var[1] - var[2] + var[4] + var[5]) if ot == "C" else (var[1] + var[5])

    elif bar_type == BarrierType.DOWN_AND_OUT:
        if spot <= bar_lvl:
            return math.exp(-rd * td) * rebate if reb_pmt == RebatePaymentTiming.AT_MATURITY else rebate
        if strike > bar_lvl:
            return (var[1] - var[3] + var[6]) if ot == "C" else (var[1] - var[2] + var[3] - var[4] + var[6])
        else:
            return (var[2] - var[4] + var[6]) if ot == "C" else var[6]

    raise ValueError("Unknown BarrierType")


def sim_barr_ame_qnt_func(
    bar_type: BarrierType,
    o_type: str,
    forward: float,
    spot: float,
    strike: float,
    bar_lvl: float,
    reb_type: RebateStyle,
    rebate: float,
    reb_pmt: RebatePaymentTiming,
    vol: float,
    te: float,
    td: float,
    rd: float,
    rf: float,
    quanto: QuantoInfo,
) -> float:
    """VBA `SimBarrAmeQntFunc`: continuously-monitored American barrier
    option with the quanto convexity adjustment folded into the drift
    (`mu`) and the quanto discount factor (`quanto.df`) used in place of
    both rd and rf in the reflection-principle terms."""
    ot = o_type.upper()
    rebate = _normalize_rebate(
        RebateStyle.QNT if reb_type == RebateStyle.QNT else reb_type, rebate, strike
    )

    mu = rd - rf + 0.5 * (vol * vol + quanto.vol_qto_num ** 2 - quanto.vol_qto_de ** 2)

    terms = _barrier_terms(
        forward,
        spot,
        strike,
        bar_lvl,
        vol,
        te,
        td,
        ot,
        rd_like=mu,
        rf_like=rf,
        use_forward_as_spot=True,
        quanto_df=quanto.df,
    )
    var = terms["var"]
    eta, eta_apo = terms["eta"], terms["eta_apo"]
    mu_hat, mu_apo = terms["mu_hat"], terms["mu_apo"]
    epsil_apo = terms["epsil_apo"]
    y_, z_ = terms["y"], terms["z"]
    y_apo = math.log(bar_lvl / spot) / (vol * math.sqrt(te)) + eta_apo * vol * math.sqrt(te)

    var[5] = rebate * math.exp(-quanto.df * td) * (
        cnd(epsil_apo * y_ - epsil_apo * vol * math.sqrt(te))
        - my_pow(bar_lvl / spot, 2 * eta) * cnd(epsil_apo * y_apo - epsil_apo * vol * math.sqrt(te))
    )
    var[6] = rebate * (
        my_pow(bar_lvl / spot, (mu_hat + mu_apo) / (vol * vol)) * cnd(epsil_apo * z_)
        + my_pow(bar_lvl / spot, (mu_hat - mu_apo) / (vol * vol))
        * cnd(epsil_apo * z_ - 2 * epsil_apo * (mu_apo / vol) * math.sqrt(te))
    )

    if bar_type in (BarrierType.UP_AND_OUT, BarrierType.DOWN_AND_OUT) and reb_pmt == RebatePaymentTiming.AT_MATURITY:
        z_mat = math.log(bar_lvl / spot) / (vol * math.sqrt(te)) + mu_hat / vol * math.sqrt(te)
        var[6] = rebate * math.exp(-rd * td) * (
            my_pow(bar_lvl / spot, (2 * mu_hat) / (vol * vol)) * cnd(epsil_apo * z_mat)
            + cnd(epsil_apo * z_mat - 2 * epsil_apo * (mu_hat / vol) * math.sqrt(te))
        )

    if bar_type == BarrierType.UP_AND_IN:
        if spot > bar_lvl:
            return fxd_bs_model(ot, forward, strike, vol, te, td, quanto.df, quanto.df)
        if strike > bar_lvl:
            return var[1] + var[5] if ot == "C" else var[1] - var[2] + var[4] + var[5]
        else:
            return var[2] - var[3] + var[4] + var[5] if ot == "C" else var[3] + var[5]

    elif bar_type == BarrierType.UP_AND_OUT:
        if spot > bar_lvl:
            return math.exp(-quanto.df * td) * rebate if reb_pmt == RebatePaymentTiming.AT_MATURITY else rebate
        if strike > bar_lvl:
            return var[6] if ot == "C" else var[2] - var[4] + var[6]
        else:
            return (var[1] - var[2] + var[3] - var[4] + var[6]) if ot == "C" else (var[1] - var[3] + var[6])

    elif bar_type == BarrierType.DOWN_AND_IN:
        if spot <= bar_lvl:
            return fxd_bs_model(ot, forward, strike, vol, te, td, quanto.df, quanto.df)
        if strike > bar_lvl:
            return var[3] + var[5] if ot == "C" else var[2] - var[3] + var[4] + var[5]
        else:
            return (var[1] - var[2] + var[4] + var[5]) if ot == "C" else (var[1] + var[5])

    elif bar_type == BarrierType.DOWN_AND_OUT:
        if spot <= bar_lvl:
            return math.exp(-quanto.df * td) * rebate if reb_pmt == RebatePaymentTiming.AT_MATURITY else rebate
        if strike > bar_lvl:
            return (var[1] - var[3] + var[6]) if ot == "C" else (var[1] - var[2] + var[3] - var[4] + var[6])
        else:
            return (var[2] - var[4] + var[6]) if ot == "C" else var[6]

    raise ValueError("Unknown BarrierType")


# ---------------------------------------------------------------------------
# Double-barrier reflection-principle model
# (VBA: fn_sigma, fn_smallf, fn_DBoption_sig, fn_DBoption_f, LowP, HighP,
#  KnockProbFunc)
# ---------------------------------------------------------------------------

def fn_sigma(z1: float, z2: float, beta: float, gamma: float, nu: float) -> float:
    """VBA `fn_sigma` (identical to `fn_DBoption_sig` in the source --
    kept as one function; see `fn_dboption_sig` alias below)."""
    term1 = math.exp((nu / 2) * (nu * (gamma ** 2) - 2 * beta))
    term2 = cnd(gamma * nu - ((beta + z1) / gamma))
    term3 = cnd(gamma * nu - ((beta + z2) / gamma))
    return term1 * (term2 - term3)


# The VBA source defines fn_DBoption_sig as a byte-for-byte duplicate of
# fn_sigma; kept as an alias rather than a second copy of the same logic.
fn_dboption_sig = fn_sigma


def fn_smallf(
    t: float, mu: float, vol: float, x_l: float, x_h: float, nu: float, z1: float, z2: float, kappa: int
) -> float:
    """VBA `fn_smallf` (identical to `fn_DBoption_f` in the source)."""
    sum1 = 0.0
    for i in range(-kappa, kappa + 1):
        u = 2 * i * (x_h - x_l)
        alpha_1 = max(z1, x_l)
        alpha_2 = min(z2, x_h)
        sum1 += math.exp(-mu * u / (vol ** 2)) * fn_sigma(alpha_1, alpha_2, -mu * t + u, vol * math.sqrt(t), nu)

    sum2 = 0.0
    for i in range(-kappa, kappa + 1):
        u = 2 * i * (x_h - x_l)
        alpha_1 = max(z1, x_l)
        alpha_2 = max(z2, x_h)
        sum2 += math.exp(mu * (2 * x_h - u) / (vol ** 2)) * fn_sigma(
            alpha_1, alpha_2, -mu * t - 2 * x_h + u, vol * math.sqrt(t), nu
        )

    return (sum1 - sum2) * 100


# The VBA source defines fn_DBoption_f as a byte-for-byte duplicate of
# fn_smallf; kept as an alias.
fn_dboption_f = fn_smallf


def low_p(mu: float, vol: float, x_l: float, x_h: float, te: float, kappa: int) -> float:
    """VBA `LowP`: probability-weighted term for hitting the low barrier
    of a double-barrier structure (reflection principle, image-charge sum)."""
    term1 = math.exp(mu * x_l / (vol ** 2))
    sum1 = 0.0
    for i in range(0, int(kappa) + 1):
        mu_l = -x_l + 2 * i * (x_h - x_l)
        d1 = (-mu_l - mu * te) / (vol * math.sqrt(te))
        d2 = (-mu_l + mu * te) / (vol * math.sqrt(te))
        sum1 += math.exp(mu * mu_l / (vol ** 2)) * cnd(d1) + math.exp(-mu * mu_l / (vol ** 2)) * cnd(d2)

    sum2 = 0.0
    for i in range(-int(kappa), 0):
        mu_l = -x_l + 2 * i * (x_h - x_l)
        d1 = (-mu_l - mu * te) / (vol * math.sqrt(te))
        d2 = (-mu_l + mu * te) / (vol * math.sqrt(te))
        sum2 += math.exp(mu * mu_l / (vol ** 2)) * cnd(-d1) + math.exp(-mu * mu_l / (vol ** 2)) * cnd(-d2)

    return term1 * (sum1 - sum2)


def high_p(mu: float, vol: float, x_l: float, x_h: float, te: float, kappa: int) -> float:
    """VBA `HighP`: mirror of `low_p` for the high barrier."""
    term1 = math.exp(mu * x_h / (vol ** 2))
    sum1 = 0.0
    for i in range(0, int(kappa) + 1):
        mu_h = x_h + 2 * i * (x_h - x_l)
        d3 = (-mu_h - mu * te) / (vol * math.sqrt(te))
        d4 = (-mu_h + mu * te) / (vol * math.sqrt(te))
        sum1 += math.exp(mu * mu_h / (vol ** 2)) * cnd(d3) + math.exp(-mu * mu_h / (vol ** 2)) * cnd(d4)

    sum2 = 0.0
    for i in range(-int(kappa), 0):
        mu_h = x_h + 2 * i * (x_h - x_l)
        d3 = (-mu_h - mu * te) / (vol * math.sqrt(te))
        d4 = (-mu_h + mu * te) / (vol * math.sqrt(te))
        sum2 += math.exp(mu * mu_h / (vol ** 2)) * cnd(-d3) + math.exp(-mu * mu_h / (vol ** 2)) * cnd(-d4)

    return term1 * (sum1 - sum2)


def knock_prob_func(
    bar_type: str,
    prob_type: str,
    mu: float,
    vol: float,
    low_x: float,
    high_x: float,
    te: float,
    kappa: int,
    dstep: float,
) -> float:
    """VBA `KnockProbFunc`. bar_type: 'Double' or 'Single'. prob_type:
    'Low' or 'High'. `dstep` is kept for interface fidelity but the VBA
    `Step Dstep` inside a `For i = 0 To kappa` loop only behaves like a
    plain integer step when Dstep is an integer; non-integer steps are
    reproduced faithfully via a manual loop rather than `range()`."""
    bt = bar_type.upper()
    pt = prob_type.upper()

    if bt == "DOUBLE":
        sum1 = 0.0
        i = 0.0
        while i <= kappa:
            if pt == "LOW":
                low_u = -low_x + 2 * i * (high_x - low_x)
                d1 = (-low_u - mu * te) / (vol * math.sqrt(te))
                d2 = (-low_u + mu * te) / (vol * math.sqrt(te))
                sum1 += math.exp(mu * low_u / (vol ** 2)) * cnd(d1) + math.exp(-mu * low_u / (vol ** 2)) * cnd(d2)
            elif pt == "HIGH":
                high_u = high_x + 2 * i * (high_x - low_x)
                d3 = (-high_u - mu * te) / (vol * math.sqrt(te))
                d4 = (-high_u + mu * te) / (vol * math.sqrt(te))
                sum1 += math.exp(mu * high_u / (vol ** 2)) * cnd(d3) + math.exp(-mu * high_u / (vol ** 2)) * cnd(d4)
            i += dstep

        sum2 = 0.0
        i = -kappa
        while i <= -1:
            if pt == "LOW":
                low_u = -low_x + 2 * i * (high_x - low_x)
                d1 = (-low_u - mu * te) / (vol * math.sqrt(te))
                d2 = (-low_u + mu * te) / (vol * math.sqrt(te))
                sum2 += math.exp(mu * low_u / (vol ** 2)) * cnd(-d1) + math.exp(-mu * low_u / (vol ** 2)) * cnd(-d2)
            elif pt == "HIGH":
                high_u = high_x + 2 * i * (high_x - low_x)
                d3 = (-high_u - mu * te) / (vol * math.sqrt(te))
                d4 = (-high_u + mu * te) / (vol * math.sqrt(te))
                sum2 += math.exp(mu * high_u / (vol ** 2)) * cnd(-d3) + math.exp(-mu * high_u / (vol ** 2)) * cnd(-d4)
            i += dstep

        if pt == "LOW":
            return math.exp((mu / (vol ** 2)) * low_x) * (sum1 - sum2)
        elif pt == "HIGH":
            return math.exp((mu / (vol ** 2)) * high_x) * (sum1 - sum2)
        raise ValueError("Unknown ProbType")

    elif bt == "SINGLE":
        if pt == "LOW":
            z = (1 / (vol * math.sqrt(te))) * low_x + mu * math.sqrt(te) / vol
            eps = 1
            return math.exp((2 * mu / (vol ** 2)) * low_x) * cnd(eps * z) + cnd(eps * z - 2 * eps * (mu / vol) * math.sqrt(te))
        elif pt == "HIGH":
            z = (1 / (vol * math.sqrt(te))) * high_x + mu * math.sqrt(te) / vol
            eps = -1
            return math.exp((2 * mu / (vol ** 2)) * high_x) * cnd(eps * z) + cnd(eps * z - 2 * eps * (mu / vol) * math.sqrt(te))
        raise ValueError("Unknown ProbType")

    raise ValueError("Unknown BarType")


# ---------------------------------------------------------------------------
# Generic curve interpolation / discount factor (VBA: MyLinearIntp, myLinearDF)
# ---------------------------------------------------------------------------

def my_linear_intp(x_points: Sequence[float], y_points: Sequence[float], x: float) -> float:
    """VBA `MyLinearIntp`: linear interpolation with flat extrapolation
    logic replaced by linear extrapolation off the two end/near-end points,
    exactly as the original (uses the first two points below the range and
    the last two above it)."""
    n = len(x_points)
    if n < 2:
        raise ValueError("Need at least two points")

    if x < x_points[0]:
        i1, i2 = 0, 1
    elif x > x_points[-1]:
        i1, i2 = n - 2, n - 1
    else:
        i1 = i2 = None
        for ir in range(n):
            if x_points[ir] == x:
                return y_points[ir]
            elif x_points[ir] > x:
                i1, i2 = ir, ir - 1
                break
        if i1 is None:
            # x equals the last point (already handled) or numerical edge case
            i1, i2 = n - 2, n - 1

    return y_points[i1] + (y_points[i2] - y_points[i1]) * (x - x_points[i1]) / (x_points[i2] - x_points[i1])


def my_linear_df(x_points: Sequence[float], y_points: Sequence[float], x: float, calc_date: date, target_date: date) -> float:
    """VBA `myLinearDF`: builds a discount factor from a (date-as-number,
    zero-rate-in-%) curve via `my_linear_intp`, using ACT/365 to `target_date`.
    (In the VBA, `X` served double duty as both the curve x-axis value and
    `target_date - calcDate`; those are separated here into `x` -- the
    curve lookup key -- and `target_date` -- used only for the day count.)"""
    t = (target_date - calc_date).days / 365
    rate = my_linear_intp(x_points, y_points, x)
    return math.exp(-rate / 100 * t)


# ---------------------------------------------------------------------------
# Day-count conventions (VBA: MyDayCount, isleapyear, DateOfLeapYear)
# ---------------------------------------------------------------------------

def is_leap_year(d: date) -> bool:
    """VBA `isleapyear`: checks whether Feb 29 exists in `d`'s year."""
    try:
        date(d.year, 2, 29)
        return True
    except ValueError:
        return False


def date_of_leap_year(date1: date, date2: date) -> bool:
    """VBA `DateOfLeapYear`: walks year by year from date1 to date2 and
    returns True if any year in between (once its Mar-1 has been passed)
    is a leap year. Reproduced with the same year-stepping logic."""
    leap_year = False
    d1 = date1
    while True:
        if (date(d1.year, 3, 1) - d1).days < 0:  # DateDiff("d", date1, Mar1 of date1's year) > 0 check, VBA-style
            pass
        # VBA: If DateDiff("d", date1, DateSerial(Year(date2), 3, 1)) > 0
        if (date(date2.year, 3, 1) - d1).days > 0:
            if is_leap_year(date(d1.year, 1, 1)):
                leap_year = True
                break
        d1 = date(d1.year + 1, d1.month, min(d1.day, 28) if (d1.month == 2 and d1.day == 29) else d1.day)
        if d1 >= date2:
            break
    return leap_year


def my_day_count(start_date: date, end_date: date, dc_method: str) -> float:
    """VBA `MyDayCount`: ACT/360, ACT/365, 30I/360, 30E/360, ACT/365A, 'simple'."""
    method = dc_method.upper()
    y1, m1, d1 = start_date.year, start_date.month, start_date.day
    y2, m2, d2 = end_date.year, end_date.month, end_date.day

    if method == "ACT/360":
        return (end_date - start_date).days / 360

    elif method == "ACT/365":
        return (end_date - start_date).days / 365

    elif method == "30I/360":
        if d1 == 31:
            d1 = 30
        if d1 in (31, 30):
            if d2 == 31:
                d2 = 30
        if m1 == 2 and d1 in (28, 29):
            d1 = 30
        if m2 == 2 and d2 in (28, 29):
            d2 = 30
        return (y2 - y1) + (m2 - m1) / 12 + (d2 - d1) / 360

    elif method == "30E/360":
        if d1 == 31:
            d1 = 30
        if d2 == 31:
            d2 = 30
        if m1 == 2 and d1 in (28, 29):
            d1 = 30
        if m2 == 2 and d2 in (28, 29):
            d2 = 30
        return (y2 - y1) + (m2 - m1) / 12 + (d2 - d1) / 360

    elif method == "ACT/365A":
        if y1 == y2:
            denom = 366 if is_leap_year(start_date) else 365
            return (d2 - d1) / denom  # NOTE: reproduces the VBA literally --
            # this uses day-of-month d2-d1, not (end_date - start_date).days;
            # only meaningful when start/end share both year AND are being
            # read the way the original sheet fed it (kept as-is for fidelity)
        else:
            diy = 366 if date_of_leap_year(start_date, end_date) else 365
            return (end_date - start_date).days / diy

    elif method == "SIMPLE":
        return (end_date - start_date).days

    else:
        raise ValueError(f"Unknown day-count method: {dc_method}")


# ---------------------------------------------------------------------------
# Misc: normal RNG (VBA: MakeNormS, Box-Muller)
# ---------------------------------------------------------------------------

def make_norm_s() -> float:
    """VBA `MakeNormS`: Box-Muller normal draw. The VBA cached one of the
    two Box-Muller outputs across calls via `Static` variables; Python's
    `random.gauss(0, 1)` is used directly instead since there's no benefit
    to replicating that caching here."""
    return random.gauss(0, 1)


# ---------------------------------------------------------------------------
# Example: standard vs. quanto vanilla pricing (illustrates what
# `VanillaCalc_Click` did with worksheet cells, using explicit inputs
# instead of Excel Range reads/writes -- not a 1:1 port, see module docstring)
# ---------------------------------------------------------------------------

def price_vanilla_example(
    category: OptionCategory,
    o_type: str,
    deal: DealInfo,
    under: UnderInfo,
    quanto: Optional[QuantoInfo] = None,
) -> float:
    td = (deal.del_date - deal.spot_date).days / deal.dc
    te = (deal.mat_date - deal.val_date).days / deal.dc
    ot = o_type.upper()

    if category == OptionCategory.STANDARD:
        if ot == "C":
            return fxd_bs_model("C", under.forward, under.strike, under.vol, te, td, under.rd, under.rf)
        elif ot == "P":
            return fxd_bs_model("P", under.forward, under.strike, under.vol, te, td, under.rd, under.rf)
        raise ValueError("Error Otype")

    elif category == OptionCategory.QUANTO:
        if quanto is None:
            raise ValueError("quanto info required for OptionCategory.QUANTO")
        adj_forward = under.forward * math.exp(
            0.5 * (under.vol ** 2 + quanto.vol_qto_num ** 2 - quanto.vol_qto_de ** 2) * te
        )
        if ot == "C":
            return fxd_bs_model("C", adj_forward, under.strike, under.vol, te, td, quanto.df, quanto.df)
        elif ot == "P":
            return fxd_bs_model("P", adj_forward, under.strike, under.vol, te, td, quanto.df, quanto.df)
        raise ValueError("Error Otype")

    raise ValueError("Unknown OptionCategory")


# ---------------------------------------------------------------------------
# Smoke tests
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # 1. Sanity-check the CND approximation against scipy's exact CDF.
    if _scipy_norm is not None:
        max_err = max(abs(cnd(x / 10) - norm_cdf_exact(x / 10)) for x in range(-50, 51))
        print(f"max |CND_approx - exact| over x in [-5, 5]: {max_err:.2e}")

    # 2. Vanilla call/put parity check on fxd_bs_model.
    f0, x, sigma, te, td, rd = 1300.0, 1300.0, 0.10, 1.0, 1.0, 0.03
    call = fxd_bs_model("C", f0, x, sigma, te, td, rd, rd)
    put = fxd_bs_model("P", f0, x, sigma, te, td, rd, rd)
    parity_lhs = call - put
    parity_rhs = math.exp(-rd * td) * (f0 - x)
    print(f"call={call:.6f} put={put:.6f} parity_lhs={parity_lhs:.6f} parity_rhs={parity_rhs:.6f}")

    # 3. Up-and-in + up-and-out should reconstruct the vanilla (no rebate).
    bar_lvl = 1400.0
    ui = sim_barr_euro_func(
        BarrierType.UP_AND_IN, "C", f0, x, bar_lvl, RebateStyle.KRO_USD, 0.0, sigma, te, td, rd, rd
    )
    uo = sim_barr_euro_func(
        BarrierType.UP_AND_OUT, "C", f0, x, bar_lvl, RebateStyle.KRO_USD, 0.0, sigma, te, td, rd, rd
    )
    print(f"up-in={ui:.6f} up-out={uo:.6f} sum={ui + uo:.6f} vanilla={call:.6f}")

    # 4. Day count sanity check.
    print("ACT/360:", my_day_count(date(2026, 1, 1), date(2026, 7, 1), "ACT/360"))
    print("ACT/365:", my_day_count(date(2026, 1, 1), date(2026, 7, 1), "act/365"))

    # 5. mx_fx_option_pricer TEST

    calc_date = date(2026, 9, 17)  # ASSUMPTION: change if your curve date differs
    target_expiry = date(2027, 9, 17)

    tenor_dates = [
        date(2026, 9, 24), date(2026, 10, 19), date(2026, 11, 19),
        date(2026, 12, 17), date(2027, 3, 18), date(2027, 6, 17),
        date(2027, 9, 17), date(2028, 9, 19), date(2029, 9, 19),
        date(2031, 9, 18),
    ]

    # One row per tenor (same order as tenor_dates), 19 points each.
    # Vols are in percent; strikes ascending.
    vols_pct = [
        [10.696, 10.6959999962767, 10.6054105029142, 10.6054105029142, 10.5343274805392, 10.4740000005249,
         10.484218281573, 10.4937093946007, 10.5027319417692, 10.5114725547198, 10.52, 10.5912725441104,
         10.6653106109984, 10.7428994344936, 10.8259092235742, 10.9170000011041, 11.0802399688777, 11.2771957021899,
         11.5359999835115, 11.536],
        [10.3825, 10.3824999998595, 10.2782269895602, 10.2782269895602, 10.1964097838872, 10.1270000008493,
         10.1368207407597, 10.1459483346975, 10.1546271348912, 10.1630327087967, 10.171, 10.2452494294991,
         10.3245722716176, 10.4077732882005, 10.4969316558713, 10.5950000005809, 10.7639660725361, 10.9687304167145,
         11.2394999984743, 11.2395],
        [10.394, 10.3939999998608, 10.2835044351453, 10.2835044351453, 10.1966973253266, 10.1230000010006,
         10.1346812248568, 10.1455486101289, 10.1558884471819, 10.1659060256762, 10.175, 10.2523038413117,
         10.3385020812799, 10.4289684312851, 10.5260331744023, 10.6330000003287, 10.822608844664, 11.0537677280963,
         11.3619999969334, 11.362],
        [10.396, 10.3959999998896, 10.285439842495, 10.285439842495, 10.1984423844676, 10.1245000010264,
         10.1392907375893, 10.1530654738531, 10.1661822277853, 10.1788975766992, 10.19, 10.2742403484104,
         10.3720143066604, 10.4748017391113, 10.58533631753, 10.7075000001701, 10.9085349446675, 11.1544603304924,
         11.4839999954192, 11.484],
        [10.178, 10.1779999999285, 10.0638323075815, 10.0638323075815, 9.97373275593051, 9.89700000112962,
         9.9161126297464, 9.93395173804993, 9.9509682591844, 9.96748503264999, 9.98, 10.0658231017099, 10.1806567115373,
         10.3016030714755, 10.4320944357446, 10.577000000345, 10.8204974686502, 11.1222157976395, 11.5339999985013,
         11.534],
        [10.1725, 10.1724999987588, 10.0349101659961, 10.0349101659961, 9.92637221600135, 9.83400000004214,
         9.85518372367613, 9.87498873665464, 9.89390569200132, 9.91228570376426, 9.924, 10.0068841208437,
         10.1387232705462, 10.2777787680591, 10.4282299950425, 10.5960000003372, 10.8623264893039, 11.1947473914606,
         11.6534999969579, 11.6535],
        [10.2565, 10.2564999993837, 10.0992169265027, 10.0992169265027, 9.9751111819305, 9.86950000005083,
         9.88651588377638, 9.90243212919623, 9.91763880019034, 9.93241448743224, 9.94, 10.0152528224686,
         10.1636782650999, 10.3203719517112, 10.4903028806998, 10.6805000004918, 10.9859224908378, 11.3712110164623,
         11.9114999917561, 11.9115],
        [10.6175, 10.6174999993003, 10.4180018952471, 10.4180018952471, 10.2596988540581, 10.1245000000793,
         10.144085640507, 10.162495246034, 10.1801552930206, 10.1973704234104, 10.198, 10.2048852114755,
         10.3941903726668, 10.5929734701351, 10.808681116511, 11.0515000000836, 11.3867432286342, 11.8139113409064,
         12.424499998578, 12.4245],
        [10.7175, 10.7174999997026, 10.5231749390815, 10.5231749390815, 10.3674728830343, 10.2335000000534,
         10.2570498883019, 10.279288926298, 10.3007086895018, 10.3068543447509, 10.313, 10.3804494459494,
         10.4478988918988, 10.686816942319, 10.9454131053018, 11.2375000000072, 11.6067570956785, 12.0819042136184,
         12.7744999976606, 12.7745],
        [10.95525, 10.9552499986011, 10.7571752172066, 10.7571752172066, 10.5960349560487, 10.4557500000267,
         10.4867082045771, 10.5161899074438, 10.5234599382959, 10.5307299691479, 10.538, 10.6220063701795,
         10.706012740359, 10.7900191105385, 11.1413357443751, 11.5382499999903, 11.9404201227062, 12.4611720158183,
         13.2367499979417, 13.23675],
    ]

    strikes = [
        [1348.77481454877, 1356.03406418576, 1361.12482946773, 1361.12482946773, 1365.11940951241, 1368.50957132264,
         1371.46078302378, 1374.20197618277, 1376.80783963655, 1379.33227574383, 1381.79514544072, 1384.36351737212,
         1387.03154747016, 1389.82753236088, 1392.81886675153, 1396.10140700802, 1399.9246258324, 1404.53749654929,
         1410.59891284136, 1418.83578731594],
        [1313.85956553038, 1328.54271392346, 1338.94472062688, 1338.94472062688, 1347.10659414929, 1354.03073517457,
         1360.033195856, 1365.61200343219, 1370.91650718238, 1376.05401436181, 1380.92364208951, 1386.19392375347,
         1391.82432015071, 1397.72999235691, 1404.05852192937, 1411.01948851776, 1419.14813299905, 1428.99896599178,
         1442.02518859968, 1459.68843497204],
        [1286.97637404925, 1307.1951840123, 1321.59535997507, 1321.59535997507, 1332.90837203513, 1342.51287010717,
         1350.825410326, 1358.55880959171, 1365.91679739384, 1373.04546102662, 1379.51687366151, 1386.72756136471,
         1394.76789388157, 1403.20634385336, 1412.26027361589, 1422.23784216193, 1433.96695582995, 1448.26633531647,
         1467.33343806697, 1492.98142564269],
        [1268.3045337222, 1292.27296028616, 1309.37947130889, 1309.37947130889, 1322.84022434838, 1334.2810202661,
         1344.17929795353, 1353.39764569308, 1362.17565760171, 1370.68504069792, 1378.11501952138, 1386.5700631377,
         1396.38344944419, 1406.70002862485, 1417.79417411781, 1430.05551068791, 1444.47487349534, 1462.11403255115,
         1485.75048043288, 1517.4413443244],
        [1225.8751079998, 1258.02619330362, 1281.13738517269, 1281.13738517269, 1299.37641569626, 1314.90957908416,
         1328.32901828796, 1340.85428709849, 1352.80199758956, 1364.39882317182, 1373.18588402412, 1383.66490103653,
         1397.68610659776, 1412.45368010804, 1428.38670045926, 1446.07969739519, 1467.14564440559, 1493.24851263987,
         1528.8736228944, 1575.73785115572],
        [1192.55505787705, 1230.93668638717, 1258.88768482474, 1258.88768482474, 1280.93687268844, 1299.70203587026,
         1315.81334896128, 1330.87608153736, 1345.26340053836, 1359.24234628461, 1368.15167200469, 1379.16787923401,
         1396.69074584919, 1415.17274410634, 1435.1693595917, 1457.46783034168, 1484.09625940495, 1517.33308254851,
         1563.20107435238, 1623.10168728283],
        [1163.13250064928, 1206.88813522364, 1239.08884720659, 1239.08884720659, 1264.49713344615, 1286.11901040718,
         1304.64225240638, 1321.96844417751, 1338.52220361602, 1354.60680045091, 1362.86427839025, 1373.39590502197,
         1394.16803413502, 1416.09730334783, 1439.8791235983, 1466.49719565294, 1498.56838612066, 1539.02598855959,
         1595.75956641309, 1668.75825938289],
        [1073.59298987438, 1133.44206112816, 1178.5601240409, 1178.5601240409, 1214.36159997744, 1244.93788253658,
         1271.04158504141, 1295.57787045952, 1319.11514575439, 1342.05943709933, 1342.89853551842, 1344.08789539115,
         1376.78870105609, 1411.12673705513, 1448.38834038486, 1490.3331720691, 1540.63017828578, 1604.71880221237,
         1696.3262647807, 1815.06533960791],
        [1015.05736575258, 1085.66510461871, 1139.20244798581, 1139.20244798581, 1182.09899674242, 1219.00907007575,
         1250.78544315244, 1280.79305757684, 1309.69520682427, 1317.98767222378, 1326.28013762329, 1339.06100604565,
         1351.841874468, 1397.11400566443, 1446.11490533217, 1501.46190064954, 1567.75402157317, 1653.05640456926,
         1777.39699004522, 1938.92209481426],
        [925.941289819563, 1012.52899342289, 1079.09502391135, 1079.09502391135, 1133.24864943423, 1180.39353377483,
         1221.49469162228, 1260.63559592946, 1270.28753439369, 1279.93947285793, 1289.59141132216, 1306.64443982672,
         1323.69746833129, 1340.75049683585, 1412.06666676606, 1492.63901027299, 1588.10246204474, 1711.71376588203,
         1895.81334740217, 2140.00886424158],
    ]

    # Build smile_table: one list of {"strike", "vol"} dicts per tenor (vol as decimal)
    smile_table = [
        [{"strike": k, "vol": v / 100} for k, v in zip(k_row, v_row)]
        for k_row, v_row in zip(strikes, vols_pct)
    ]
    o_type = "p"
    s0 = 1382.20
    f0 = 1369.6137
    x = 1357.1996
    sigma = mx_fx_option_vol_calc(tenor_dates, smile_table, target_expiry, x, calc_date)
    print(f"sigma={sigma:.12f}")
    te = my_day_count(calc_date, target_expiry, "act/365")
    td = my_day_count(date(2026, 9, 21), date(2027, 9, 21), "act/365")
    rf = 0.044259353360886
    rd = 0.0351116727930189

    # 1
    print(f"\n#.1 Premium: %USD, Delta: %USD")
    opt_type = "c"
    prem_type = ccy_to_prem_type("%USD")
    greeks_quo_type = greeks_quote_type("%USD")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type , greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")
    # 2
    print(f"\n#.2 Premium: %USD, Delta: %KRW")
    opt_type = "c"
    prem_type = ccy_to_prem_type("%USD")
    greeks_quo_type = greeks_quote_type("%KRW")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type , greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")
    # 3
    print(f"\n#.3 Premium: %KRW, Delta: %KRW")
    opt_type = "c"
    prem_type = ccy_to_prem_type("%KRW")
    greeks_quo_type = greeks_quote_type("%KRW")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type , greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")
    # 4
    print(f"\n#.4 Premium: %KRW, Delta: %USD")
    opt_type = "c"
    prem_type = ccy_to_prem_type("%KRW")
    greeks_quo_type = greeks_quote_type("%USD")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf,rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")
    # 5
    print(f"\n#.5 Premium: USD/KRW, Delta: %USD")
    opt_type = "c"
    prem_type = ccy_to_prem_type("USD/KRW")
    greeks_quo_type = greeks_quote_type("%USD")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf,rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")
    # 6
    print("\n#.6 Premium: KRW/USD, Delta: %USD")
    opt_type = "c"
    prem_type = ccy_to_prem_type("KRW/USD")
    greeks_quo_type = greeks_quote_type("%USD")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf,rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")
    # 7
    print(f"\n#.7 Premium: USD/KRW, Delta: %KRW")
    opt_type = "c"
    prem_type = ccy_to_prem_type("USD/KRW")
    greeks_quo_type = greeks_quote_type("%KRW")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf,rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")
    # 8
    print(f"\n #.8 Premium: KRW/USD, Delta: %KRW")
    opt_type = "c"
    prem_type = ccy_to_prem_type("KRW/USD")
    greeks_quo_type = greeks_quote_type("%KRW")
    price = mx_fx_option_pricer(opt_type, prem_type, s0, f0, x, sigma, te, td, rf, rd)
    print(f"price={price:.12f}")
    delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"delta={delta:.12f}")
    forward_delta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Forward Delta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"forward_delta={forward_delta:.12f}")
    gamma = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Gamma"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"gamma={gamma:.12f}")
    vega = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vega"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vega={vega:.12f}")
    theta = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Theta"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"theta={theta:.12f}")
    rho_rd = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rd"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rd={rho_rd:.12f}")
    rho_rf = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Rho rf"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"rho_rf={rho_rf:.12f}")
    vanna = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Vanna"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"vanna={vanna:.12f}")
    volga = mx_fx_option_grk_calc(opt_type, prem_type, greeks_quo_type, greeks_type("Volga"), s0, f0, x, sigma, te, td, rf, rd)
    print(f"volga={volga:.12f}")