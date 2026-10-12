"""
FX Option Pricer — Black-76 Model

Prices European FX options using the Black-76 (Black-1976) forward-price model:
C = DF_dom * [F*N(d1) - K*N(d2)]
P = DF_dom * [K*N(-d2) - F*N(-d1)]

where the forward F is typically the market FX forward rate (already embeds
the interest-rate differential), so only the domestic discount factor is
needed for discounting the expected payoff.

If you prefer to price directly off spot with domestic/foreign rates
(Garman-Kohlhagen), see the `from_spot` classmethod which first builds the
forward internally.
"""

from dataclasses import dataclass
from math import log, sqrt, exp
from statistics import NormalDist

_N = NormalDist()


@dataclass
class FXOptionResult:
    price: float
    delta_fwd: float # dPrice/dForward
    gamma_fwd: float # d2Price/dForward2
    vega: float # dPrice/dVol (per 1.00 = 100 vol pts)
    theta: float # dPrice/dT (per year, T decreasing)
    rho_dom: float # dPrice/d(domestic rate), per 1.00 (100%)


class Black76FXOption:
    """
    Black-76 pricer for a European FX option quoted on the forward.

    Parameters
    ----------
    forward : float
    FX forward rate for the option's expiry/delivery (domestic per foreign, e.g. KRW per USD).
    strike : float
    Strike rate, same quoting convention as forward.
    vol : float
    Annualized Black volatility (e.g. 0.10 for 10%).
    T : float
    Time to expiry in years (ACT/365 or day-count of choice).
    r_dom : float
    Domestic risk-free rate (continuously compounded), used only for discounting.
    option_type : str
    'call' or 'put'. A call is the right to buy the base/foreign currency
    (i.e. right to receive foreign ccy, pay domestic ccy) at the strike.
    """

    def __init__(self, forward: float, strike: float, vol: float, T: float,
        r_dom: float, option_type: str = "call"):
        if forward <= 0 or strike <= 0:
            raise ValueError("forward and strike must be positive")
        if vol < 0:
            raise ValueError("vol must be non-negative")
        if T < 0:
            raise ValueError("T must be non-negative")
        if option_type not in ("call", "put"):
            raise ValueError("option_type must be 'call' or 'put'")

        self.F = forward
        self.K = strike
        self.vol = vol
        self.T = T
        self.r_dom = r_dom
        self.option_type = option_type


    @classmethod
    def from_spot(cls, spot: float, strike: float, vol: float, T: float, r_dom: float, r_for: float, option_type: str = "call"):
        """
        Build the pricer from spot + domestic/foreign rates (Garman-Kohlhagen
        style inputs), computing the forward as F = S * exp((r_dom - r_for) * T)
        internally, then applying Black-76 on that forward.
        """
        forward = spot * exp((r_dom - r_for) * T)
        return cls(forward, strike, vol, T, r_dom, option_type)

# -- internals -----------------------------------------------------

    def _d1_d2(self):
        if self.T == 0 or self.vol == 0:
            return None, None
        vol_sqrt_T = self.vol * sqrt(self.T)
        d1 = (log(self.F / self.K) + 0.5 * self.vol ** 2 * self.T) / vol_sqrt_T
        d2 = d1 - vol_sqrt_T
        return d1, d2

    def _discount_factor(self):
        return exp(-self.r_dom * self.T)

# -- public API ------------------------------------------------------

    def price(self) -> float:
        df = self._discount_factor()
        d1, d2 = self._d1_d2()

        if d1 is None:
            # zero vol or zero time: payoff is deterministic
            intrinsic = max(self.F - self.K, 0.0) if self.option_type == "call" \
        else max(self.K - self.F, 0.0)
            return df * intrinsic

        if self.option_type == "call":
            return df * (self.F * _N.cdf(d1) - self.K * _N.cdf(d2))
        else:
            return df * (self.K * _N.cdf(-d2) - self.F * _N.cdf(-d1))

    def greeks(self) -> FXOptionResult:
        df = self._discount_factor()
        d1, d2 = self._d1_d2()
        price = self.price()

        if d1 is None:
            return FXOptionResult(price, 0.0, 0.0, 0.0, 0.0, 0.0)

        pdf_d1 = _N.pdf(d1)
        sign = 1.0 if self.option_type == "call" else -1.0

        delta_fwd = df * sign * _N.cdf(sign * d1)
        gamma_fwd = df * pdf_d1 / (self.F * self.vol * sqrt(self.T))
        vega = df * self.F * pdf_d1 * sqrt(self.T) # per 1.00 vol (divide by 100 for per-vol-point)

        # Theta: d(price)/dT with T decreasing (time decay), holding vol/rate fixed
        term1 = -df * self.F * pdf_d1 * self.vol / (2 * sqrt(self.T))
        if self.option_type == "call":
            term2 = -self.r_dom * df * (self.F * _N.cdf(d1) - self.K * _N.cdf(d2))
        else:
            term2 = -self.r_dom * df * (self.K * _N.cdf(-d2) - self.F * _N.cdf(-d1))
        theta = term1 + term2

        rho_dom = -self.T * price # dPrice/dr_dom, holding forward fixed

        return FXOptionResult(price, delta_fwd, gamma_fwd, vega, theta, rho_dom)


if __name__ == "__main__":
    # Example: USD/KRW call, forward = 1380.00, strike = 1400.00,
    # vol = 9%, 3 months to expiry, domestic (KRW) rate = 3.0%
    opt = Black76FXOption(
    forward=1357.1996,
    strike=1357.1996,
    vol=0.0994,
    T=1,
    r_dom=0.03584,
    option_type="call",
    )

    price = opt.price()
    g = opt.greeks()

    print(f"Price: {price:.4f}")
    print(f"Delta (fwd): {g.delta_fwd:.6f}")
    print(f"Gamma (fwd): {g.gamma_fwd:.8f}")
    print(f"Vega (per 100%): {g.vega:.4f} (per 1 vol pt: {g.vega/100:.6f})")
    print(f"Theta (per yr): {g.theta:.4f} (per day: {g.theta/365:.6f})")
    print(f"Rho (domestic): {g.rho_dom:.4f}")

    # Example: pricing directly from spot + rates (Garman-Kohlhagen inputs)
    opt2 = Black76FXOption.from_spot(
    spot=1368.60,
    strike=1357.1996,
    vol=0.0994,
    T=1,
    r_dom=0.03584,
    r_for=0.0439, # USD rate as the foreign currency
    option_type="call",
    )
    print(f"From spot, implied forward: {opt2.F:.4f}")
    print(f"Price: {opt2.price():.4f}")
