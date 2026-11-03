import math

import pytest

from ias19tfr.curva import Svensson


def test_curva_piatta():
    """Con beta1 = beta2 = beta3 = 0 la curva e piatta: forward = spot = e^beta0 - 1."""
    c = Svensson(3.0, 0.0, 0.0, 0.0, 1.5, 10.0)
    fwd = c.forward_annui(50)
    assert fwd == pytest.approx([math.exp(0.03) - 1] * 50)


def test_forward_costante_oltre_scadenza_massima():
    c = Svensson(2.5, -1.0, 1.0, 2.0, 1.5, 9.0)
    fwd = c.forward_annui(50, scadenza_max=30)
    assert all(fwd[29:] == fwd[29])
    # coerenza spot-forward
    prod = 1.0
    for t in range(30):
        prod *= 1 + fwd[t]
    assert prod == pytest.approx((1 + c.spot_annuo(30)) ** 30)
