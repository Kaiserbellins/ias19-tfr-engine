"""Proprieta del modello verificate su dati sintetici."""
from dataclasses import replace

import numpy as np
import pytest

from ias19tfr.engine import proietta, valuta
from ias19tfr.model import ALIQUOTA_TFR, Dipendente, Ipotesi, IpotesiFinanziarie, TavolaPermanenza


def tavola(d=0.05, a=0.02):
    eta = np.arange(112)
    lx = 100000 * np.exp(-0.0001 * eta ** 2)
    lx[-1] = 0
    base = {1: lx, 2: lx}
    dim = {(s, p): np.full(112, d) for s in (1, 2) for p in (1, 2, 3)}
    ant = {(s, p): np.full(112, a) for s in (1, 2) for p in (1, 2, 3)}
    return TavolaPermanenza.costruisci(base, dim, ant)


def ipotesi(att=0.035, infl=0.02):
    fin = IpotesiFinanziarie(inflazione=np.full(50, infl), attualizzazione=np.full(50, att),
                             incremento_reale={p: np.full(50, 0.01) for p in (1, 2, 3)})
    return Ipotesi(tavola(), fin)


def dip(fip=0.0, futuro=20, tfr=30000.0):
    return Dipendente(1, 1, 45, 1, 15, futuro, 3, 40000.0, tfr, fip)


def test_equivalenza_finanziaria():
    """Se attualizzazione = rivalutazione il DBO coincide con il TFR contabile (anche con TFR futuro)."""
    ip = ipotesi()
    r = ip.finanziarie.tasso_rivalutazione()
    ip = replace(ip, finanziarie=replace(ip.finanziarie, attualizzazione=r.copy()))
    for fip in (0.0, ALIQUOTA_TFR):
        res = valuta([dip(fip)], ip)[0]
        assert res.dbo == pytest.approx(30000.0, rel=1e-12)


def test_stock_pagato_per_intero():
    f = proietta(dip(), ipotesi())
    assert f.w1.sum() == pytest.approx(30000.0, rel=1e-12)
    assert len(f.w1) == 20


def test_senza_tfr_futuro_metodi_coincidono():
    res = valuta([dip(ALIQUOTA_TFR)], ipotesi())[0]
    assert res.csl == 0
    assert res.dbo_lineare == pytest.approx(res.dbo, rel=1e-12)


def test_tasso_piu_alto_dbo_piu_basso():
    a = valuta([dip()], ipotesi(att=0.03))[0].dbo
    b = valuta([dip()], ipotesi(att=0.04))[0].dbo
    assert b < a


def test_turnover_piu_alto_riduce_dbo_con_attualizzazione_sopra_rivalutazione():
    from ias19tfr.scenari import sensitivity_assoluta
    sa = sensitivity_assoluta([dip(ALIQUOTA_TFR)], ipotesi())
    su = sa["Tasso di turnover +1.00%"][0]
    giu = sa["Tasso di turnover -1.00%"][0]
    base = valuta([dip(ALIQUOTA_TFR)], ipotesi())[0].dbo
    # uscite anticipate = meno anni di sconto netto (i > r): DBO piu vicino al TFR contabile
    assert giu < base < su
