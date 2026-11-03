"""Riconciliazione con il foglio di riferimento Garofalo al 30.09.2026.

Il file contiene dati di clienti e non e nel repository: impostare
IAS19_GAROFALO_XLSX con il percorso del foglio per eseguire questi test.
"""
import os

import numpy as np
import openpyxl
import pytest

from ias19tfr.engine import valuta
from ias19tfr.io_excel import dati_da_foglio_storico, ipotesi_da_excel
from ias19tfr.scenari import griglia_sensitivity, scomposizione_csl

PATH = os.environ.get("IAS19_GAROFALO_XLSX")
pytestmark = pytest.mark.skipif(not PATH or not os.path.exists(PATH), reason="foglio Garofalo non disponibile")


@pytest.fixture(scope="module")
def caso():
    ip = ipotesi_da_excel(PATH)
    dv = dati_da_foglio_storico(PATH)
    res = valuta(dv.dipendenti, ip)
    wb = openpyxl.load_workbook(PATH, data_only=True, read_only=True)
    return ip, dv, res, wb


def _pilota(wb):
    """Righe di Pilota: (azienda, griglia 3x3 di sensitivity) dalla colonna T in poi."""
    ws = wb["Pilota"]
    out = {}
    for r in ws.iter_rows(min_row=11, max_row=44, min_col=20, max_col=29, values_only=True):
        if isinstance(r[0], (int, float)):
            out[int(r[0])] = np.array(r[1:10], dtype=float).reshape(3, 3)
    return out


def test_totali_gruppo(caso):
    _, _, res, _ = caso
    t = res[0]
    assert t.n == 3244
    assert t.ccp == pytest.approx(14_743_550.00, abs=0.01)
    assert t.psl == pytest.approx(13_855_968.13, abs=0.01)
    assert t.csl == pytest.approx(-503_161.08, abs=0.01)
    assert t.dbo == pytest.approx(13_352_807.06, abs=0.01)


def test_dbo_per_azienda_e_sensitivity(caso):
    ip, dv, res, wb = caso
    ref = _pilota(wb)
    griglia = griglia_sensitivity(dv.dipendenti, ip)
    for k, g in ref.items():
        if k == 0:
            continue
        if k not in res:
            assert np.allclose(g, 0)
            continue
        assert res[k].dbo == pytest.approx(g[1, 1], abs=0.02), k
        assert np.allclose(griglia[k], g, atol=0.02), k
    assert np.allclose(griglia[0], ref[0], atol=0.05)


def test_scomposizione_csl(caso):
    ip, dv, res, wb = caso
    dec = scomposizione_csl(dv.dipendenti, ip, res)
    soc = wb["Soc01"]
    ra, inf, sal = [soc.cell(17, c).value for c in (2, 3, 4)]
    assert dec[0]["DC-R/A"] == pytest.approx(ra, abs=1)
    assert dec[0]["DC-Inf"] == pytest.approx(inf, abs=1)
    assert dec[0]["DC-Sal"] == pytest.approx(sal, abs=1)
