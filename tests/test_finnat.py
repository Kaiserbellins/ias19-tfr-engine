"""Gruppo Banca Finnat: valutazioni 31/03, 30/06, 30/09/2026 e movimentazione dal 31/12/2025.

Riferimento: prospetto delle poste calcolato da Francesco (poste_Finnat_IAS19_rev_30092026.xlsx).
Serve la cartella IAS19_FINNAT_DIR con i fogli storici convertiti in .xlsx:
f2025_12.xlsx, f2026_03.xlsx, f2026_06.xlsx, f2026_09.xlsx.
"""
import os
from datetime import date

import pytest

from ias19tfr.engine import valuta
from ias19tfr.io_excel import dati_da_foglio_storico, ipotesi_da_excel
from ias19tfr.riconciliazione import movimentazione

DIR = os.environ.get("IAS19_FINNAT_DIR")
pytestmark = pytest.mark.skipif(not (DIR and os.path.isdir(DIR)), reason="file Finnat non disponibili")

# SC e VF per azienda (1 Banca Finnat, 3 InvestiRE, 4 Fiduciaria); DBO e P/G attesi per azienda
CASI = {
    "2026_03": (date(2026, 3, 31), {1: (0, 0), 3: (48951.18, -17501.09), 4: (13275.89, -225840.88)},
                {1: (1468195.30, 13782.14), 3: (4038737.15, 75090.42), 4: (404079.66, -1354.24)}),
    "2026_06": (date(2026, 6, 30), {1: (0, -39231.46), 3: (93532.06, -1052493.37), 4: (24990.28, -232655.646018)},
                {1: (1465657.27, -15282.61), 3: (3124354.58, 20076.18), 4: (415071.96, -4279.62)}),
    "2026_09": (date(2026, 9, 30), {1: (0, -39231.46), 3: (100608.98, -1592957.53), 4: (36967.86, -235632.17)},
                {1: (1405118.92, 52968.26), 3: (2405120.64, 227168.41), 4: (406356.47, 16639.02)}),
}


@pytest.mark.parametrize("periodo", sorted(CASI))
def test_poste_finnat(periodo):
    fine, poste, attesi = CASI[periodo]
    prev = os.path.join(DIR, "f2025_12.xlsx")
    cur = os.path.join(DIR, f"f{periodo}.xlsx")
    ip_old, ip_new = ipotesi_da_excel(prev), ipotesi_da_excel(cur)
    dprev, dcur = dati_da_foglio_storico(prev), dati_da_foglio_storico(cur)
    # aziende attribuite per intervalli di righe di Pilota (codici 1-2-3 nei dati, aziende 1-3-4)
    assert {e.azienda for e in dcur.dipendenti} == {1, 3, 4}
    apertura = {k: r.dbo for k, r in valuta(dprev.dipendenti, ip_old).items() if k}
    assert sum(apertura.values()) == pytest.approx(6_148_184.33, abs=0.01)
    m = movimentazione(dcur.dipendenti, ip_old, ip_new, apertura,
                       {k: v[0] for k, v in poste.items()}, {k: v[1] for k, v in poste.items()},
                       date(2025, 12, 31), fine)
    for k, (dbo, pg) in attesi.items():
        assert m[k].dbo_chiusura == pytest.approx(dbo, abs=0.02)
        assert m[k].utile == pytest.approx(pg, abs=0.02)
