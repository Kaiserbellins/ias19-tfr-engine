"""Movimentazione 31/12/2025 -> 30/09/2026 del caso Garofalo (Tabella 7 della relazione GHC).

Servono tre file, indicati da variabili d'ambiente (i dati dei clienti non sono nel repository):
    IAS19_GAROFALO_XLSX       valutazione al 30/09/2026
    IAS19_GAROFALO_PREV_XLSX  valutazione al 31/12/2025 (convertita da .xls in .xlsx)
    IAS19_SCHEMA_XLSX         schema dati cliente al 30/09/2026
"""
import os
from datetime import date

import pytest

from ias19tfr.engine import valuta
from ias19tfr.io_excel import dati_da_foglio_storico, dati_da_schema_cliente, ipotesi_da_excel
from ias19tfr.model import ALIQUOTA_TFR
from ias19tfr.riconciliazione import movimentazione

CUR = os.environ.get("IAS19_GAROFALO_XLSX")
PREV = os.environ.get("IAS19_GAROFALO_PREV_XLSX")
SCHEMA = os.environ.get("IAS19_SCHEMA_XLSX")
pytestmark = pytest.mark.skipif(not all(p and os.path.exists(p) for p in (CUR, PREV, SCHEMA)),
                                reason="file Garofalo non disponibili")

EUTONIA, CITTA_DI_ROMA = 27, 32
TRASFERIMENTO_CITTA_DI_ROMA = 921_968.00  # TFR acquisito nel periodo


@pytest.fixture(scope="module")
def dati():
    cur = dati_da_foglio_storico(CUR)
    grandi = {k for k in {e.azienda for e in cur.dipendenti}
              if all(e.fip == pytest.approx(ALIQUOTA_TFR) for e in cur.dipendenti if e.azienda == k)}
    sch = dati_da_schema_cliente(SCHEMA, aziende_almeno_50=grandi, aziende_escluse={EUTONIA})
    return cur, sch


def test_schema_cliente_come_input_tfr(dati):
    cur, sch = dati
    assert len(sch.dipendenti) == len(cur.dipendenti) == 3244
    for a, b in zip(sch.dipendenti, cur.dipendenti):
        assert (a.azienda, a.eta, a.sesso, a.anzianita, a.futuro, a.carriera, a.salario, a.tfr) == \
               (b.azienda, b.eta, b.sesso, b.anzianita, b.futuro, b.carriera, b.salario, b.tfr)
        assert a.fip == pytest.approx(b.fip, abs=1e-9)


def test_movimentazione_tabella_7(dati):
    cur, sch = dati
    ip_old, ip_new = ipotesi_da_excel(PREV), ipotesi_da_excel(CUR)
    prev = dati_da_foglio_storico(PREV)
    apertura = {k: r.dbo for k, r in valuta(prev.dipendenti, ip_old).items() if k}
    m = movimentazione(
        sch.dipendenti, ip_old, ip_new, apertura,
        service_cost={k: a.accantonamento for k, a in sch.aziende.items()},
        variazioni_fondo={k: a.utilizzi for k, a in sch.aziende.items()},
        trasferimenti={CITTA_DI_ROMA: TRASFERIMENTO_CITTA_DI_ROMA},
        inizio=date(2025, 12, 31), fine=date(2026, 9, 30))
    t = m[0]
    assert t.dbo_apertura == pytest.approx(14_468_054.84, abs=0.01)
    assert t.service_cost == pytest.approx(433_966.67, abs=0.01)
    assert t.nie == pytest.approx(224_572.20, abs=0.01)
    assert t.variazioni_fondo == pytest.approx(-852_212.07, abs=0.01)
    assert t.dbo_chiusura == pytest.approx(13_352_807.06, abs=0.01)
    assert t.utile == pytest.approx(921_574.59, abs=0.02)
    assert t.utile_esperienza + t.utile_demografiche + t.utile_finanziarie == pytest.approx(t.utile)
    assert m[1].nie == pytest.approx(11_809.31, abs=0.01)
    assert m[1].utile == pytest.approx(16_030.00, abs=0.01)
    assert m[EUTONIA].utile == pytest.approx(466_561.30, abs=0.01)
    assert m[CITTA_DI_ROMA].utile == pytest.approx(262_629.93, abs=0.01)
