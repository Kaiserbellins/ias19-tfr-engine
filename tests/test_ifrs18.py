from datetime import date

from ias19tfr.ifrs18 import Periodo, colonne
from ias19tfr.riconciliazione import Movimentazione


def _m(sc, nie, vf, chiusura, fin, esp):
    return {0: Movimentazione(1000.0, sc, nie, vf, chiusura, utile_esperienza=esp, utile_finanziarie=fin)}


def test_trimestri_per_differenza():
    p = [Periodo(date(2026, 3, 31), _m(10, 5, -100, 900, 10, 5)),
         Periodo(date(2026, 6, 30), _m(25, 10, -150, 860, 10, 15))]
    (e1, q1), (e2, q2), (e3, ytd) = colonne(p, 0)
    assert (e1, e2, e3) == ("1° trim. 2026", "2° trim. 2026", "Progressivo 6 mesi 2026")
    assert q2["dbo_apertura"] == 900 and q2["dbo_chiusura"] == 860
    assert (q2["service_cost"], q2["nie"], q2["variazioni_fondo"]) == (15, 5, -50)
    assert q2["perdite_finanziarie"] == 0 and q2["perdite_esperienza"] == -10
    for v in (q1, q2, ytd):
        assert abs(v["dbo_apertura"] + v["ce"] + v["oci"] + v["variazioni_fondo"] - v["dbo_chiusura"]) < 1e-9
