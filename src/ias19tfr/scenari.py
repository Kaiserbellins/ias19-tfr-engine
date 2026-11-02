"""Valutazioni derivate: scomposizione del CSL, sensitivity, curtailment."""
from __future__ import annotations

from dataclasses import replace
from typing import Dict, Iterable, List, Sequence, Tuple

import numpy as np

from .engine import valuta
from .model import Dipendente, Ipotesi, Risultato


def _senza_crescita(ip: Ipotesi, inflazione_salari: bool) -> Ipotesi:
    """Ipotesi con salari costanti (o solo indicizzati all'inflazione), rivalutazione invariata."""
    fin = ip.finanziarie
    zero = {p: np.zeros_like(g) for p, g in fin.incremento_reale.items()}
    # l'inflazione resta per la rivalutazione; per i salari si azzera con un'incidenza "fittizia"
    nuovo = replace(fin, incremento_reale=zero, rivalutazione=fin.tasso_rivalutazione())
    if not inflazione_salari:
        # salario nominale costante: (1+pi)(1+g) = 1  =>  g = 1/(1+pi) - 1
        nuovo = replace(nuovo, incremento_reale={p: 1.0 / (1.0 + fin.inflazione) - 1.0 for p in zero})
    return replace(ip, finanziarie=nuovo)


def scomposizione_csl(dip: Sequence[Dipendente], ip: Ipotesi,
                      base: Dict[int, Risultato] | None = None) -> Dict[int, Dict[str, float]]:
    """CSL = DC-R/A + DC-Inf + DC-Sal (per azienda; chiave 0 = totale).

    DC-R/A: CSL a salari costanti (effetto rivalutazione/attualizzazione)
    DC-Inf: incremento dovuto alla sola inflazione dei salari
    DC-Sal: incremento dovuto alla crescita reale (residuo)
    """
    base = base or valuta(dip, ip)
    ra = valuta(dip, _senza_crescita(ip, inflazione_salari=False))
    inf = valuta(dip, _senza_crescita(ip, inflazione_salari=True))
    out = {}
    for k, r in base.items():
        out[k] = {
            "DC-R/A": ra[k].csl,
            "DC-Inf": inf[k].csl - ra[k].csl,
            "DC-Sal": r.csl - inf[k].csl,
            "CSL": r.csl,
        }
    return out


GRIGLIA_GHC = (0.9, 1.0, 1.1)


def griglia_sensitivity(dip: Sequence[Dipendente], ip: Ipotesi, metodo: str = "financial",
                        fattori_att: Sequence[float] = GRIGLIA_GHC,
                        fattori_infl: Sequence[float] = GRIGLIA_GHC) -> Dict[int, np.ndarray]:
    """Griglia DBO: righe = fattori sull'inflazione, colonne = fattori sull'attualizzazione.

    Variazioni relative applicate a tutta la curva (come nel foglio `Pilota`).
    """
    out: Dict[int, np.ndarray] = {}
    for i, fi in enumerate(fattori_infl):
        for j, fa in enumerate(fattori_att):
            ipx = replace(ip, finanziarie=ip.finanziarie.scala(attualizzazione=fa, inflazione=fi))
            res = valuta(dip, ipx)
            for k, r in res.items():
                out.setdefault(k, np.zeros((len(fattori_infl), len(fattori_att))))[i, j] = r.dbo_metodo(metodo)
    return out


def sensitivity_assoluta(dip: Sequence[Dipendente], ip: Ipotesi, metodo: str = "financial",
                         d_att: float = 0.0025, d_infl: float = 0.0025, d_turnover: float = 0.01
                         ) -> Dict[str, Dict[int, float]]:
    """Variazioni assolute, nello stile della relazione BE."""
    scen: List[Tuple[str, Ipotesi]] = [
        (f"Tasso di attualizzazione +{d_att:.2%}", replace(ip, finanziarie=ip.finanziarie.scala(shift_attualizzazione=d_att))),
        (f"Tasso di attualizzazione -{d_att:.2%}", replace(ip, finanziarie=ip.finanziarie.scala(shift_attualizzazione=-d_att))),
        (f"Tasso di inflazione +{d_infl:.2%}", replace(ip, finanziarie=ip.finanziarie.scala(shift_inflazione=d_infl))),
        (f"Tasso di inflazione -{d_infl:.2%}", replace(ip, finanziarie=ip.finanziarie.scala(shift_inflazione=-d_infl))),
    ]
    out = {}
    for nome, ipx in scen:
        out[nome] = {k: r.dbo_metodo(metodo) for k, r in valuta(dip, ipx).items()}
    # turnover: richiede la ricostruzione della tavola, gestita da chi possiede le basi grezze
    return out


def curtailment(dip: Sequence[Dipendente], ip: Ipotesi, coinvolti: Iterable[int],
                metodo: str = "financial", nuova_fip: float | None = None) -> Dict[int, float]:
    """Effetto di riduzione: DBO prima meno DBO dopo l'evento, a parita di ipotesi.

    `coinvolti`: codici dei dipendenti interessati. Se `nuova_fip` e None i dipendenti escono
    dal piano (DBO dopo = 0); altrimenti cambia la loro aliquota FIP (es. passaggio a >= 50 dip.).
    """
    cod = set(coinvolti)
    prima = valuta([e for e in dip if e.codice in cod], ip)
    if nuova_fip is None:
        return {k: r.dbo_metodo(metodo) for k, r in prima.items()}
    dopo = valuta([replace(e, fip=nuova_fip) for e in dip if e.codice in cod], ip)
    return {k: prima[k].dbo_metodo(metodo) - dopo[k].dbo_metodo(metodo) for k in prima}
