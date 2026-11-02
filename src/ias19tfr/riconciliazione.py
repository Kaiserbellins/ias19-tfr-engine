"""Movimentazione del DBO nel periodo e utili/perdite attuariali.

    P/G = DBO apertura + SC + NIE + VF (+/- curtailment) - DBO chiusura   (positivo = utile)

La scomposizione degli utili/perdite usa tre valutazioni sulla popolazione di chiusura:
    finanziarie : DBO(ipotesi demo vecchie, fin nuove) - DBO(ipotesi tutte vecchie)
    demografiche: DBO(tutte nuove) - DBO(demo vecchie, fin nuove)
    esperienza  : residuo
Le perdite sono positive (incrementano la passivita), come nella Tabella 6.1 della relazione BE.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Dict, Optional, Sequence

from .engine import valuta
from .model import Dipendente, Ipotesi


@dataclass
class Movimenti:
    dbo_apertura: float
    service_cost: float      # di norma l'accantonamento contabile di periodo
    utilizzi: float          # benefits paid, negativo
    frazione_anno: float     # es. 0.75 per il 30/09
    tasso_nie: float         # tasso annuo per l'interest cost
    curtailment: float = 0.0
    trasferimenti: float = 0.0

    def nie(self, utilizzi_meta_periodo: bool = True) -> float:
        base = self.dbo_apertura + (self.utilizzi / 2 if utilizzi_meta_periodo else 0.0)
        return base * self.tasso_nie * self.frazione_anno

    def dbo_atteso(self, nie: Optional[float] = None) -> float:
        nie = self.nie() if nie is None else nie
        return (self.dbo_apertura + self.service_cost + nie + self.utilizzi
                + self.trasferimenti - self.curtailment)


def tasso_piatto_equivalente(forward, flussi) -> float:
    """Tasso unico che restituisce lo stesso valore attuale dei flussi scontati con la curva."""
    import numpy as np

    flussi = np.asarray(flussi, dtype=float)
    fwd = np.asarray(forward[: len(flussi)], dtype=float)
    v = np.cumprod(1.0 / (1.0 + fwd))
    pv = (flussi * v).sum()
    lo, hi = -0.05, 0.20
    t = np.arange(1, len(flussi) + 1)
    for _ in range(100):
        mid = (lo + hi) / 2
        if (flussi / (1 + mid) ** t).sum() > pv:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def scomposizione_pg(dip: Sequence[Dipendente], ip_vecchie: Ipotesi, ip_nuove: Ipotesi,
                     dbo_atteso: Dict[int, float], metodo: str = "financial") -> Dict[int, Dict[str, float]]:
    a = valuta(dip, ip_vecchie)
    b = valuta(dip, replace(ip_vecchie, finanziarie=ip_nuove.finanziarie))
    c = valuta(dip, ip_nuove)
    out = {}
    for k in c:
        fin = b[k].dbo_metodo(metodo) - a[k].dbo_metodo(metodo)
        demo = c[k].dbo_metodo(metodo) - b[k].dbo_metodo(metodo)
        tot = c[k].dbo_metodo(metodo) - dbo_atteso.get(k, 0.0)
        out[k] = {
            "esperienza": tot - fin - demo,
            "ipotesi demografiche": demo,
            "ipotesi finanziarie": fin,
            "totale (perdita +)": tot,
        }
    return out
