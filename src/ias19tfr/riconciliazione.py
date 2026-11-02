"""Movimentazione del DBO nel periodo e utili/perdite attuariali.

    DBO atteso = DBO apertura + SC + NIE + VF
    P/G        = DBO atteso - DBO chiusura        (positivo = utile, come nella Tabella 7 GHC)

Regole (confermate da Francesco il 02/11/2026, verificate sul caso Garofalo 30/09/2026):
- SC  = accantonamento contabile di periodo fornito dal cliente (foglio `codici`);
- NIE = DBO apertura x forward dell'anno 1 della valutazione di apertura x giorni/365
        (gli utilizzi del periodo non entrano nel calcolo);
- VF  = variazioni del fondo fornite dal cliente (utilizzi negativi) piu eventuali trasferimenti
        in entrata/uscita (es. acquisizione di una societa nel periodo: Citta di Roma 2026).

Scomposizione degli utili/perdite, sulla popolazione di chiusura:
    ipotesi finanziarie : DBO(demografia vecchia, finanza nuova) - DBO(ipotesi tutte vecchie)
    ipotesi demografiche: DBO(tutte nuove) - DBO(demografia vecchia, finanza nuova)
    esperienza          : residuo
Le variazioni sono espresse come utili (positivo = riduzione della passivita).
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date
from typing import Dict, Mapping, Optional, Sequence

from .engine import valuta
from .model import Dipendente, Ipotesi


def frazione_anno(inizio: date, fine: date) -> float:
    return (fine - inizio).days / 365.0


@dataclass
class Movimentazione:
    dbo_apertura: float
    service_cost: float
    nie: float
    variazioni_fondo: float   # utilizzi + trasferimenti
    dbo_chiusura: float
    utile_esperienza: float = 0.0
    utile_demografiche: float = 0.0
    utile_finanziarie: float = 0.0

    @property
    def dbo_atteso(self) -> float:
        return self.dbo_apertura + self.service_cost + self.nie + self.variazioni_fondo

    @property
    def utile(self) -> float:
        return self.dbo_atteso - self.dbo_chiusura


def movimentazione(dip_chiusura: Sequence[Dipendente], ip_apertura: Ipotesi, ip_chiusura: Ipotesi,
                   dbo_apertura: Mapping[int, float], service_cost: Mapping[int, float],
                   variazioni_fondo: Mapping[int, float], inizio: date, fine: date,
                   metodo: str = "financial",
                   trasferimenti: Optional[Mapping[int, float]] = None) -> Dict[int, Movimentazione]:
    """Movimentazione per azienda (chiave 0 = totale)."""
    fr = frazione_anno(inizio, fine)
    tasso = float(ip_apertura.finanziarie.attualizzazione[0])
    a = valuta(dip_chiusura, ip_apertura)
    b = valuta(dip_chiusura, replace(ip_apertura, finanziarie=ip_chiusura.finanziarie))
    c = valuta(dip_chiusura, ip_chiusura)
    trasferimenti = trasferimenti or {}
    chiavi = (set(c) | set(dbo_apertura) | set(service_cost) | set(variazioni_fondo) | set(trasferimenti)) - {0}
    out: Dict[int, Movimentazione] = {}
    for k in chiavi:
        dbo_c = c[k].dbo_metodo(metodo) if k in c else 0.0
        fin = (b[k].dbo_metodo(metodo) - a[k].dbo_metodo(metodo)) if k in c else 0.0
        demo = (dbo_c - b[k].dbo_metodo(metodo)) if k in c else 0.0
        d0 = dbo_apertura.get(k, 0.0)
        m = Movimentazione(d0, service_cost.get(k, 0.0), d0 * tasso * fr,
                           variazioni_fondo.get(k, 0.0) + trasferimenti.get(k, 0.0), dbo_c,
                           utile_finanziarie=-fin, utile_demografiche=-demo)
        m.utile_esperienza = m.utile - m.utile_finanziarie - m.utile_demografiche
        out[k] = m
    tot = Movimentazione(*(sum(getattr(m, f) for m in out.values()) for f in
                           ("dbo_apertura", "service_cost", "nie", "variazioni_fondo", "dbo_chiusura",
                            "utile_esperienza", "utile_demografiche", "utile_finanziarie")))
    out[0] = tot
    return out
