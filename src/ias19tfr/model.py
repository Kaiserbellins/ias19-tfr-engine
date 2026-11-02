"""Strutture dati del motore: dipendenti, basi tecniche, risultati."""
from __future__ import annotations

from dataclasses import dataclass, field, replace
from typing import Dict, List, Optional

import numpy as np

# Aliquota annua di accantonamento TFR al netto del contributo 0,5%: 1/13,5 - 0,5%
# (valore usato nei fogli di riferimento, arrotondato a 6 decimali).
ALIQUOTA_TFR = 0.069074

# Imposta sostitutiva sulla rivalutazione dal 01/01/2015.
IMPOSTA_RIVALUTAZIONE = 0.17

N_ANNI = 50  # orizzonte massimo di proiezione


@dataclass(frozen=True)
class Dipendente:
    codice: int
    azienda: int
    eta: int            # X(k,0), eta intera
    sesso: int          # Y(k): 1 = M, 2 = F
    anzianita: int      # M(k,0), anni di servizio pregresso
    futuro: int         # T(k), anni di servizio futuro (uscita certa all'anno T)
    carriera: int       # R(k) = p, categoria (colonna di tavola e di incremento)
    salario: float      # S(k,0), salario annuo lordo
    tfr: float          # W(k,0), TFR contabile
    fip: float          # u(k), quota di salario che esce dall'azienda (0 .. ALIQUOTA_TFR)


@dataclass
class TavolaPermanenza:
    """l_x per sesso e categoria: lx[(sesso, p)] = array indicizzato per eta 0..111."""
    lx: Dict[tuple, np.ndarray]

    def sopravvivenza(self, sesso: int, p: int, eta: int) -> float:
        """Probabilita di restare attivo da eta a eta+1."""
        col = self.lx[(sesso, p)]
        den = col[eta]
        return float(col[eta + 1] / den) if den > 0 else 0.0

    @classmethod
    def costruisci(cls, lx_base: Dict[int, np.ndarray], dimissioni: Dict[tuple, np.ndarray],
                   anticipazioni: Dict[tuple, np.ndarray], radice: float = 100000.0) -> "TavolaPermanenza":
        """Combina la tavola di mortalita (es. RG48) con dimissioni e anticipazioni.

        l_x = l_{x-1} * (L_x / L_{x-1}) * (1 - d_x) * (1 - a_x), come nel foglio `Tav`.
        """
        out = {}
        for (sesso, p), d in dimissioni.items():
            base = lx_base[sesso]
            a = anticipazioni[(sesso, p)]
            n = len(base)
            col = np.zeros(n)
            col[0] = radice
            for x in range(1, n):
                ratio = base[x] / base[x - 1] if base[x - 1] > 0 else 0.0
                col[x] = col[x - 1] * ratio * (1 - d[x]) * (1 - a[x])
            out[(sesso, p)] = col
        return cls(out)


@dataclass
class IpotesiFinanziarie:
    inflazione: np.ndarray                 # pi_t, t = 1..N_ANNI (indice 0 = anno 1)
    attualizzazione: np.ndarray            # i_t forward annui
    incremento_reale: Dict[int, np.ndarray]  # g_p,t per categoria
    imposta: float = IMPOSTA_RIVALUTAZIONE
    rivalutazione: Optional[np.ndarray] = None  # se None: (1,5% + 75% pi) * (1 - imposta)

    def tasso_rivalutazione(self) -> np.ndarray:
        if self.rivalutazione is not None:
            return self.rivalutazione
        return (0.015 + 0.75 * self.inflazione) * (1 - self.imposta)

    def scala(self, attualizzazione: float = 1.0, inflazione: float = 1.0,
              shift_attualizzazione: float = 0.0, shift_inflazione: float = 0.0) -> "IpotesiFinanziarie":
        """Ipotesi modificate per la sensitivity (variazioni relative e/o assolute)."""
        return replace(
            self,
            attualizzazione=self.attualizzazione * attualizzazione + shift_attualizzazione,
            inflazione=self.inflazione * inflazione + shift_inflazione,
            # se cambia l'inflazione la rivalutazione va ricalcolata dalla regola di legge
            rivalutazione=(self.rivalutazione if (inflazione == 1.0 and shift_inflazione == 0.0)
                           else None),
        )


@dataclass
class Ipotesi:
    tavola: TavolaPermanenza
    finanziarie: IpotesiFinanziarie
    aliquota_tfr: float = ALIQUOTA_TFR


@dataclass
class Flussi:
    """Flussi per anno di uscita t = 1..T di un dipendente (equivalenti ai fogli W1, W2, C1, C2)."""
    w1: np.ndarray   # TFR contabile pagato in t
    w2: np.ndarray   # TFR contabile rivalutato e attualizzato
    c1: np.ndarray   # contributi futuri attualizzati all'epoca di maturazione
    c2: np.ndarray   # TFR futuro rivalutato fino a t e attualizzato
    pagamenti: np.ndarray  # erogazioni attese nominali (TFR rivalutato + quote future rivalutate)
    pv_pagamenti: np.ndarray
    dbo_lineare: np.ndarray  # contributo per anno alla DBO con linear proration
    sc_fin: float     # service cost pro futuro (financial proration)
    sc_lin: float     # service cost pro futuro (linear proration)


@dataclass
class Risultato:
    """Risultati aggregati (per azienda o gruppo)."""
    n: int = 0
    ccp: float = 0.0
    psl: float = 0.0
    c1: float = 0.0
    c2: float = 0.0
    dbo_lineare: float = 0.0
    sc_fin: float = 0.0
    sc_lin: float = 0.0
    per_anno: Dict[str, np.ndarray] = field(default_factory=lambda: {
        k: np.zeros(N_ANNI) for k in ("w1", "w2", "c1", "c2", "pagamenti", "pv_pagamenti")})

    @property
    def dw(self) -> float:
        return self.psl - self.ccp

    @property
    def csl(self) -> float:
        return self.c2 - self.c1

    @property
    def dbo(self) -> float:
        return self.psl + self.csl

    @property
    def duration(self) -> float:
        pv = self.per_anno["pv_pagamenti"]
        tot = pv.sum()
        return float((np.arange(1, N_ANNI + 1) * pv).sum() / tot) if tot else 0.0

    def aggiungi(self, f: Flussi, tfr: float) -> None:
        t = len(f.w1)
        self.n += 1
        self.ccp += tfr
        self.psl += f.w2.sum()
        self.c1 += f.c1.sum()
        self.c2 += f.c2.sum()
        self.dbo_lineare += f.dbo_lineare.sum()
        self.sc_fin += f.sc_fin
        self.sc_lin += f.sc_lin
        for k in self.per_anno:
            self.per_anno[k][:t] += getattr(f, k)

    def unisci(self, altro: "Risultato") -> None:
        for a in ("n", "ccp", "psl", "c1", "c2", "dbo_lineare", "sc_fin", "sc_lin"):
            setattr(self, a, getattr(self, a) + getattr(altro, a))
        for k in self.per_anno:
            self.per_anno[k] += altro.per_anno[k]

    def dbo_metodo(self, metodo: str) -> float:
        return self.dbo if metodo == "financial" else self.dbo_lineare
