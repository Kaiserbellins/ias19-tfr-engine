"""Proiezione individuale e valutazione IAS 19 del TFR.

Metodo principale: Financial proration (Annibali, Bellini 2004)
    DBO = W~ + (C~ - C)
    W~ = TFR contabile rivalutato e attualizzato
    C~ = TFR futuro rivalutato e attualizzato
    C  = contributi futuri attualizzati all'epoca di maturazione

Metodo alternativo: linear proration (PUCM classico)
    DBO = sum_t P_t v_t B_t * n / (n + t), con B_t beneficio atteso all'uscita in t.

Convenzioni (riprodotte dal foglio di riferimento):
- anni interi dalla data di valutazione, pagamenti e accantonamenti a fine anno;
- all'anno T(k) l'intero stock residuo viene pagato (uscita certa);
- le anticipazioni sono un'uscita finanziaria: ogni anno riducono lo stock della frazione a_x.
"""
from __future__ import annotations

from typing import Dict, Iterable, List

import numpy as np

from .model import Dipendente, Flussi, Ipotesi, Risultato


def proietta(e: Dipendente, ip: Ipotesi) -> Flussi:
    fin = ip.finanziarie
    n = max(int(e.futuro), 1)
    riv = fin.tasso_rivalutazione()
    infl = fin.inflazione
    att = fin.attualizzazione
    g = fin.incremento_reale.get(e.carriera)
    if g is None:
        raise ValueError(f"Dipendente {e.codice}: categoria di carriera {e.carriera} senza incrementi salariali")
    quota = e.salario * max(ip.aliquota_tfr - e.fip, 0.0)

    w1 = np.zeros(n); w2 = np.zeros(n); c1 = np.zeros(n); c2 = np.zeros(n)
    pag = np.zeros(n); pv = np.zeros(n); lin = np.zeros(n)
    stock = 1.0
    R = 1.0; V = 1.0; crescita = 1.0
    cum_b = 0.0       # sum_j a_j / R_j
    cum_c = 0.0       # sum_j a_j v_j
    a1_su_R1 = 0.0
    sc_fin = 0.0; sc_lin = 0.0
    n_pass = max(e.anzianita, 0)
    for t in range(1, n + 1):
        s = ip.tavola.sopravvivenza(e.sesso, e.carriera, e.eta + t - 1)
        P = stock * (1.0 - s) if t < n else stock
        stock *= s
        R *= 1.0 + riv[t - 1]
        V /= 1.0 + att[t - 1]
        crescita *= (1.0 + infl[t - 1]) * (1.0 + g[t - 1])
        a = quota * crescita
        cum_b += a / R
        cum_c += a * V
        if t == 1:
            a1_su_R1 = a / R
        w1[t - 1] = e.tfr * P
        w2[t - 1] = e.tfr * P * R * V
        c1[t - 1] = P * cum_c
        c2[t - 1] = P * R * V * cum_b
        beneficio = e.tfr * R + R * cum_b           # nominale all'uscita
        pag[t - 1] = P * beneficio
        pv[t - 1] = P * beneficio * V
        # linear proration: pro-rata solo se ci sono quote future in azienda
        if quota > 0:
            prorata = n_pass / (n_pass + t) if (n_pass + t) > 0 else 1.0
            lin[t - 1] = P * V * beneficio * prorata
            sc_lin += P * V * beneficio / (n_pass + t)
        else:
            lin[t - 1] = P * V * e.tfr * R
        sc_fin += P * R * V * a1_su_R1
    return Flussi(w1, w2, c1, c2, pag, pv, lin, sc_fin, sc_lin)


def valuta(dipendenti: Iterable[Dipendente], ip: Ipotesi) -> Dict[int, Risultato]:
    """Valuta tutti i dipendenti; restituisce i risultati per azienda (chiave 0 = totale)."""
    out: Dict[int, Risultato] = {}
    tot = Risultato()
    for e in dipendenti:
        f = proietta(e, ip)
        out.setdefault(e.azienda, Risultato()).aggiungi(f, e.tfr)
    for r in out.values():
        tot.unisci(r)
    out[0] = tot
    return out


def flussi_individuali(dipendenti: Iterable[Dipendente], ip: Ipotesi) -> List[Flussi]:
    return [proietta(e, ip) for e in dipendenti]
