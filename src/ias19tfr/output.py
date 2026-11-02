"""Excel dei risultati (versione di base: il formato definitivo arriva con il passo 3)."""
from __future__ import annotations

from typing import Dict

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font

from .io_excel import DatiValutazione
from .model import Risultato

EUR = '#,##0.00'


def scrivi_risultati(path: str, dv: DatiValutazione, res: Dict[int, Risultato],
                     csl: Dict[int, Dict[str, float]], sens: Dict[int, np.ndarray], metodo: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Riepilogo"
    hdr = ["Azienda", "Nome", "Dipendenti", "CCP", "DW", "PSL", "DC-R/A", "DC-Inf", "DC-Sal", "CSL",
           "DBO financial", "DBO lineare", "SC pro futuro", "Duration"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
    for k in sorted(res, key=lambda x: (x == 0, x)):
        r = res[k]
        nome = "Totale" if k == 0 else (dv.aziende[k].nome if k in dv.aziende else str(k))
        d = csl[k]
        sc = r.sc_fin if metodo == "financial" else r.sc_lin
        ws.append([k or "", nome, r.n, r.ccp, r.dw, r.psl, d["DC-R/A"], d["DC-Inf"], d["DC-Sal"], r.csl,
                   r.dbo, r.dbo_lineare, sc, round(r.duration, 2)])
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=13):
        for c in row:
            c.number_format = EUR

    wa = wb.create_sheet("Per anno")
    wa.append(["Azienda", "Anno", "CCP pagato", "PSL", "C", "C~", "Erogazioni attese", "Valore attuale"])
    for k in sorted(res, key=lambda x: (x == 0, x)):
        pa = res[k].per_anno
        for t in range(len(pa["w1"])):
            if pa["w1"][t] == 0 and pa["c1"][t] == 0:
                continue
            wa.append([k or "Totale", t + 1, pa["w1"][t], pa["w2"][t], pa["c1"][t], pa["c2"][t],
                       pa["pagamenti"][t], pa["pv_pagamenti"][t]])

    wsn = wb.create_sheet("Sensitivity")
    wsn.append(["Azienda", "Inflazione \\ Attualizzazione", "-10%", "100%", "+10%"])
    etich = ["-10%", "100%", "+10%"]
    for k in sorted(sens, key=lambda x: (x == 0, x)):
        g = sens[k]
        for i in range(g.shape[0]):
            wsn.append([k or "Totale", etich[i]] + [float(x) for x in g[i]])
    wb.save(path)
