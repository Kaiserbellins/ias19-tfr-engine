"""Excel dei risultati: schede per azienda e tabelle di disclosure.

Fogli prodotti:
- Riepilogo        : scheda di valutazione per azienda (CCP, DW, PSL, DC-R/A, DC-Inf, DC-Sal, CSL, DBO),
                     con valori medi per dipendente e rapporto su CCP (come le schede della relazione GHC)
- Movimentazione   : DBO apertura, SC, NIE, VF, DBO chiusura, utili/perdite e scomposizione (Tabella 7)
- Sensitivity      : griglia 3x3 su attualizzazione e inflazione, valori assoluti e relativi
- Sensitivity ass. : variazioni assolute (stile BE), se calcolate
- Andamento        : per anno di uscita CCP pagato, PSL, CSL, DBO e residui (dati della Figura 1)
- Disclosure       : duration, service cost pro futuro, erogazioni attese anni 1-5, 6-10 e oltre
- Statistiche      : dipendenti, salario medio e TFR medio per classi di eta e anzianita
- Dipendenti       : risultati individuali
"""
from __future__ import annotations

from typing import Dict, List, Mapping, Optional, Sequence

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .engine import proietta
from .io_excel import DatiValutazione
from .model import Ipotesi, Risultato

EUR = '#,##0.00;[Red]-#,##0.00'
PCT = '0.00%'
BOLD = Font(bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DDE5F0")

CLASSI_ETA = [(0, 19, "fino a 19"), (20, 24, "20-24"), (25, 29, "25-29"), (30, 34, "30-34"),
              (35, 39, "35-39"), (40, 44, "40-44"), (45, 49, "45-49"), (50, 54, "50-54"),
              (55, 59, "55-59"), (60, 64, "60-64"), (65, 200, "65 e oltre")]
CLASSI_ANZ = [(0, 5, "fino a 5"), (6, 10, "6-10"), (11, 15, "11-15"), (16, 20, "16-20"),
              (21, 25, "21-25"), (26, 30, "26-30"), (31, 35, "31-35"), (36, 200, "36 e oltre")]


def _intestazione(ws, valori: Sequence) -> None:
    ws.append(list(valori))
    r = ws.max_row
    for c in ws[r]:
        c.font = BOLD
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")


def _formato(ws, min_row: int, min_col: int, max_col: int, fmt: str) -> None:
    for row in ws.iter_rows(min_row=min_row, min_col=min_col, max_col=max_col):
        for c in row:
            if isinstance(c.value, (int, float)):
                c.number_format = fmt


def _larghezze(ws, prima: int = 28, altre: int = 16) -> None:
    for i in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(i)].width = prima if i == 1 else altre


def _ordine(chiavi):
    return sorted(chiavi, key=lambda x: (x == 0, x))


def _nome(dv: DatiValutazione, k: int) -> str:
    if k == 0:
        return "Totale gruppo"
    return dv.aziende[k].nome if k in dv.aziende else f"Azienda {k}"


def scrivi_risultati(path: str, dv: DatiValutazione, res: Dict[int, Risultato],
                     csl: Dict[int, Dict[str, float]], sens: Dict[int, np.ndarray], metodo: str,
                     ip: Optional[Ipotesi] = None, movimenti: Optional[Mapping] = None,
                     sens_assoluta: Optional[Dict[str, Dict[int, float]]] = None,
                     data_valutazione: str = "") -> None:
    wb = Workbook()
    _riepilogo(wb.active, dv, res, csl, metodo, data_valutazione)
    if movimenti:
        _movimentazione(wb.create_sheet("Movimentazione"), dv, movimenti)
    _sensitivity(wb.create_sheet("Sensitivity"), dv, sens)
    if sens_assoluta:
        _sensitivity_assoluta(wb.create_sheet("Sensitivity ass."), dv, sens_assoluta)
    _andamento(wb.create_sheet("Andamento"), dv, res)
    _disclosure(wb.create_sheet("Disclosure"), dv, res, metodo)
    _statistiche(wb.create_sheet("Statistiche"), dv)
    if ip is not None:
        _dipendenti(wb.create_sheet("Dipendenti"), dv, ip)
    wb.save(path)


def _riepilogo(ws, dv, res, csl, metodo, data_val):
    ws.title = "Riepilogo"
    ws.append([f"Valutazione IAS 19 del TFR {('al ' + data_val) if data_val else ''}".strip()])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"Metodo: {'Financial proration' if metodo == 'financial' else 'Linear proration'}"])
    ws.append([])
    voci = [("Current Concern Provision", "CCP"), ("Difference CCP", "DW"), ("Past Service Liability", "PSL"),
            ("Difference CCP: R/A", "DC-R/A"), ("Difference CCP: Infla", "DC-Inf"),
            ("Difference CCP: Salari", "DC-Sal"), ("Current Service Liability", "CSL"),
            ("Total Difference", "DW+CSL"), ("Defined Benefit Obligation", "DBO")]
    _intestazione(ws, ["Azienda", "Dipendenti"] + [s for _, s in voci] + ["DBO lineare", "DBO / CCP"])
    for k in _ordine(res):
        r, d = res[k], csl[k]
        dbo = r.dbo if metodo == "financial" else r.dbo_lineare
        valori = [r.ccp, r.dw, r.psl, d["DC-R/A"], d["DC-Inf"], d["DC-Sal"], r.csl, r.dw + r.csl, r.dbo]
        ws.append([_nome(dv, k), r.n] + valori + [r.dbo_lineare, (dbo / r.ccp) if r.ccp else None])
        if k == 0:
            for c in ws[ws.max_row]:
                c.font = BOLD
    _formato(ws, 5, 3, 12, EUR)
    _formato(ws, 5, 13, 13, PCT)
    _larghezze(ws)
    ws.freeze_panes = "C5"

    # scheda nel formato della relazione per il totale e per ciascuna azienda
    ws.append([])
    for k in _ordine(res):
        r, d = res[k], csl[k]
        ws.append([])
        ws.append([_nome(dv, k)])
        ws[ws.max_row][0].font = Font(bold=True, size=12)
        _intestazione(ws, ["Voce", "Sigla", "Financial", "Medie", "Fin/CCP"])
        valori = {"CCP": r.ccp, "DW": r.dw, "PSL": r.psl, "DC-R/A": d["DC-R/A"], "DC-Inf": d["DC-Inf"],
                  "DC-Sal": d["DC-Sal"], "CSL": r.csl, "DW+CSL": r.dw + r.csl, "DBO": r.dbo}
        inizio = ws.max_row + 1
        for nome, sigla in voci:
            v = valori[sigla]
            ws.append([nome, sigla, v, v / r.n if r.n else None, v / r.ccp if r.ccp else None])
        _formato(ws, inizio, 3, 4, EUR)
        _formato(ws, inizio, 5, 5, PCT)


def _movimentazione(ws, dv, mov):
    _intestazione(ws, ["Azienda", "DBO apertura", "Service Cost", "Net Interest Expense", "Variazioni fondo",
                       "DBO atteso", "DBO chiusura", "Utile/(Perdita) attuariale", "di cui esperienza",
                       "di cui ipotesi demografiche", "di cui ipotesi finanziarie"])
    for k in _ordine(mov):
        m = mov[k]
        ws.append([_nome(dv, k), m.dbo_apertura, m.service_cost, m.nie, m.variazioni_fondo, m.dbo_atteso,
                   m.dbo_chiusura, m.utile, m.utile_esperienza, m.utile_demografiche, m.utile_finanziarie])
        if k == 0:
            for c in ws[ws.max_row]:
                c.font = BOLD
    _formato(ws, 2, 2, 11, EUR)
    _larghezze(ws)
    ws.freeze_panes = "B2"
    ws.append([])
    ws.append(["Utile/(Perdita) = DBO apertura + SC + NIE + VF - DBO chiusura; positivo = utile."])


def _sensitivity(ws, dv, sens):
    etich = ["-10%", "100%", "+10%"]
    for k in _ordine(sens):
        g = sens[k]
        base = g[1, 1]
        ws.append([_nome(dv, k)])
        ws[ws.max_row][0].font = Font(bold=True, size=12)
        _intestazione(ws, ["Inflazione \\ Attualizzazione"] + etich + [""] + etich)
        inizio = ws.max_row + 1
        for i, e in enumerate(etich):
            rel = [x / base if base else None for x in g[i]]
            ws.append([e] + [float(x) for x in g[i]] + [None] + rel)
        _formato(ws, inizio, 2, 4, EUR)
        _formato(ws, inizio, 6, 8, PCT)
        ws.append([])
    _larghezze(ws)


def _sensitivity_assoluta(ws, dv, sa):
    chiavi = _ordine({k for v in sa.values() for k in v})
    _intestazione(ws, ["Scenario"] + [_nome(dv, k) for k in chiavi])
    for nome, valori in sa.items():
        ws.append([nome] + [valori.get(k) for k in chiavi])
    _formato(ws, 2, 2, len(chiavi) + 1, EUR)
    _larghezze(ws, 34)


def _andamento(ws, dv, res):
    _intestazione(ws, ["Azienda", "Anno", "CCP pagato", "PSL", "CSL", "DBO", "CCP residuo", "DBO residuo"])
    for k in _ordine(res):
        pa = res[k].per_anno
        csl = pa["c2"] - pa["c1"]
        dbo = pa["w2"] + csl
        res_ccp = res[k].ccp - np.cumsum(pa["w1"])
        res_dbo = res[k].dbo - np.cumsum(dbo)
        ultimo = int(np.max(np.nonzero(pa["w1"] + np.abs(pa["c1"]))[0])) + 1 if (pa["w1"] + np.abs(pa["c1"])).any() else 0
        for t in range(ultimo):
            ws.append([_nome(dv, k), t + 1, pa["w1"][t], pa["w2"][t], csl[t], dbo[t], res_ccp[t], res_dbo[t]])
    _formato(ws, 2, 3, 8, EUR)
    _larghezze(ws)
    ws.freeze_panes = "C2"


def _disclosure(ws, dv, res, metodo):
    _intestazione(ws, ["Azienda", "Duration (anni)", "Service cost pro futuro", "Erogazioni anno 1",
                       "Erogazioni anno 2", "Erogazioni anno 3", "Erogazioni anno 4", "Erogazioni anno 5",
                       "Erogazioni anni 6-10", "Erogazioni oltre 10 anni"])
    for k in _ordine(res):
        r = res[k]
        p = r.per_anno["pagamenti"]
        sc = r.sc_fin if metodo == "financial" else r.sc_lin
        ws.append([_nome(dv, k), round(r.duration, 2), sc] + [float(x) for x in p[:5]]
                  + [float(p[5:10].sum()), float(p[10:].sum())])
    _formato(ws, 2, 3, 10, EUR)
    _larghezze(ws)
    ws.append([])
    ws.append(["Erogazioni attese nominali: TFR maturato rivalutato e quote future rivalutate, per anno di uscita."])


def _classe(v, classi):
    for i, (lo, hi, _) in enumerate(classi):
        if lo <= v <= hi:
            return i
    return len(classi) - 1


def _statistiche(ws, dv):
    gruppi: Dict[int, list] = {0: list(dv.dipendenti)}
    for e in dv.dipendenti:
        gruppi.setdefault(e.azienda, []).append(e)
    titoli = [("Tab. 1 - Distribuzione per eta e anzianita (numero dipendenti)", None),
              ("Tab. 2 - Retribuzioni medie", "salario"), ("Tab. 3 - TFR accantonato medio", "tfr")]
    for k in _ordine(gruppi):
        ws.append([_nome(dv, k)])
        ws[ws.max_row][0].font = Font(bold=True, size=12)
        for titolo, campo in titoli:
            ws.append([titolo])
            _intestazione(ws, ["Eta \\ Anzianita"] + [c[2] for c in CLASSI_ANZ] + ["Totale"])
            n = np.zeros((len(CLASSI_ETA), len(CLASSI_ANZ)))
            s = np.zeros_like(n)
            for e in gruppi[k]:
                i, j = _classe(e.eta, CLASSI_ETA), _classe(e.anzianita, CLASSI_ANZ)
                n[i, j] += 1
                if campo:
                    s[i, j] += getattr(e, campo)
            inizio = ws.max_row + 1
            for i, c in enumerate(CLASSI_ETA):
                if campo:
                    riga = [s[i, j] / n[i, j] if n[i, j] else 0 for j in range(len(CLASSI_ANZ))]
                    tot = s[i].sum() / n[i].sum() if n[i].sum() else 0
                else:
                    riga, tot = list(n[i]), n[i].sum()
                ws.append([c[2]] + [float(x) for x in riga] + [float(tot)])
            if campo:
                riga = [s[:, j].sum() / n[:, j].sum() if n[:, j].sum() else 0 for j in range(len(CLASSI_ANZ))]
                tot = s.sum() / n.sum() if n.sum() else 0
            else:
                riga, tot = list(n.sum(axis=0)), n.sum()
            ws.append(["Totale"] + [float(x) for x in riga] + [float(tot)])
            _formato(ws, inizio, 2, len(CLASSI_ANZ) + 2, EUR if campo else '0')
            ws.append([])
    _larghezze(ws, 30, 12)


def _dipendenti(ws, dv, ip):
    _intestazione(ws, ["Codice", "Azienda", "Eta", "Sesso", "Anzianita", "Futuro", "Carriera", "Salario",
                       "TFR (CCP)", "Aliquota FIP", "DW", "PSL", "CSL", "DBO", "DBO lineare"])
    for e in dv.dipendenti:
        f = proietta(e, ip)
        psl = float(f.w2.sum())
        csl = float(f.c2.sum() - f.c1.sum())
        ws.append([e.codice, e.azienda, e.eta, e.sesso, e.anzianita, e.futuro, e.carriera, e.salario, e.tfr,
                   e.fip, psl - e.tfr, psl, csl, psl + csl, float(f.dbo_lineare.sum())])
    _formato(ws, 2, 8, 9, EUR)
    _formato(ws, 2, 10, 10, '0.0000%')
    _formato(ws, 2, 11, 15, EUR)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
