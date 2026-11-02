"""Prospetto delle movimentazioni trimestrali per la transizione all'IFRS 18.

L'IFRS 18 (Presentazione e informativa nel bilancio, in vigore dal 1° gennaio 2027) si applica
retroattivamente: i periodi comparativi 2026 vanno ripresentati e i bilanci intermedi 2027 (IAS 34)
richiedono i comparativi dei corrispondenti trimestri 2026. Per il TFR la nuova classificazione del
conto economico è:

- service cost                -> categoria operativa;
- net interest expense (NIE)  -> categoria di finanziamento (interessi su passività che non derivano
                                 da operazioni di sola raccolta di finanziamenti);
- utili/perdite attuariali    -> altre componenti del conto economico complessivo, non riclassificabili
                                 (invariato rispetto allo IAS 19);
- benefici pagati e trasferimenti -> movimenti patrimoniali, fuori dal conto economico.

Le valutazioni infrannuali sono progressive dall'inizio dell'esercizio (stessa DBO di apertura): i
valori del singolo trimestre si ottengono per differenza tra progressivi consecutivi.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Dict, List, Mapping, Optional, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .riconciliazione import Movimentazione

VOCI = [
    # (chiave, descrizione, categoria IFRS 18)
    ("dbo_apertura", "Passività (DBO) di inizio periodo", ""),
    ("service_cost", "Service cost", "Conto economico - categoria operativa"),
    ("nie", "Net interest expense (interest cost)", "Conto economico - categoria di finanziamento"),
    ("ce", "Totale a conto economico", ""),
    ("perdite_finanziarie", "(Utili)/perdite attuariali da ipotesi finanziarie", "OCI - non riclassificabili"),
    ("perdite_demografiche", "(Utili)/perdite attuariali da ipotesi demografiche", "OCI - non riclassificabili"),
    ("perdite_esperienza", "(Utili)/perdite attuariali da esperienza", "OCI - non riclassificabili"),
    ("oci", "Totale altre componenti del conto economico complessivo", ""),
    ("variazioni_fondo", "Benefici pagati, trasferimenti e altre variazioni", "Stato patrimoniale (fuori CE)"),
    ("dbo_chiusura", "Passività (DBO) di fine periodo", ""),
]
TOTALI = {"dbo_apertura", "ce", "oci", "dbo_chiusura"}


@dataclass
class Periodo:
    fine: date
    movimenti: Mapping[int, Movimentazione]   # progressivi dall'inizio dell'esercizio, 0 = totale


def _valori(m: Movimentazione) -> Dict[str, float]:
    v = {
        "dbo_apertura": m.dbo_apertura,
        "service_cost": m.service_cost,
        "nie": m.nie,
        "perdite_finanziarie": -m.utile_finanziarie,
        "perdite_demografiche": -m.utile_demografiche,
        "perdite_esperienza": -m.utile_esperienza,
        "variazioni_fondo": m.variazioni_fondo,
        "dbo_chiusura": m.dbo_chiusura,
    }
    v["ce"] = v["service_cost"] + v["nie"]
    v["oci"] = v["perdite_finanziarie"] + v["perdite_demografiche"] + v["perdite_esperienza"]
    return {c: (0.0 if abs(x) < 0.005 else x) for c, x in v.items()}


def _apertura(periodi: Sequence[Periodo]) -> str:
    """Data di inizio esercizio (31/12 dell'anno precedente l'ultimo periodo)."""
    return f"31/12/{periodi[-1].fine.year - 1}"


def colonne(periodi: Sequence[Periodo], k: int, trimestri: bool = False) -> List[tuple]:
    """Colonne (etichetta, valori).

    Default: una colonna per valutazione, progressiva dall'inizio dell'esercizio (DBO di apertura
    sempre quella di fine esercizio precedente). Con `trimestri`: singoli trimestri per differenza
    tra progressivi consecutivi, seguiti dal progressivo dell'ultimo periodo.
    """
    vuoto = {c: 0.0 for c, _, _ in VOCI}
    if not trimestri:
        return [(f"{p.fine:%d/%m/%Y} ({p.fine.month} mesi)", _valori(p.movimenti[k]) if k in p.movimenti else vuoto)
                for p in periodi]
    out = []
    prec: Optional[Dict[str, float]] = None
    for p in periodi:
        cur = _valori(p.movimenti[k]) if k in p.movimenti else dict(vuoto)
        if prec is None:
            disc = dict(cur)
        else:
            disc = {c: (0.0 if abs(cur[c] - prec[c]) < 0.005 else cur[c] - prec[c]) for c in cur}
            disc["dbo_apertura"] = prec["dbo_chiusura"]
            disc["dbo_chiusura"] = cur["dbo_chiusura"]
        trimestre = (p.fine.month - 1) // 3 + 1
        out.append((f"{trimestre}° trim. {p.fine.year}", disc))
        prec = cur
    ultimo = periodi[-1]
    out.append((f"Progressivo {ultimo.fine.month} mesi {ultimo.fine.year}",
                _valori(ultimo.movimenti[k]) if k in ultimo.movimenti else vuoto))
    return out


EUR = '#,##0.00;[Red]-#,##0.00'
TESTA = PatternFill("solid", fgColor="DDE5F0")
TOT = PatternFill("solid", fgColor="F2F2F2")
SOTTILE = Side(style="thin", color="BFBFBF")


def _blocco(ws, titolo: str, cols: List[tuple], confronto: Optional[tuple] = None) -> None:
    ws.append([titolo])
    ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    tutte = cols + ([confronto] if confronto else [])
    ws.append(["Voce", "Classificazione IFRS 18"] + [c[0] for c in tutte])
    r0 = ws.max_row
    for c in ws[r0]:
        c.font = Font(bold=True)
        c.fill = TESTA
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for chiave, descr, cat in VOCI:
        ws.append([descr, cat] + [c[1].get(chiave) for c in tutte])
        riga = ws[ws.max_row]
        for c in riga[2:]:
            c.number_format = EUR
        if chiave in TOTALI:
            for c in riga:
                c.font = Font(bold=True)
                c.fill = TOT
    # controllo di quadratura
    ws.append(["Controllo: apertura + CE + OCI + variazioni - chiusura", ""] +
              [round(c[1]["dbo_apertura"] + c[1]["ce"] + c[1]["oci"] + c[1]["variazioni_fondo"]
                     - c[1]["dbo_chiusura"], 2) for c in tutte])
    for c in ws[ws.max_row]:
        c.font = Font(italic=True, color="808080")
    for row in ws.iter_rows(min_row=r0, max_row=ws.max_row):
        for c in row:
            c.border = Border(top=SOTTILE, bottom=SOTTILE, left=SOTTILE, right=SOTTILE)
    ws.append([])
    ws.append([])


def scrivi_prospetto(path: str, periodi: Sequence[Periodo], nomi: Mapping[int, str], cliente: str,
                     confronto: Optional[Mapping[int, Movimentazione]] = None,
                     etichetta_confronto: str = "Esercizio precedente",
                     note: Sequence[str] = (), trimestri: bool = False) -> None:
    periodi = sorted(periodi, key=lambda p: p.fine)
    anno = periodi[-1].fine.year
    wb = Workbook()
    ws = wb.active
    ws.title = "Gruppo"
    ws.append([f"{cliente} - Movimentazione trimestrale del TFR (IAS 19) {anno} ai fini dell'IFRS 18"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["Importi in euro. " + ("Trimestri ottenuti per differenza tra valutazioni progressive."
                                      if trimestri else
                                      f"Valori progressivi dal {_apertura(periodi)}: "
                                      "passività di inizio periodo sempre quella di fine esercizio precedente.")])
    ws.append([])
    conf = (lambda k: (etichetta_confronto, _valori(confronto[k])) if confronto and k in confronto else None)
    _blocco(ws, cliente, colonne(periodi, 0, trimestri), conf(0))
    ws2 = wb.create_sheet("Per società")
    chiavi = sorted({k for p in periodi for k in p.movimenti if k})
    for k in chiavi:
        _blocco(ws2, nomi.get(k, f"Azienda {k}"), colonne(periodi, k, trimestri), conf(k))
    for w in (ws, ws2):
        w.column_dimensions["A"].width = 52
        w.column_dimensions["B"].width = 40
        for col in "CDEFGH":
            w.column_dimensions[col].width = 18
        w.freeze_panes = "C1"

    # sintesi per categoria IFRS 18 e trimestre
    ws3 = wb.create_sheet("Sintesi CE")
    ws3.append(["Società", "Periodo", "Categoria operativa (SC)", "Categoria di finanziamento (NIE)",
                "Totale conto economico", "OCI (utili)/perdite attuariali"])
    for c in ws3[1]:
        c.font = Font(bold=True)
        c.fill = TESTA
        c.alignment = Alignment(wrap_text=True, horizontal="center")
    for k in [0] + chiavi:
        for etich, v in colonne(periodi, k, trimestri):
            ws3.append([cliente if k == 0 else nomi.get(k, str(k)), etich, v["service_cost"], v["nie"], v["ce"], v["oci"]])
            for c in ws3[ws3.max_row][2:]:
                c.number_format = EUR
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 24
    for col in "CDEF":
        ws3.column_dimensions[col].width = 22
    ws3.freeze_panes = "C2"

    ws4 = wb.create_sheet("Note")
    righe = [
        "Note metodologiche",
        "",
        "1. L'IFRS 18 sostituisce lo IAS 1 dagli esercizi che iniziano il 1° gennaio 2027 e si applica "
        "retroattivamente: il conto economico 2026 va ripresentato con le nuove categorie e i bilanci intermedi "
        "2027 (IAS 34) richiedono i comparativi dei trimestri 2026.",
        "2. Service cost: costo del personale, categoria operativa.",
        "3. Net interest expense: interessi su una passività che non deriva da operazioni di sola raccolta di "
        "finanziamenti, classificati nella categoria di finanziamento.",
        "4. Utili e perdite attuariali: altre componenti del conto economico complessivo non riclassificabili "
        "(IAS 19, par. 120 e 122), senza effetti sul conto economico; scomposti tra ipotesi finanziarie, "
        "ipotesi demografiche ed esperienza.",
        "5. Benefici pagati, trasferimenti e altre variazioni del fondo: movimenti dello stato patrimoniale, "
        "fuori dal conto economico.",
        ("6. Le valutazioni infrannuali sono progressive dall'inizio dell'esercizio; i valori del singolo "
         "trimestre sono la differenza tra progressivi consecutivi. La DBO di inizio trimestre è la DBO della "
         "valutazione del trimestre precedente." if trimestri else
         "6. Ogni colonna è progressiva dall'inizio dell'esercizio: la passività di inizio periodo è sempre "
         "quella di fine esercizio precedente e i flussi si riferiscono ai mesi trascorsi."),
        "7. Il net interest expense è calcolato sulla DBO di inizio esercizio al tasso forward del primo anno "
        "della valutazione di chiusura dell'esercizio precedente, per i giorni trascorsi su 365.",
    ] + list(note)
    for r in righe:
        ws4.append([r])
    ws4["A1"].font = Font(bold=True, size=12)
    ws4.column_dimensions["A"].width = 140
    for row in ws4.iter_rows(min_row=3):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)
