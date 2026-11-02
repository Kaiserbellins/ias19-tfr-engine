"""Relazione tecnica Word generata a partire da una relazione modello (.docx).

Il modello e una relazione gia emessa (es. "Relazione GHCSPA IAS19 2026_09.docx"): se ne conservano
stili, copertina, intestazioni, testi metodologici e formule, e si rigenerano le parti che dipendono
dalla valutazione:

- nome del cliente e data di valutazione (proprieta del documento e campi collegati);
- Tabelle 1-6 (RG48, turnover e anticipazioni, tavole di permanenza, tassi annui) dalle ipotesi;
- incrementi salariali per categoria, elenco delle societa e numero dei dipendenti;
- caratteristiche del collettivo (tabelle per societa, eta e anzianita);
- Tabella 7 (movimentazione per societa) e scomposizione degli utili/perdite attuariali;
- informativa aggiuntiva IAS 19: riconciliazione della DBO, ipotesi principali, sensitivity con
  variazioni assolute, duration, service cost pro futuro, erogazioni attese;
- allegati: scheda di valutazione e grafico dell'andamento prospettico per ogni societa;
- sensitivity 3x3 (valori assoluti e relativi) per ogni societa.

Le tabelle, che nel modello sono immagini incollate da Excel, diventano tabelle Word. La
numerazione delle didascalie viene ricalcolata; sommario e indici si aggiornano all'apertura
del documento in Word (campo `updateFields`).
"""
from __future__ import annotations

import copy
import io
import re
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Mapping, Optional, Sequence, Tuple

import numpy as np

from .io_excel import DatiValutazione
from .model import N_ANNI, Ipotesi, Risultato

MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto",
        "settembre", "ottobre", "novembre", "dicembre"]

CLASSI_ETA = [(0, 24, "fino a 24"), (25, 29, "25-29"), (30, 34, "30-34"), (35, 39, "35-39"),
              (40, 44, "40-44"), (45, 49, "45-49"), (50, 54, "50-54"), (55, 59, "55-59"),
              (60, 200, "60 e oltre")]
CLASSI_ANZ = [(0, 4, "0-4"), (5, 9, "5-9"), (10, 14, "10-14"), (15, 19, "15-19"),
              (20, 24, "20-24"), (25, 29, "25-29"), (30, 200, "30 e oltre")]

ETA_INGRESSO = 15          # le tavole di permanenza sono mostrate con radice 100.000 a questa eta
LARGHEZZA_CM = 15.4        # larghezza utile della pagina del modello
COL_CCP = "2A78D6"
COL_DBO = "EB6834"
ARANCIO = "FFC000"
GIALLO = "FFFF00"
ROSSO = "FF0000"
GRIGIO = "DDE5F0"

DEFINIZIONI = [
    ("(CCP) Current Concern Provision", "valore contabile del TFR aziendale"),
    ("(DW) Differenziale di valutazione del TFR contabile", ""),
    ("(PSL) Past Service Liability", "valore attuale medio delle prestazioni pagabili in futuro, come "
     "contropartita dell'attività lavorativa pregressa"),
    ("(CSL) Current Service Liability", "valore attuale medio delle prestazioni pagabili in futuro come "
     "contropartita dei servizi offerti, al netto dei contributi futuri"),
    ("(DBO) Defined Benefit Obligation", "valore attuale medio delle prestazioni pagabili in futuro come "
     "contropartita ai servizi lavorativi resi sino alla data di valutazione"),
]


@dataclass
class DatiRelazione:
    cliente: str
    data_valutazione: date
    dv: DatiValutazione
    res: Dict[int, Risultato]
    csl: Dict[int, Dict[str, float]]
    sens: Dict[int, np.ndarray]
    ip: Ipotesi
    metodo: str = "financial"
    movimenti: Optional[Mapping] = None
    sens_assoluta: Optional[Dict[str, Dict[int, float]]] = None
    data_precedente: Optional[date] = None
    ip_precedente: Optional[Ipotesi] = None
    trasferimenti: Dict[int, float] = field(default_factory=dict)
    categorie: Dict[int, str] = field(default_factory=dict)
    data_emissione: Optional[date] = None


# ---------------------------------------------------------------- formati

def data_lunga(d: date) -> str:
    return f"{d.day} {MESI[d.month - 1]} {d.year}"


def numero(v: float, dec: int = 2) -> str:
    s = f"{v:,.{dec}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def euro(v: float, dec: int = 2) -> str:
    return "-" if v is None or abs(v) < 0.5 * 10 ** -dec else numero(v, dec)


def pct(v: float, dec: int = 2) -> str:
    return numero(100 * v, dec) + "%"


# ---------------------------------------------------------------- XML di basso livello

def _ns():
    from docx.oxml.ns import qn
    return qn


def _testo(el) -> str:
    qn = _ns()
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _sostituisci(p, vecchio: str, nuovo: str) -> bool:
    """Sostituisce `vecchio` con `nuovo` nel paragrafo anche se il testo e spezzato su piu run."""
    qn = _ns()
    if not vecchio or vecchio == nuovo:
        return False
    fatto = False
    inizio = 0
    while True:
        nodi = list(p.iter(qn("w:t")))
        full = "".join(t.text or "" for t in nodi)
        i = full.find(vecchio, inizio)
        if i < 0:
            return fatto
        fine = i + len(vecchio)
        off = 0
        primo = None
        for t in nodi:
            s = t.text or ""
            a, b = off, off + len(s)
            off = b
            if b <= i or a >= fine:
                continue
            if primo is None:
                primo = t
                t.text = s[: i - a] + nuovo + (s[fine - a:] if fine <= b else "")
            else:
                t.text = s[fine - a:] if fine < b else ""
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        fatto = True
        inizio = i + len(nuovo)


def _rimuovi(el) -> None:
    el.getparent().remove(el)


def _senza_segnalibri(el):
    qn = _ns()
    for tag in ("w:bookmarkStart", "w:bookmarkEnd"):
        for b in list(el.iter(qn(tag))):
            b.getparent().remove(b)
    return el


def _clona(el):
    return _senza_segnalibri(copy.deepcopy(el))


def _primo_rpr(p):
    qn = _ns()
    r = p.find(qn("w:r"))
    if r is not None and r.find(qn("w:rPr")) is not None:
        return copy.deepcopy(r.find(qn("w:rPr")))
    return None


def _run(testo: str, rpr=None, grassetto: Optional[bool] = None):
    from docx.oxml import OxmlElement
    qn = _ns()
    r = OxmlElement("w:r")
    if rpr is not None:
        rpr = copy.deepcopy(rpr)
        r.append(rpr)
    if grassetto is not None:
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.append(rpr)
        for b in rpr.findall(qn("w:b")):
            rpr.remove(b)
        if grassetto:
            rpr.insert(0, OxmlElement("w:b"))
    parti = testo.split("\t")
    for j, parte in enumerate(parti):
        if j:
            r.append(OxmlElement("w:tab"))
        if parte:
            t = OxmlElement("w:t")
            t.text = parte
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(t)
    return r


def _paragrafo(proto, parti: Sequence[Tuple[str, Optional[bool]]]):
    """Copia di `proto` (stile e numerazione) con nuovo contenuto: parti = [(testo, grassetto)]."""
    qn = _ns()
    p = _clona(proto)
    rpr = _primo_rpr(proto)
    for c in list(p):
        if c.tag != qn("w:pPr"):
            p.remove(c)
    for testo, g in parti:
        p.append(_run(testo, rpr, g))
    return p


def _interruzione_pagina():
    from docx.oxml import OxmlElement
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(_ns()("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def _didascalia(proto, testo: str, centrata: bool = True):
    """Didascalia con campo SEQ copiato dal prototipo e testo descrittivo nuovo."""
    qn = _ns()
    p = _clona(proto)
    figli = list(p)
    fine = None
    in_seq = False
    for i, c in enumerate(figli):
        if c.tag == qn("w:fldSimple") and "SEQ" in (c.get(qn("w:instr")) or ""):
            fine = i
            break
        if c.tag == qn("w:r"):
            it = c.find(qn("w:instrText"))
            if it is not None and "SEQ" in (it.text or ""):
                in_seq = True
            fc = c.find(qn("w:fldChar"))
            if in_seq and fc is not None and fc.get(qn("w:fldCharType")) == "end":
                fine = i
                break
    if fine is None:
        return _paragrafo(proto, [(testo, None)])
    for c in figli[fine + 1:]:
        p.remove(c)
    p.append(_run(f" - {testo}", None))
    if centrata:
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:jc")) is None:
            from docx.oxml import OxmlElement
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), "center")
            ppr.append(jc)
    return p


def _inserisci_dopo(ancora, elementi: Sequence) -> object:
    for el in elementi:
        ancora.addnext(el)
        ancora = el
    return ancora


def _numera_didascalie(body) -> None:
    """Ricalcola i numeri delle didascalie (campi SEQ Tabella / SEQ Figura) nell'ordine del documento."""
    qn = _ns()
    conta: Dict[str, int] = {}
    for p in body.iter(qn("w:p")):
        for fs in p.findall(qn("w:fldSimple")):
            m = re.search(r"SEQ\s+(\w+)", fs.get(qn("w:instr")) or "")
            if m:
                conta[m.group(1)] = conta.get(m.group(1), 0) + 1
                ts = list(fs.iter(qn("w:t")))
                if ts:
                    ts[0].text = str(conta[m.group(1)])
                    for t in ts[1:]:
                        t.text = ""
        stato, seq, primo = None, None, False
        for r in p.iter(qn("w:r")):
            fc = r.find(qn("w:fldChar"))
            it = r.find(qn("w:instrText"))
            if fc is not None:
                tipo = fc.get(qn("w:fldCharType"))
                if tipo == "begin":
                    stato, seq = "istr", None
                elif tipo == "separate":
                    stato = "ris" if seq else None
                    if seq:
                        conta[seq] = conta.get(seq, 0) + 1
                        primo = True
                elif tipo == "end":
                    stato, seq = None, None
            elif it is not None and stato == "istr":
                m = re.search(r"SEQ\s+(\w+)", it.text or "")
                if m:
                    seq = m.group(1)
            elif stato == "ris":
                for t in r.findall(qn("w:t")):
                    t.text = str(conta[seq]) if primo else ""
                    primo = False


# ---------------------------------------------------------------- tabelle esistenti (Tabelle 1-6)

def _testo_cella(tc, testo: str, grassetto: Optional[bool] = None) -> None:
    qn = _ns()
    ps = tc.findall(qn("w:p"))
    p = ps[0]
    for x in ps[1:]:
        tc.remove(x)
    runs = p.findall(qn("w:r"))
    rpr = None
    if runs and runs[0].find(qn("w:rPr")) is not None:
        rpr = runs[0].find(qn("w:rPr"))
    else:
        ppr = p.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:rPr")) is not None:
            rpr = ppr.find(qn("w:rPr"))
    for c in list(p):
        if c.tag != qn("w:pPr"):
            p.remove(c)
    p.append(_run(testo, rpr, grassetto))


def _riempi_tabella(tbl, righe: Sequence[Tuple[str, Sequence[str]]]) -> None:
    """Riscrive una tabella del modello: righe = [("h"|"d", celle)].

    La prima riga del modello fa da prototipo per le intestazioni, la seconda per i dati; il
    numero di colonne si adatta e la larghezza totale resta quella del modello.
    """
    from docx.oxml import OxmlElement
    qn = _ns()
    trs = tbl.findall(qn("w:tr"))
    proto = {"h": copy.deepcopy(trs[0]), "d": copy.deepcopy(trs[1] if len(trs) > 1 else trs[0])}
    for tr in trs:
        tbl.remove(tr)
    ncol = max(len(c) for _, c in righe)
    grid = tbl.find(qn("w:tblGrid"))
    totale = sum(int(g.get(qn("w:w")) or 0) for g in grid.findall(qn("w:gridCol"))) or 8787
    larg = totale // ncol
    for g in list(grid):
        grid.remove(g)
    for _ in range(ncol):
        g = OxmlElement("w:gridCol")
        g.set(qn("w:w"), str(larg))
        grid.append(g)
    for tipo, celle in righe:
        tr = copy.deepcopy(proto[tipo])
        tcs = tr.findall(qn("w:tc"))
        while len(tcs) < ncol:
            nuovo = copy.deepcopy(tcs[-1])
            tcs[-1].addnext(nuovo)
            tcs.append(nuovo)
        for tc in tcs[ncol:]:
            tr.remove(tc)
        tcs = tcs[:ncol]
        for tc, testo in zip(tcs, list(celle) + [""] * (ncol - len(celle))):
            tcpr = tc.find(qn("w:tcPr"))
            if tcpr is not None:
                for x in tcpr.findall(qn("w:gridSpan")) + tcpr.findall(qn("w:vMerge")):
                    tcpr.remove(x)
                w = tcpr.find(qn("w:tcW"))
                if w is not None:
                    w.set(qn("w:w"), str(larg))
                    w.set(qn("w:type"), "dxa")
            _testo_cella(tc, testo)
        tbl.append(tr)


# ---------------------------------------------------------------- tabelle nuove

class _Costruttore:
    """Crea tabelle Word nel documento con lo stile del modello."""

    def __init__(self, doc):
        self.doc = doc
        stili = [s.name for s in doc.styles if s.type == 3]
        self.stile = "Table Grid" if "Table Grid" in stili else None

    def tabella(self, righe: Sequence[Sequence[str]], larghezze_cm: Sequence[float],
                intestazioni: int = 1, allinea: Optional[Sequence[str]] = None,
                dimensione: float = 8, grassetto_righe: Sequence[int] = (),
                colori: Optional[Mapping[Tuple[int, int], str]] = None,
                unioni: Sequence[Tuple[int, int, int, int]] = ()):
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Cm, Pt
        ncol = len(larghezze_cm)
        t = self.doc.add_table(rows=len(righe), cols=ncol)
        if self.stile:
            t.style = self.doc.styles[self.stile]
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        allinea = allinea or (["l"] + ["r"] * (ncol - 1))
        mappa = {"l": WD_ALIGN_PARAGRAPH.LEFT, "r": WD_ALIGN_PARAGRAPH.RIGHT, "c": WD_ALIGN_PARAGRAPH.CENTER}
        for i, riga in enumerate(righe):
            for j in range(ncol):
                cell = t.cell(i, j)
                cell.width = Cm(larghezze_cm[j])
                testo = riga[j] if j < len(riga) else ""
                p = cell.paragraphs[0]
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = 1.0
                pf.first_line_indent = Pt(0)
                p.alignment = mappa["c"] if i < intestazioni else mappa[allinea[j]]
                run = p.add_run(testo)
                run.font.size = Pt(dimensione)
                if i < intestazioni or i in grassetto_righe:
                    run.bold = True
                if i < intestazioni:
                    _sfondo(cell, GRIGIO)
                if colori and (i, j) in colori:
                    _sfondo(cell, colori[(i, j)])
        for r0, c0, r1, c1 in unioni:
            a = t.cell(r0, c0)
            testo = a.text
            m = a.merge(t.cell(r1, c1))
            for p in m.paragraphs[1:]:
                p._p.getparent().remove(p._p)
            for r in m.paragraphs[0].runs[1:]:
                r._r.getparent().remove(r._r)
            if m.paragraphs[0].runs:
                m.paragraphs[0].runs[0].text = testo
        el = t._tbl
        from docx.oxml import OxmlElement
        qn = _ns()
        for i, tr in enumerate(el.findall(qn("w:tr"))):
            trpr = tr.get_or_add_trPr()
            trpr.append(OxmlElement("w:cantSplit"))
            if i < intestazioni:
                trpr.append(OxmlElement("w:tblHeader"))
        grid = el.find(qn("w:tblGrid"))
        for g in list(grid):
            grid.remove(g)
        for w in larghezze_cm:
            g = OxmlElement("w:gridCol")
            g.set(qn("w:w"), str(int(w * 567)))
            grid.append(g)
        tblpr = el.find(qn("w:tblPr"))
        lay = OxmlElement("w:tblLayout")
        lay.set(qn("w:type"), "fixed")
        tblpr.append(lay)
        el.getparent().remove(el)
        return el

    def immagine(self, png: bytes, larghezza_cm: float):
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Cm
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(io.BytesIO(png), width=Cm(larghezza_cm))
        el = p._p
        el.getparent().remove(el)
        return el


def _sfondo(cell, colore: str) -> None:
    from docx.oxml import OxmlElement
    qn = _ns()
    tcpr = cell._tc.get_or_add_tcPr()
    for s in tcpr.findall(qn("w:shd")):
        tcpr.remove(s)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), colore)
    tcpr.append(shd)


# ---------------------------------------------------------------- contenuti

def _nome(r: DatiRelazione, k: int) -> str:
    if k == 0:
        return r.cliente
    return r.dv.aziende[k].nome if k in r.dv.aziende else f"Azienda {k}"


def _ordine(chiavi) -> List[int]:
    return sorted(chiavi, key=lambda x: (x != 0, x))


def _categorie_usate(r: DatiRelazione) -> List[int]:
    return sorted({e.carriera for e in r.dv.dipendenti})


def _nome_categoria(r: DatiRelazione, p: int) -> str:
    return r.categorie.get(p, f"{p}")


def _intervalli_costanti(valori: Sequence[Tuple[float, ...]], eta0: int) -> List[Tuple[str, Tuple[float, ...]]]:
    """Raggruppa eta consecutive con le stesse frequenze: [("16..111", (d1, d2, ...))]."""
    out = []
    inizio = 0
    for x in range(1, len(valori) + 1):
        if x == len(valori) or not np.allclose(valori[x], valori[inizio], atol=1e-12):
            a, b = eta0 + inizio, eta0 + x - 1
            out.append((f"{a}..{b}" if b > a else f"{a}", valori[inizio]))
            inizio = x
    return out


def _gruppi_categorie(ip: Ipotesi, sesso: int, cat: Sequence[int]) -> List[Tuple[int, str]]:
    """Categorie con la stessa tavola di permanenza, mostrate in un'unica colonna: [(p, "1-3")]."""
    gruppi: List[List[int]] = []
    for p in cat:
        for g in gruppi:
            if np.allclose(ip.tavola.lx[(sesso, g[0])], ip.tavola.lx[(sesso, p)], rtol=0, atol=1e-9):
                g.append(p)
                break
        else:
            gruppi.append([p])
    out = []
    for g in gruppi:
        consecutive = g == list(range(g[0], g[-1] + 1))
        etich = (f"{g[0]}-{g[-1]}" if consecutive else ",".join(map(str, g))) if len(g) > 1 else str(g[0])
        out.append((g[0], etich))
    return out


def _tabelle_ipotesi(r: DatiRelazione, tabelle: Dict[str, object]) -> None:
    ip = r.ip
    comp = ip.tavola.componenti
    cat_tutte = _categorie_usate(r)
    if comp is not None and "rg48" in tabelle:
        base = comp[0]
        righe = [("h", ["x", "MLx", "FLx"])]
        righe += [("d", [str(x), numero(base[1][x], 6), numero(base[2][x], 6)]) for x in range(len(base[1]))]
        _riempi_tabella(tabelle["rg48"], righe)
    for sesso, s in ((1, "M"), (2, "F")):
        gruppi = _gruppi_categorie(ip, sesso, cat_tutte)
        cat = [p for p, _ in gruppi]
        etich = {p: e for p, e in gruppi}
        chiave = "freq_m" if sesso == 1 else "freq_f"
        if comp is not None and chiave in tabelle:
            dim, ant = comp[1], comp[2]
            righe = []
            for nome, dati in ((f"{s}Dim", dim), (f"{s}Ant", ant)):
                if righe:
                    righe.append(("d", [""] * (len(cat) + 1)))
                righe.append(("h", ["Età"] + [f"{nome}x,p={etich[p]}" for p in cat]))
                valori = [tuple(dati[(sesso, p)][x] for p in cat) for x in range(len(dati[(sesso, cat[0])]))]
                for eta, v in _intervalli_costanti(valori, 0):
                    righe.append(("d", [eta] + [numero(f, 6) for f in v]))
            _riempi_tabella(tabelle[chiave], righe)
        chiave = "lx_m" if sesso == 1 else "lx_f"
        if chiave in tabelle:
            righe = [("h", ["Età"] + [f"{s}Lx,p={etich[p]}" for p in cat])]
            cols = []
            for p in cat:
                lx = ip.tavola.lx[(sesso, p)]
                rif = lx[ETA_INGRESSO] if lx[ETA_INGRESSO] > 0 else 1.0
                cols.append([100000.0 if x <= ETA_INGRESSO else 100000.0 * lx[x] / rif for x in range(len(lx))])
            for x in range(len(cols[0])):
                righe.append(("d", [str(x)] + [numero(c[x], 6) for c in cols]))
            _riempi_tabella(tabelle[chiave], righe)
    if "tassi" in tabelle:
        fin = ip.finanziarie
        riv = fin.tasso_rivalutazione()
        righe = [("h", ["Anno", "Infla", "Rival", "Attua"])]
        righe += [("d", [str(t + 1), pct(fin.inflazione[t], 4), pct(riv[t], 4), pct(fin.attualizzazione[t], 4)])
                  for t in range(len(fin.inflazione))]
        _riempi_tabella(tabelle["tassi"], righe)


def _descrivi_serie(v: np.ndarray, dec: int = 2) -> str:
    if np.allclose(v, v[0], atol=1e-10):
        return pct(float(v[0]), dec)
    return f"{pct(float(v[0]), dec)} il primo anno, variabile nel tempo (da {pct(float(v.min()), dec)} a {pct(float(v.max()), dec)})"


def _tasso_equivalente(pagamenti: np.ndarray, pv: float) -> float:
    """Tasso costante che, applicato alle erogazioni attese, restituisce lo stesso valore attuale."""
    t = np.arange(1, len(pagamenti) + 1)
    lo, hi = -0.5, 1.0
    for _ in range(200):
        m = (lo + hi) / 2
        if (pagamenti / (1 + m) ** t).sum() > pv:
            lo = m
        else:
            hi = m
    return (lo + hi) / 2


def _statistiche(r: DatiRelazione, cb: _Costruttore, proto_didascalia) -> List:
    dip = r.dv.dipendenti
    out = []
    # per societa
    righe = [["Società", "N.", "Età media", "Anzianità media", "Retribuzione media", "TFR medio", "TFR totale"]]
    gruppi: Dict[int, list] = {}
    for e in dip:
        gruppi.setdefault(e.azienda, []).append(e)
    for k in sorted(gruppi) + [0]:
        g = dip if k == 0 else gruppi[k]
        n = len(g)
        righe.append([_nome(r, k) if k else "Totale", numero(n, 0), numero(np.mean([e.eta for e in g]), 1),
                      numero(np.mean([e.anzianita for e in g]), 1), euro(np.mean([e.salario for e in g])),
                      euro(np.mean([e.tfr for e in g])), euro(sum(e.tfr for e in g))])
    out.append(cb.tabella(righe, [4.4, 1.1, 1.4, 1.7, 2.3, 2.1, 2.4], grassetto_righe=[len(righe) - 1]))
    out.append(_didascalia(proto_didascalia, "Caratteristiche del collettivo per società"))
    # eta x anzianita
    n = np.zeros((len(CLASSI_ETA), len(CLASSI_ANZ)))
    sal = np.zeros(len(CLASSI_ETA))
    tfr = np.zeros(len(CLASSI_ETA))
    for e in dip:
        i = next(i for i, c in enumerate(CLASSI_ETA) if c[0] <= e.eta <= c[1])
        j = next(j for j, c in enumerate(CLASSI_ANZ) if c[0] <= e.anzianita <= c[1])
        n[i, j] += 1
        sal[i] += e.salario
        tfr[i] += e.tfr
    righe = [["Età \\ Anzianità"] + [c[2] for c in CLASSI_ANZ] + ["Totale"]]
    for i, c in enumerate(CLASSI_ETA):
        righe.append([c[2]] + [numero(x, 0) if x else "-" for x in n[i]] + [numero(n[i].sum(), 0)])
    righe.append(["Totale"] + [numero(x, 0) for x in n.sum(axis=0)] + [numero(n.sum(), 0)])
    out.append(cb.tabella(righe, [2.6] + [1.4] * len(CLASSI_ANZ) + [1.6], grassetto_righe=[len(righe) - 1]))
    out.append(_didascalia(proto_didascalia, "Distribuzione dei dipendenti per età e anzianità di servizio"))
    righe = [["Classe di età", "N.", "Retribuzione media", "TFR medio"]]
    for i, c in enumerate(CLASSI_ETA):
        m = n[i].sum()
        righe.append([c[2], numero(m, 0), euro(sal[i] / m) if m else "-", euro(tfr[i] / m) if m else "-"])
    m = n.sum()
    righe.append(["Totale", numero(m, 0), euro(sal.sum() / m), euro(tfr.sum() / m)])
    out.append(cb.tabella(righe, [3.5, 2.0, 3.5, 3.5], grassetto_righe=[len(righe) - 1]))
    out.append(_didascalia(proto_didascalia, "Retribuzione media e TFR medio per classe di età"))
    return out


def _tabella7(r: DatiRelazione, cb: _Costruttore):
    mov = r.movimenti
    dc = r.data_valutazione.strftime("%d/%m/%Y")
    if not mov:
        righe = [["Società", "Dipendenti", "CCP", "DBO"]]
        for k in _ordine(r.res):
            x = r.res[k]
            righe.append([_nome(r, k), numero(x.n, 0), euro(x.ccp), euro(x.dbo_metodo(r.metodo))])
        return cb.tabella(righe, [5.4, 2.2, 3.9, 3.9], grassetto_righe=[1])
    dp = r.data_precedente.strftime("%d/%m/%Y") if r.data_precedente else "apertura"
    righe = [["Società", "Passività IAS19 (DBO)", "Service Cost", "Net Interest Expense", "Passività IAS19 (DBO)",
              "Variazioni fondo", "Perdita/Guadagno Attuariale"],
             ["Aziende", dc, "SC", "NIE", dp, "VF", "P/G Att."]]
    colori = {}
    for k in _ordine(mov):
        m = mov[k]
        righe.append([_nome(r, k), euro(m.dbo_chiusura), euro(m.service_cost), euro(m.nie), euro(m.dbo_apertura),
                      euro(m.variazioni_fondo), euro(m.utile)])
    colori[(2, 0)] = "FBD4B4"
    return cb.tabella(righe, [3.1, 2.15, 1.85, 1.85, 2.15, 2.1, 2.2], intestazioni=2, grassetto_righe=[2],
                      colori=colori, dimensione=6.5)


def _tabella_scomposizione(r: DatiRelazione, cb: _Costruttore):
    righe = [["Società", "Utile/(Perdita) attuariale", "da esperienza", "da ipotesi demografiche",
              "da ipotesi finanziarie"]]
    for k in _ordine(r.movimenti):
        m = r.movimenti[k]
        righe.append([_nome(r, k), euro(m.utile), euro(m.utile_esperienza), euro(m.utile_demografiche),
                      euro(m.utile_finanziarie)])
    return cb.tabella(righe, [4.6, 2.7, 2.7, 2.7, 2.7], grassetto_righe=[1], dimensione=7)


def _informativa(r: DatiRelazione, cb: _Costruttore, h1, h2, corpo, didasc) -> List:
    out = [_paragrafo(h1, [("Informativa aggiuntiva ai sensi dello IAS 19", None)])]
    out.append(_paragrafo(corpo, [(
        "Si riportano di seguito le informazioni integrative richieste dai paragrafi 135-147 del principio "
        "IAS 19, riferite al Gruppo nel suo complesso. Il dettaglio per società è riportato nelle tabelle "
        "seguenti e nel file dei risultati della valutazione.", None)]))
    t = r.res[0]
    dbo = t.dbo_metodo(r.metodo)
    dc = data_lunga(r.data_valutazione)

    # riconciliazione
    if r.movimenti:
        m = r.movimenti[0]
        dp = data_lunga(r.data_precedente) if r.data_precedente else "inizio periodo"
        tr = sum(r.trasferimenti.values())
        giorni = (r.data_valutazione - r.data_precedente).days if r.data_precedente else None
        tasso = float(r.ip_precedente.finanziarie.attualizzazione[0]) if r.ip_precedente else None
        out.append(_paragrafo(h2, [("Riconciliazione della passività", None)]))
        testo = (f"La tabella seguente riconcilia la passività al {dp} con quella al {dc}. ")
        if tasso is not None and giorni is not None:
            testo += (f"L'interest cost è calcolato applicando alla passività di apertura il tasso di "
                      f"attualizzazione del primo anno della curva utilizzata nella valutazione precedente "
                      f"({pct(tasso, 4)}) per la frazione di anno trascorsa ({giorni} giorni su 365). ")
        testo += ("Il service cost è pari all'accantonamento contabile del periodo comunicato dalle società. "
                  "Gli utili e le perdite attuariali sono scomposti tra effetto della variazione delle ipotesi "
                  "finanziarie (tassi di attualizzazione, inflazione, incrementi retributivi), effetto della "
                  "variazione delle ipotesi demografiche ed effetto dell'esperienza, cioè degli scostamenti tra "
                  "quanto ipotizzato nella valutazione precedente e quanto effettivamente verificatosi.")
        out.append(_paragrafo(corpo, [(testo, None)]))
        righe = [["Voce", "Importo"],
                 [f"Passività (DBO) al {dp}", euro(m.dbo_apertura)],
                 ["Service cost", euro(m.service_cost)],
                 ["Interest cost (NIE)", euro(m.nie)],
                 ["Benefici erogati e altre variazioni del fondo", euro(m.variazioni_fondo - tr)]]
        if tr:
            righe.append(["Trasferimenti e acquisizioni di società", euro(tr)])
        righe += [["(Utili)/perdite attuariali da ipotesi finanziarie", euro(-m.utile_finanziarie)],
                  ["(Utili)/perdite attuariali da ipotesi demografiche", euro(-m.utile_demografiche)],
                  ["(Utili)/perdite attuariali da esperienza", euro(-m.utile_esperienza)],
                  [f"Passività (DBO) al {dc}", euro(m.dbo_chiusura)]]
        out.append(cb.tabella(righe, [10.0, 4.0], grassetto_righe=[1, len(righe) - 1]))
        out.append(_didascalia(didasc, "Riconciliazione della passività del Gruppo"))

    # ipotesi principali
    fin = r.ip.finanziarie
    pv = float(t.per_anno["pv_pagamenti"].sum())
    equivalente = _tasso_equivalente(t.per_anno["pagamenti"], pv) if pv else float(fin.attualizzazione[0])
    out.append(_paragrafo(h2, [("Principali ipotesi attuariali", None)]))
    righe = [["Ipotesi", "Valore"],
             ["Tasso di attualizzazione", f"curva forward: {pct(float(fin.attualizzazione[0]), 4)} il primo anno; "
                                          f"tasso costante equivalente {pct(equivalente, 2)}"],
             ["Tasso di inflazione", _descrivi_serie(fin.inflazione)],
             ["Tasso di rivalutazione del TFR", _descrivi_serie(fin.tasso_rivalutazione(), 4)
              + " (1,50% + 75% dell'inflazione, al netto dell'imposta sostitutiva)"]]
    for p in _categorie_usate(r):
        righe.append([f"Incremento salariale reale, categoria {_nome_categoria(r, p)}",
                      _descrivi_serie(fin.incremento_reale[p])])
    comp = r.ip.tavola.componenti
    if comp is not None:
        for nome, dati in (("Frequenza annua di dimissione e altre cause", comp[1]),
                           ("Frequenza annua di anticipazione", comp[2])):
            gruppi: Dict[str, List[int]] = {}
            for p in _categorie_usate(r):
                v = dati[(1, p)][ETA_INGRESSO + 1:66]
                s = pct(float(v[0])) if np.allclose(v, v[0]) else f"da {pct(float(v.min()))} a {pct(float(v.max()))} per età"
                gruppi.setdefault(s, []).append(p)
            righe.append([nome, "; ".join(f"{s} (categorie {', '.join(map(str, ps))})" for s, ps in gruppi.items())])
        righe.append(["Mortalità", "tavola RG48 della Ragioneria Generale dello Stato"])
    righe.append(["Duration media della passività", f"{numero(t.duration, 1)} anni"])
    out.append(cb.tabella(righe, [6.0, 9.0], allinea=["l", "l"]))
    out.append(_didascalia(didasc, "Principali ipotesi attuariali"))

    # sensitivity assoluta
    if r.sens_assoluta:
        out.append(_paragrafo(h2, [("Analisi di sensitività sulle ipotesi principali", None)]))
        out.append(_paragrafo(corpo, [(
            "Come richiesto dal paragrafo 145 dello IAS 19, la tabella seguente riporta l'effetto sulla passività "
            "di variazioni ragionevolmente possibili delle ipotesi attuariali principali, ciascuna considerata "
            "separatamente e a parità delle altre. La variazione dell'inflazione si riflette sulla rivalutazione "
            "del TFR e sulle retribuzioni nominali; la variazione del turnover è applicata in valore assoluto "
            "alle frequenze di dimissione.", None)]))
        righe = [["Ipotesi", "DBO", "Variazione", "Variazione %"], ["Valutazione base", euro(dbo), "-", "-"]]
        for nome, v in r.sens_assoluta.items():
            x = v.get(0, 0.0)
            righe.append([_it(nome), euro(x), euro(x - dbo), pct((x - dbo) / dbo) if dbo else "-"])
        out.append(cb.tabella(righe, [6.6, 3.2, 3.0, 2.2], grassetto_righe=[1]))
        out.append(_didascalia(didasc, "Sensitività della passività del Gruppo alle ipotesi principali"))

    # duration, service cost, erogazioni
    out.append(_paragrafo(h2, [("Duration, service cost e erogazioni attese", None)]))
    out.append(_paragrafo(corpo, [(
        "La tabella seguente riporta, per ciascuna società, la duration media finanziaria della passività, il "
        "service cost atteso per l'esercizio successivo (nullo per le società con almeno 50 dipendenti, i cui "
        "accantonamenti sono versati al Fondo di Tesoreria INPS o alla previdenza complementare) e le "
        "erogazioni attese, in valore nominale, per anno di pagamento (importi in euro).", None)]))
    righe = [["Società", "Duration (anni)", "Service cost pro futuro", "Erogazioni anno 1", "Erogazioni anni 2-5",
              "Erogazioni anni 6-10", "Erogazioni oltre 10 anni"]]
    for k in _ordine(r.res):
        x = r.res[k]
        p = x.per_anno["pagamenti"]
        sc = x.sc_fin if r.metodo == "financial" else x.sc_lin
        righe.append([_nome(r, k), numero(x.duration, 1), euro(sc, 0), euro(float(p[0]), 0),
                      euro(float(p[1:5].sum()), 0), euro(float(p[5:10].sum()), 0), euro(float(p[10:].sum()), 0)])
    out.append(cb.tabella(righe, [3.8, 1.5, 2.0, 2.0, 2.0, 2.0, 2.1], grassetto_righe=[1], dimensione=7))
    out.append(_didascalia(didasc, "Duration, service cost pro futuro ed erogazioni attese per società"))
    return out


def _scheda(r: DatiRelazione, k: int, cb: _Costruttore):
    x, d = r.res[k], r.csl[k]
    voci = [("Current Concern Provision", "CCP", x.ccp), ("Difference CCP", "DW", x.dw),
            ("Past Service Liability", "PSL", x.psl), ("Difference CCP: R/A", "DC-R/A", d["DC-R/A"]),
            ("Difference CCP: Infla", "DC-Inf", d["DC-Inf"]), ("Difference CCP: Salari", "DC-Sal", d["DC-Sal"]),
            ("Current Service Liability", "CSL", x.csl), ("Total Difference", "DW+CSL", x.dw + x.csl),
            ("Defined Benefit Obligation", "DBO", x.dbo)]
    etich = "Financial" if r.metodo == "financial" else "Financial (rif.)"
    righe = [[_nome(r, k), "", "Risultati", "", ""], ["", "", etich, "Medie", "Fin/CCP"]]
    for nome, sigla, v in voci:
        righe.append([nome, sigla, numero(v, 0), numero(v / x.n, 0) if x.n else "-",
                      pct(v / x.ccp) if x.ccp else "-"])
    if r.metodo != "financial":
        righe.append(["DBO linear proration", "DBO", numero(x.dbo_lineare, 0),
                      numero(x.dbo_lineare / x.n, 0) if x.n else "-", pct(x.dbo_lineare / x.ccp) if x.ccp else "-"])
    colori = {(2, 2): ARANCIO, (10, 2): ARANCIO}
    return cb.tabella(righe, [5.0, 2.4, 2.6, 1.8, 1.9], intestazioni=2, allinea=["l", "c", "r", "r", "r"],
                      colori=colori, unioni=[(0, 0, 1, 1), (0, 2, 0, 4)], dimensione=8)


def _grafico(r: DatiRelazione, k: int) -> bytes:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    x = r.res[k]
    pa = x.per_anno
    csl = pa["c2"] - pa["c1"]
    dbo_anno = pa["w2"] + csl
    ccp_res = np.concatenate([[x.ccp], x.ccp - np.cumsum(pa["w1"])])
    dbo_res = np.concatenate([[x.dbo], x.dbo - np.cumsum(dbo_anno)])
    attivi = np.nonzero(np.abs(pa["w1"]) + np.abs(dbo_anno) > 0.5)[0]
    n = int(min(max(attivi[-1] + 1 if len(attivi) else 1, 10), 30))
    scala, unita = (1e6, "milioni di euro") if max(x.ccp, x.dbo) >= 2e6 else (1e3, "migliaia di euro")

    fig, (a1, a2) = plt.subplots(2, 1, figsize=(6.3, 7.6), dpi=200)
    for ax in (a1, a2):
        ax.grid(True, color="#d9d9d9", linewidth=0.6)
        ax.set_axisbelow(True)
        for s in ("top", "right"):
            ax.spines[s].set_visible(False)
        for s in ("left", "bottom"):
            ax.spines[s].set_color("#8c8c8c")
        ax.tick_params(colors="#404040", labelsize=8)
        ax.set_ylabel(unita, fontsize=8, color="#404040")
        ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: numero(v, 2)))
    t0 = np.arange(0, n + 1)
    a1.plot(t0, ccp_res[: n + 1] / scala, color="#" + COL_CCP, linewidth=2, label="TFR contabile residuo (CCP)")
    a1.plot(t0, dbo_res[: n + 1] / scala, color="#" + COL_DBO, linewidth=2, label="Passività IAS 19 residua (DBO)")
    a1.set_title("Andamento prospettico del fondo", fontsize=9, loc="left", color="#202020")
    a1.set_xlim(0, n)
    a1.set_ylim(bottom=min(0, float(min(ccp_res[: n + 1].min(), dbo_res[: n + 1].min())) / scala))
    a1.legend(fontsize=8, frameon=False)
    t1 = np.arange(1, n + 1)
    a2.plot(t1, pa["w1"][:n] / scala, color="#" + COL_CCP, linewidth=2, label="TFR contabile erogato (CCP)")
    a2.plot(t1, dbo_anno[:n] / scala, color="#" + COL_DBO, linewidth=2, label="Quota di DBO per anno di uscita")
    a2.set_title("Variazioni periodiche", fontsize=9, loc="left", color="#202020")
    a2.set_xlim(1, n)
    a2.set_xlabel("anni dalla data di valutazione", fontsize=8, color="#404040")
    a2.legend(fontsize=8, frameon=False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return buf.getvalue()


def _tabella_sensitivity(r: DatiRelazione, k: int, cb: _Costruttore, relativa: bool):
    g = r.sens[k]
    base = g[1, 1]
    v = g / base if (relativa and base) else g
    fmt = (lambda z: pct(z)) if relativa else (lambda z: numero(z, 0))
    righe = [[_nome(r, k), "", "Attualizzazione", "", "", "", ""],
             ["", "", "-10%", "", "100%", "", "+10%"]]
    etich = ["-10%", "", "100%", "", "+10%"]
    colori = {}
    for i in range(5):
        riga = ["Infla" if i == 0 else "", etich[i]]
        for j in range(5):
            gi, gj = i // 2, j // 2
            if i % 2 == 0 and j % 2 == 0:
                riga.append(fmt(v[gi, gj]))
            else:
                if i % 2 == 0:          # differenza orizzontale: colonna esterna - colonna centrale
                    z = v[gi, 0 if j == 1 else 2] - v[gi, 1]
                elif j % 2 == 0:        # differenza verticale: riga esterna - riga centrale
                    z = v[0 if i == 1 else 2, gj] - v[1, gj]
                else:                   # diagonale: angolo - base
                    z = v[0 if i == 1 else 2, 0 if j == 1 else 2] - v[1, 1]
                riga.append(fmt(z))
                colori[(i + 2, j + 2)] = GIALLO
        righe.append(riga)
    for j in range(7):
        colori.setdefault((0, j), ARANCIO)
        colori.setdefault((1, j), ARANCIO)
    for i in range(2, 7):
        colori[(i, 0)] = ARANCIO
        colori[(i, 1)] = ARANCIO
    colori[(0, 0)] = ROSSO
    return cb.tabella(righe, [1.3, 1.6, 2.1, 2.0, 2.1, 2.0, 2.1], intestazioni=2,
                      allinea=["c"] * 7, colori=colori,
                      unioni=[(0, 0, 1, 1), (0, 2, 0, 6), (2, 0, 6, 0)], dimensione=8)


def _it(testo: str) -> str:
    return re.sub(r"(\d)\.(\d)", r"\1,\2", testo)


def _tabella_sens_assoluta_societa(r: DatiRelazione, cb: _Costruttore):
    scen = list(r.sens_assoluta)
    brevi = [_it(s).replace("Tasso di attualizzazione", "Attual.").replace("Tasso di inflazione", "Infl.")
             .replace("Tasso di turnover", "Turnover") for s in scen]
    righe = [["Società", "DBO base"] + brevi]
    for k in _ordine(r.res):
        righe.append([_nome(r, k), numero(r.res[k].dbo_metodo(r.metodo), 0)]
                     + [numero(r.sens_assoluta[s].get(k, 0.0), 0) for s in scen])
    larg = [3.4, 1.8] + [(LARGHEZZA_CM - 5.2) / len(scen)] * len(scen)
    return cb.tabella(righe, larg, grassetto_righe=[1], dimensione=6.5)


# ---------------------------------------------------------------- composizione

def _trova(body, cond, dopo=None):
    vedi = dopo is None
    for el in body:
        if not vedi:
            vedi = el is dopo
            continue
        if cond(el):
            return el
    return None


def _stile(el) -> str:
    qn = _ns()
    ppr = el.find(qn("w:pPr"))
    if ppr is None or ppr.find(qn("w:pStyle")) is None:
        return ""
    return ppr.find(qn("w:pStyle")).get(qn("w:val"))


def _e_p(el) -> bool:
    return el.tag == _ns()("w:p")


def _e_tbl(el) -> bool:
    return el.tag == _ns()("w:tbl")


def _tabella_prima_di(didascalia):
    el = didascalia.getprevious()
    while el is not None and not _e_tbl(el):
        el = el.getprevious()
    return el


def scrivi_relazione(modello: str, uscita: str, r: DatiRelazione) -> None:
    import docx
    from docx.oxml import OxmlElement

    qn = _ns()
    doc = docx.Document(modello)
    body = doc.element.body
    cb = _Costruttore(doc)

    # --- cliente e data (proprieta del documento collegate ai campi del modello)
    cp = doc.core_properties
    vecchio_cliente = (cp.keywords or "").strip()
    vecchia_data = (cp.comments or "").strip()
    nuova_data = data_lunga(r.data_valutazione)
    cp.keywords = r.cliente
    cp.comments = nuova_data
    parti = [body] + [h._element for s in doc.sections for h in
                      (s.header, s.footer, s.first_page_header, s.first_page_footer, s.even_page_header,
                       s.even_page_footer)]
    for parte in parti:
        for p in parte.iter(qn("w:p")):
            _sostituisci(p, vecchio_cliente, r.cliente)
            _sostituisci(p, vecchia_data, nuova_data)

    # --- stato del documento: data di emissione
    emissione = (r.data_emissione or date.today()).strftime("%d/%m/%Y")
    stato = _trova(body, lambda el: _e_tbl(el) and _testo(el).startswith("Versione"))
    if stato is not None:
        tr = stato.findall(qn("w:tr"))[1]
        for t in tr.iter(qn("w:t")):
            if re.fullmatch(r"\d{2}/\d{2}/\d{4}", t.text or ""):
                t.text = emissione

    # --- prototipi (solo nel corpo della relazione, dopo sommario e indici)
    def paragrafi_corpo():
        tutti = [el for el in body if _e_p(el)]
        i0 = next(i for i, p in enumerate(tutti) if _stile(p) in ("Titolo1", "Heading1"))
        return tutti[i0:]
    paragrafi = paragrafi_corpo()
    h1 = paragrafi[0]
    def sottolivello(p):
        il = p.find(".//" + qn("w:ilvl"))
        return _stile(p) in ("Titolo1", "Heading1") and il is not None and il.get(qn("w:val")) == "1"
    h2 = next((p for p in paragrafi if sottolivello(p)), None)
    if h2 is None:
        h2 = next(p for p in paragrafi if _stile(p) in ("Titolo2", "Heading2"))
    corpo = next(p for p in paragrafi if _testo(p).startswith("La valutazione in esame"))
    didasc_t = next(p for p in paragrafi if _testo(p).startswith("Tabella") and "Risultati della valutazione" in _testo(p))
    didasc_f = next(p for p in paragrafi if _testo(p).startswith("Figura"))
    proto_didasc_t = copy.deepcopy(didasc_t)
    proto_didasc_f = copy.deepcopy(didasc_f)

    # --- Tabelle 1-6
    def per_didascalia(frammento, *altri):
        for p in paragrafi:
            t = _testo(p).strip()
            if t.startswith("Tabella") and frammento in t and all(a in t for a in altri):
                return _tabella_prima_di(p)
        return None
    tabelle = {
        "rg48": per_didascalia("Tavola di sopravvivenza RG48"),
        "freq_m": per_didascalia("Probabilità demografica", "(Maschi)"),
        "lx_m": per_didascalia("tre cause di eliminazione", "(Maschi)"),
        "freq_f": per_didascalia("Probabilità demografica", "(Femmine)"),
        "lx_f": per_didascalia("tre cause di eliminazione", "(Femmine)"),
        "tassi": per_didascalia("Tassi annui di inflazione"),
    }
    _tabelle_ipotesi(r, {k: v for k, v in tabelle.items() if v is not None})

    # --- incrementi salariali per categoria
    cat_par = [p for p in paragrafi if _stile(p) == "Paragrafoelenco" and _testo(p).startswith("Categoria")]
    if cat_par:
        nuovi = []
        for p in _categorie_usate(r):
            g = r.ip.finanziarie.incremento_reale[p]
            nuovi.append(_paragrafo(cat_par[0], [(f"Categoria {_nome_categoria(r, p)}\t\t\t"
                                                  f"tasso annuo reale {_descrivi_serie(g)}", None)]))
        _inserisci_dopo(cat_par[-1], nuovi)
        for p in cat_par:
            _rimuovi(p)

    # --- numero dipendenti ed elenco societa
    n_tot = r.res[0].n
    for p in paragrafi:
        if "i calcoli hanno riguardato n." in _testo(p):
            m = re.search(r"n\. ([\d\.]+) dipendenti", _testo(p))
            if m:
                _sostituisci(p, m.group(0), f"n. {numero(n_tot, 0)} dipendenti")
    soc_par = [p for p in paragrafi if re.search(r"dipendenti n\. \d+$", _testo(p).strip())]
    if soc_par:
        nuovi = [_paragrafo(soc_par[0], [(f"{_nome(r, k)}\t\t\tdipendenti n. {numero(r.res[k].n, 0)}", None)])
                 for k in sorted(k for k in r.res if k)]
        _inserisci_dopo(soc_par[-1], nuovi)
        for p in soc_par:
            _rimuovi(p)

    # --- caratteristiche del collettivo
    ancora = next((p for p in paragrafi if _testo(p).startswith("Al fine di esaminare la struttura")), None)
    if ancora is not None:
        _inserisci_dopo(ancora, [OxmlElement("w:p")] + _statistiche(r, cb, proto_didasc_t))

    # --- definizione degli utili/perdite attuariali dopo il NIE
    nie = next((p for p in paragrafi if _testo(p).startswith("Net Interest Expense")), None)
    if nie is not None:
        _inserisci_dopo(nie, [_paragrafo(nie, [
            ("Utili e perdite attuariali: ", None),
            ("variazioni della passività dovute alla modifica delle ipotesi finanziarie (tassi di attualizzazione, "
             "inflazione, incrementi retributivi), alla modifica delle ipotesi demografiche (dimissioni, "
             "anticipazioni, mortalità) e agli scostamenti tra le ipotesi della valutazione precedente e quanto "
             "effettivamente verificatosi nel periodo (esperienza). Sono rilevati tra le altre componenti del "
             "conto economico complessivo (§120 e §127).", None)])])

    # --- Tabella 7 e scomposizione degli utili/perdite
    img7 = didasc_t.getprevious()
    while img7 is not None and not (_e_p(img7) and img7.findall(".//" + qn("w:drawing"))):
        img7 = img7.getprevious()
    t7 = _tabella7(r, cb)
    if img7 is not None:
        img7.addprevious(t7)
        _rimuovi(img7)
    else:
        didasc_t.addprevious(t7)
    ultimo = didasc_t
    if r.movimenti:
        ultimo = _inserisci_dopo(ultimo, [
            OxmlElement("w:p"),
            _paragrafo(corpo, [(
                "Gli utili e le perdite attuariali della Tabella precedente sono scomposti, per ciascuna società, "
                "tra effetto dell'esperienza, delle ipotesi demografiche e delle ipotesi finanziarie (valori "
                "positivi = utili, cioè riduzione della passività):", None)]),
            _tabella_scomposizione(r, cb),
            _didascalia(proto_didasc_t, "Scomposizione degli utili e delle perdite attuariali"),
        ])
    # informativa aggiuntiva dopo la Tabella 7, prima della firma
    _inserisci_dopo(ultimo, [OxmlElement("w:p")] + _informativa(r, cb, h1, h2, corpo, proto_didasc_t)
                    + [OxmlElement("w:p")])

    # --- allegati: una scheda e un grafico per societa
    paragrafi = paragrafi_corpo()
    inizio = next((p for p in paragrafi if _testo(p).startswith("Scheda di valutazione attuariale al")), None)
    sens_titolo = next((p for p in paragrafi if _testo(p).strip() == "Analisi di sensitività"), None)
    if inizio is not None and sens_titolo is not None:
        proto_h = copy.deepcopy(inizio)
        defs = []
        el = inizio
        while el is not None and el is not sens_titolo:
            t = _testo(el)
            if _e_p(el) and t.startswith("(") and len(defs) < 5:
                defs.append(copy.deepcopy(el))
            el = el.getnext()
        da_togliere = []
        el = inizio
        while el is not None and el is not sens_titolo:
            da_togliere.append(el)
            el = el.getnext()
        blocchi = []
        for i, k in enumerate(_ordine(r.res)):
            nome = _nome(r, k)
            if i:
                blocchi.append(_interruzione_pagina())
            blocchi.append(_paragrafo(proto_h, [(
                f"Scheda di valutazione attuariale al {nuova_data} del Fondo di Trattamento Fine Rapporto sulla "
                f"base del principio contabile IAS 19 – {nome}", None)]))
            blocchi.append(OxmlElement("w:p"))
            blocchi.append(_scheda(r, k, cb))
            blocchi.append(_didascalia(proto_didasc_t, f"Scheda di valutazione attuariale {nome}"))
            for proto, (sigla, testo) in zip(defs, DEFINIZIONI):
                blocchi.append(_paragrafo(proto, [(f"{sigla}{': ' + testo if testo else ''}", None)]))
            blocchi.append(OxmlElement("w:p"))
            blocchi.append(cb.immagine(_grafico(r, k), 14.0))
            blocchi.append(_didascalia(proto_didasc_f,
                                       f"Andamento prospettico del Fondo di Trattamento Fine Rapporto e delle "
                                       f"variazioni periodiche in base alla valutazione al {nuova_data} – {nome}"))
        blocchi.append(_interruzione_pagina())
        inizio.addprevious(blocchi[0])
        _inserisci_dopo(blocchi[0], blocchi[1:])
        for el in da_togliere:
            _rimuovi(el)

    # --- sensitivity 3x3 per societa
    if sens_titolo is not None:
        intro = sens_titolo.getnext()
        while intro is not None and not (_e_p(intro) and _testo(intro).strip()):
            intro = intro.getnext()
        if intro is not None:
            _sostituisci(intro, "nei tassi di rivalutazione e/o di attualizzazione",
                         "del tasso di inflazione (e quindi dei tassi di rivalutazione e di incremento retributivo "
                         "nominale) e/o dei tassi di attualizzazione")
            el = intro.getnext()
            fine = body.find(qn("w:sectPr"))
            while el is not None and el is not fine:
                nxt = el.getnext()
                if not (_e_p(el) and el.find(".//" + qn("w:sectPr")) is not None):
                    _rimuovi(el)
                el = nxt
            nuovi = [_paragrafo(corpo, [(
                "Nelle celle evidenziate sono riportate le differenze tra valori adiacenti della griglia: tra le "
                "colonne, l'effetto della sola variazione dei tassi di attualizzazione; tra le righe, l'effetto "
                "della sola variazione dell'inflazione; nelle celle diagonali, la differenza rispetto alla "
                "valutazione base.", None)])]
            for k in _ordine(r.sens):
                nome = _nome(r, k)
                nuovi += [OxmlElement("w:p"), _tabella_sensitivity(r, k, cb, False),
                          _didascalia(proto_didasc_t, f"Sensitività (valori assoluti) – {nome}"),
                          OxmlElement("w:p"), _tabella_sensitivity(r, k, cb, True),
                          _didascalia(proto_didasc_t, f"Sensitività (valori relativi) – {nome}")]
            if r.sens_assoluta:
                nuovi += [_interruzione_pagina(), _paragrafo(corpo, [(
                    "La tabella seguente riporta, per ciascuna società, la passività ricalcolata con variazioni "
                    "assolute delle ipotesi principali (tassi di attualizzazione e di inflazione, frequenze di "
                    "turnover), ciascuna considerata separatamente.", None)]),
                    _tabella_sens_assoluta_societa(r, cb),
                    _didascalia(proto_didasc_t, "Sensitività con variazioni assolute delle ipotesi, per società")]
            _inserisci_dopo(intro, nuovi)

    # --- numerazione didascalie e aggiornamento di sommario e indici all'apertura
    _numera_didascalie(body)
    settings = doc.settings.element
    for u in settings.findall(qn("w:updateFields")):
        settings.remove(u)
    u = OxmlElement("w:updateFields")
    u.set(qn("w:val"), "true")
    seguenti = {qn(t) for t in ("w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
                                "w:rsids", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
                                "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures",
                                "w:forceUpgrade", "w:captions", "w:readModeInkLockDown", "w:smartTagType",
                                "w:schemaLibrary", "w:shapeDefaults", "w:doNotEmbedSmartTags",
                                "w:decimalSymbol", "w:listSeparator")}
    seguenti.add("{http://schemas.openxmlformats.org/officeDocument/2006/math}mathPr")
    dopo = next((c for c in settings if c.tag in seguenti), None)
    if dopo is not None:
        dopo.addprevious(u)
    else:
        settings.append(u)
    doc.save(uscita)
