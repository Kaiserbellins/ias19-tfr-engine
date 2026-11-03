"""Nota per il cliente sulla movimentazione del TFR ai fini dell'IFRS 18, dal template documentale.

Il template (es. "Relazione-Cliente.docx") fornisce copertina, metadati, tabella delle revisioni,
sommario, intestazioni e stili; i paragrafi di esempio del corpo ("Titolo 1", "Sed ut
perspiciatis...") vengono sostituiti dal contenuto generato dal prospetto IFRS 18.
"""
from __future__ import annotations

import copy
import os
import re
from datetime import date
from typing import List, Mapping, Optional, Sequence

from .ifrs18 import Periodo, colonne
from .relazione import (_e_p, _inserisci_dopo, _ns, _paragrafo, _rigenera_indici, _rimuovi, _sostituisci,
                        _stile, _testo, euro, pct)

BLU = "1328C9"
GRIGIO = "F2F2F2"
LARGHEZZA_CM = 17.0

RIGHE = [
    ("dbo_apertura", "Passività (DBO) di inizio periodo", True),
    ("service_cost", "Service cost (categoria operativa)", False),
    ("nie", "Net interest expense (categoria di finanziamento)", False),
    ("ce", "Totale a conto economico", True),
    ("perdite_finanziarie", "(Utili)/perdite attuariali: ipotesi finanziarie", False),
    ("perdite_demografiche", "(Utili)/perdite attuariali: ipotesi demografiche", False),
    ("perdite_esperienza", "(Utili)/perdite attuariali: esperienza", False),
    ("oci", "Totale altre componenti del conto economico complessivo", True),
    ("variazioni_fondo", "Benefici pagati, trasferimenti e altre variazioni", False),
    ("dbo_chiusura", "Passività (DBO) di fine periodo", True),
]


def _tabella(doc, righe: Sequence[Sequence[str]], larghezze: Sequence[float], grassetto: Sequence[int] = (),
             allinea: Optional[Sequence[str]] = None, dimensione: float = 8):
    """Tabella nello stile della tabella delle revisioni del template: intestazione blu, testo bianco."""
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.shared import Cm, Pt, RGBColor
    qn = _ns()
    ncol = len(larghezze)
    t = doc.add_table(rows=len(righe), cols=ncol)
    stili = [s.name for s in doc.styles if s.type == 3]
    if "Table Grid" in stili:
        t.style = doc.styles["Table Grid"]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    allinea = allinea or (["l"] + ["r"] * (ncol - 1))
    mappa = {"l": WD_ALIGN_PARAGRAPH.LEFT, "r": WD_ALIGN_PARAGRAPH.RIGHT, "c": WD_ALIGN_PARAGRAPH.CENTER}
    for i, riga in enumerate(righe):
        for j in range(ncol):
            cell = t.cell(i, j)
            cell.width = Cm(larghezze[j])
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.space_before = Pt(1)
            pf.space_after = Pt(1)
            p.alignment = mappa["c" if i == 0 and j else allinea[j]]
            run = p.add_run(riga[j] if j < len(riga) else "")
            run.font.size = Pt(dimensione)
            fill = None
            if i == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                fill = BLU
            elif i in grassetto:
                run.bold = True
                fill = GRIGIO
            if fill:
                tcpr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), fill)
                tcpr.append(shd)
    el = t._tbl
    grid = el.find(qn("w:tblGrid"))
    for g in list(grid):
        grid.remove(g)
    for w in larghezze:
        g = OxmlElement("w:gridCol")
        g.set(qn("w:w"), str(int(w * 567)))
        grid.append(g)
    trs = el.findall(qn("w:tr"))
    for tr in trs[:-1]:   # tabella non spezzata tra pagine (keepNext su tutte le righe tranne l'ultima)
        for par in tr.iter(qn("w:p")):
            ppr = par.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                par.insert(0, ppr)
            ppr.insert(0, OxmlElement("w:keepNext"))
    for i, tr in enumerate(trs):
        trpr = tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
        if i == 0:
            trpr.append(OxmlElement("w:tblHeader"))
    el.getparent().remove(el)
    return el


def _tabella_movimenti(doc, periodi, k):
    cols = colonne(periodi, k)
    righe = [["Voce"] + [c[0] for c in cols]]
    tot = []
    for chiave, descr, grassetto in RIGHE:
        righe.append([descr] + [euro(c[1][chiave]) for c in cols])
        if grassetto:
            tot.append(len(righe) - 1)
    n = len(cols)
    return _tabella(doc, righe, [LARGHEZZA_CM - 3.2 * n] + [3.2] * n, grassetto=tot)


def scrivi_nota(modello: str, uscita: str, periodi: Sequence[Periodo], nomi: Mapping[int, str], cliente: str,
                autore: str = "", societa_autore: str = "Eurokleis", data_emissione: Optional[date] = None,
                tasso_nie: Optional[float] = None, note: Sequence[str] = ()) -> None:
    import docx
    from docx.oxml import OxmlElement

    qn = _ns()
    periodi = sorted(periodi, key=lambda p: p.fine)
    anno = periodi[-1].fine.year
    apertura = f"31/12/{anno - 1}"
    emissione = (data_emissione or date.today()).strftime("%d/%m/%Y")
    titolo = f"TFR IAS 19 – Movimentazione {anno} per la transizione all'IFRS 18"
    doc = docx.Document(modello)
    body = doc.element.body

    # --- copertina, intestazioni e metadati
    parti = [body] + [h._element for s in doc.sections for h in
                      (s.header, s.footer, s.first_page_header, s.first_page_footer)]
    for parte in parti:
        for p in parte.iter(qn("w:p")):
            _sostituisci(p, "[TITOLO DOCUMENTO]", titolo)
            if autore:
                _sostituisci(p, "[digitare Autore/i del documento]", autore)
            _sostituisci(p, "[MAX 5], ", f"TFR, IAS 19, IFRS 18, {cliente}")
            _sostituisci(p, "[MAX 5]", f"TFR, IAS 19, IFRS 18, {cliente}")
    # segnaposto MACROBUTTON della copertina ("[digitare Autore/i del documento]")
    if autore:
        from .relazione import _primo_rpr, _run
        for it in list(body.iter(qn("w:instrText"))):
            if "MACROBUTTON" in (it.text or "") and "Autore" in (it.text or ""):
                par = it.getparent().getparent()
                rpr = _primo_rpr(par)
                for r in par.findall(qn("w:r")):
                    par.remove(r)
                par.append(_run(autore, rpr))
    cp = doc.core_properties
    cp.title = titolo
    cp.subject = f"{cliente} - movimentazione TFR {anno}"
    cp.keywords = f"TFR, IAS 19, IFRS 18, {cliente}"
    if autore:
        cp.author = autore
    nome_file = os.path.basename(uscita)
    for r in body.iter(qn("w:instrText")):
        if "FILENAME" in (r.text or ""):
            run = r.getparent()
            el = run.getnext()
            while el is not None:
                fc = el.find(qn("w:fldChar"))
                if fc is not None and fc.get(qn("w:fldCharType")) == "end":
                    break
                for t in el.findall(qn("w:t")):
                    t.text = os.path.splitext(nome_file)[0] if "-AU" in _testo(run.getparent()) else nome_file
                el = el.getnext()

    # elenchi dei metadati (Autori, Gruppo di lavoro, Revisione)
    paragrafi = [el for el in body if _e_p(el)]

    def elenco_dopo(etichetta):
        i = next((n for n, p in enumerate(paragrafi) if _testo(p).strip() == etichetta), None)
        out = []
        if i is None:
            return out
        for p in paragrafi[i + 1:]:
            if not _testo(p).strip():
                break
            out.append(p)
        return out

    if autore:
        aut = elenco_dopo("Autori")
        if aut:
            _inserisci_dopo(aut[-1], [_paragrafo(aut[0], [(autore, None)])])
            for p in aut:
                _rimuovi(p)
        gru = elenco_dopo("Gruppo di lavoro")
        if gru:
            _inserisci_dopo(gru[-1], [_paragrafo(gru[0], [(f"{autore} / {societa_autore}", None)]),
                                      _paragrafo(gru[0], [(f"{cliente} / Cliente", None)])])
            for p in gru:
                _rimuovi(p)
    for p in elenco_dopo("Revisione / Approvazione"):
        t = _testo(p)
        etich = t.split(":")[0]
        p.addprevious(_paragrafo(p, [(f"{etich}:\t[da compilare]", None)]))
        _rimuovi(p)
    rev = next((t for t in body.iter(qn("w:tbl")) if _testo(t).startswith("Ver.")), None)
    if rev is not None:
        trs = rev.findall(qn("w:tr"))
        if len(trs) > 1:
            celle = trs[1].findall(qn("w:tc"))
            valori = ["1.0", emissione, "Prima emissione del documento", autore or ""]
            for tc, v in zip(celle, valori):
                ts = list(tc.iter(qn("w:t")))
                if ts:
                    ts[0].text = v
                    for t in ts[1:]:
                        t.text = ""

    # --- corpo: prototipi dal testo di esempio, poi sostituzione
    paragrafi = [el for el in body if _e_p(el)]
    h1 = next(p for p in paragrafi if _stile(p) in ("Titolo1", "Heading1"))
    h2 = next((p for p in paragrafi if _stile(p) in ("Titolo2", "Heading2")), h1)
    occhiello = next((p for p in paragrafi if _stile(p) in ("Sottotitolo", "Subtitle")), None)
    i_h1 = paragrafi.index(h1)
    corpo = next(p for p in paragrafi[i_h1 + 1:] if _stile(p) == "" and len(_testo(p)) > 50)
    proto = {k: copy.deepcopy(v) for k, v in (("h1", h1), ("h2", h2), ("corpo", corpo),
                                               ("occhiello", occhiello)) if v is not None}
    for k in ("h1", "h2"):
        for b in list(proto[k].iter(qn("w:bookmarkStart"))) + list(proto[k].iter(qn("w:bookmarkEnd"))):
            b.getparent().remove(b)
    # i run del paragrafo di esempio possono essere in grassetto: si usa il primo run
    sect = body.find(qn("w:sectPr"))
    el = h1
    while el is not None and el is not sect:
        nxt = el.getnext()
        _rimuovi(el)
        el = nxt

    def P(testo):
        return _paragrafo(proto["corpo"], [(testo, None)])

    def H1(testo):
        return _paragrafo(proto["h1"], [(testo, None)])

    def H2(testo):
        return _paragrafo(proto["h2"], [(testo, None)])

    def occh(testo):
        if "occhiello" not in proto:
            return P(testo)
        q = copy.deepcopy(proto["occhiello"])
        runs = q.findall(qn("w:r"))
        # primo run = punto colorato del template; il testo prende lo stile dell'ultimo run
        for r in runs[1:-1]:
            q.remove(r)
        ultimo_run = runs[-1]
        for t in ultimo_run.findall(qn("w:t")):
            ultimo_run.remove(t)
        t = OxmlElement("w:t")
        t.text = " " + testo
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        ultimo_run.append(t)
        return q

    vuoto = lambda: OxmlElement("w:p")  # noqa: E731
    g = {c[0]: c[1] for c in colonne(periodi, 0)}
    ultimo = list(g.values())[-1]
    etich_ultimo = list(g)[-1]
    date_val = ", ".join(f"{p.fine:%d/%m/%Y}" for p in periodi)
    contenuto: List = [
        H1("Premessa"),
        occh(f"{cliente} – TFR al " + (", ".join(f"{p.fine:%d/%m}" for p in periodi[:-1]) + " e " if len(periodi) > 1 else "")
             + f"{periodi[-1].fine:%d/%m/%Y}"),
        P(f"L'IFRS 18 «Presentazione e informativa nel bilancio» sostituisce lo IAS 1 a partire dagli esercizi che "
          f"iniziano il 1° gennaio 2027. Il principio si applica retroattivamente: il conto economico dell'esercizio "
          f"comparativo {anno} deve essere ripresentato secondo le nuove categorie e, nei bilanci intermedi "
          f"{anno + 1}, i dati dei corrispondenti periodi {anno} devono essere riesposti."),
        P(f"La presente nota riporta, per il Gruppo e per ciascuna società, la movimentazione della passività per "
          f"Trattamento di Fine Rapporto valutata secondo lo IAS 19 alle date del {date_val}, con le componenti "
          f"di costo classificate secondo l'IFRS 18. Ogni periodo parte dalla passività al {apertura}."),
        vuoto(),
        H1("Classificazione delle componenti secondo l'IFRS 18"),
        P("Le componenti della movimentazione del TFR sono presentate come segue:"),
        _tabella(doc, [["Componente", "Prospetto", "Categoria IFRS 18"],
                       ["Service cost", "Conto economico", "Operativa"],
                       ["Net interest expense (interest cost)", "Conto economico", "Di finanziamento"],
                       ["Utili e perdite attuariali", "Altre componenti del conto economico complessivo",
                        "Non riclassificabili a conto economico"],
                       ["Benefici pagati, trasferimenti e altre variazioni", "Stato patrimoniale", "-"]],
                 [5.0, 6.0, 6.0], allinea=["l", "l", "l"]),
        vuoto(),
        P("Il net interest expense è un interesse su una passività che non deriva da operazioni di sola raccolta "
          "di finanziamenti e per questo rientra nella categoria di finanziamento; il service cost, costo del "
          "personale, resta nella categoria operativa. Il trattamento degli utili e delle perdite attuariali non "
          "cambia rispetto allo IAS 19."),
        vuoto(),
        H1("Movimentazione della passività del Gruppo"),
        _tabella_movimenti(doc, periodi, 0),
        vuoto(),
        P(f"Nel periodo chiuso al {etich_ultimo.split(' ')[0]} il conto economico accoglie oneri per "
          f"{euro(ultimo['ce'])} euro, di cui {euro(ultimo['service_cost'])} euro nella categoria operativa e "
          f"{euro(ultimo['nie'])} euro nella categoria di finanziamento. Le altre componenti del conto economico "
          f"complessivo registrano {'un utile' if ultimo['oci'] < 0 else 'una perdita'} attuariale di "
          f"{euro(abs(ultimo['oci']))} euro"
          + (f", dovut{'o' if ultimo['oci'] < 0 else 'a'} principalmente alla variazione delle ipotesi finanziarie"
             if abs(ultimo["perdite_finanziarie"]) > abs(ultimo["perdite_esperienza"]) else
             ", dovut{} principalmente all'esperienza".format('o' if ultimo['oci'] < 0 else 'a'))
          + f". I benefici pagati e le altre variazioni del fondo ammontano a {euro(ultimo['variazioni_fondo'])} euro."),
        vuoto(),
        H1("Movimentazione per società"),
    ]
    chiavi = sorted({k for p in periodi for k in p.movimenti if k})
    for k in chiavi:
        contenuto += [H2(nomi.get(k, f"Società {k}")), _tabella_movimenti(doc, periodi, k), vuoto()]
    testo_nie = (f"al tasso forward del primo anno della valutazione al {apertura} ({pct(tasso_nie, 4)}), "
                 if tasso_nie is not None else
                 f"al tasso forward del primo anno della valutazione al {apertura}, ")
    contenuto += [
        H1("Note metodologiche"),
        P("Le passività sono valutate secondo lo IAS 19 con il metodo della proiezione unitaria del credito "
          "(Projected Unit Credit Method), nella versione Financial Proration che tiene conto delle anticipazioni "
          "e dei versamenti alla previdenza complementare."),
        P(f"Il net interest expense è calcolato sulla passività al {apertura} {testo_nie}per i giorni trascorsi "
          f"su 365. Il service cost e le variazioni del fondo sono quelli comunicati dalle società."),
        P("Gli utili e le perdite attuariali sono scomposti tra effetto della variazione delle ipotesi finanziarie "
          "(tassi di attualizzazione, inflazione, incrementi retributivi), delle ipotesi demografiche ed effetto "
          "dell'esperienza (scostamento tra quanto ipotizzato e quanto effettivamente verificatosi)."),
    ] + [P(n) for n in note]
    corpo_stile = _stile(proto["corpo"])
    spaziato = []
    for i, el in enumerate(contenuto):
        spaziato.append(el)
        succ = contenuto[i + 1] if i + 1 < len(contenuto) else None
        if (succ is not None and _e_p(el) and _e_p(succ) and _testo(el).strip() and _testo(succ).strip()
                and _stile(el) == corpo_stile and _stile(succ) == corpo_stile):
            spaziato.append(OxmlElement("w:p"))
    for el in spaziato:
        sect.addprevious(el)

    _rigenera_indici(body)
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        u = OxmlElement("w:updateFields")
        u.set(qn("w:val"), "true")
        dopo = next((c for c in settings if c.tag in {qn("w:hdrShapeDefaults"), qn("w:footnotePr"),
                                                       qn("w:endnotePr"), qn("w:compat"), qn("w:docVars"),
                                                       qn("w:rsids")}), None)
        (dopo.addprevious(u) if dopo is not None else settings.append(u))
    doc.save(uscita)
