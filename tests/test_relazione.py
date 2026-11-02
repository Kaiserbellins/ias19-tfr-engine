"""Relazione Word.

I test sintetici controllano formati e sostituzioni nel testo. Il test di integrazione genera la
relazione Garofalo dal modello GHC e richiede (i documenti dei clienti non sono nel repository):
    IAS19_GAROFALO_XLSX       valutazione al 30/09/2026
    IAS19_RELAZIONE_DOCX      relazione GHC al 30/09/2026 usata come modello
"""
import os
import re
from datetime import date

import pytest
from docx.oxml import parse_xml

from ias19tfr.relazione import _numera_didascalie, _sostituisci, _testo, data_lunga, euro, numero, pct

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def test_formati_italiani():
    assert numero(13352807.064) == "13.352.807,06"
    assert numero(-604.33) == "-604,33"
    assert euro(0.001) == "-"
    assert pct(0.0207528, 4) == "2,0753%"
    assert data_lunga(date(2026, 9, 30)) == "30 settembre 2026"


def test_sostituzione_su_piu_run():
    p = parse_xml(f'<w:p {W}><w:r><w:t>Garofalo Health</w:t></w:r><w:r><w:t>care S.p.A. al 30 </w:t></w:r>'
                  f'<w:r><w:t>settembre 2026</w:t></w:r></w:p>')
    _sostituisci(p, "Garofalo Healthcare S.p.A.", "Cliente S.r.l.")
    _sostituisci(p, "30 settembre 2026", "31 dicembre 2026")
    assert _testo(p) == "Cliente S.r.l. al 31 dicembre 2026"


def test_numerazione_didascalie():
    cap = ('<w:p><w:r><w:t xml:space="preserve">Tabella </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
           '<w:r><w:instrText> SEQ Tabella \\* ARABIC </w:instrText></w:r>'
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r>'
           '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t> - x</w:t></w:r></w:p>')
    body = parse_xml(f'<w:body {W}>{cap}{cap}<w:p><w:fldSimple w:instr=" SEQ Figura "><w:r><w:t>7</w:t></w:r>'
                     f'</w:fldSimple></w:p>{cap}</w:body>')
    _numera_didascalie(body)
    testi = [_testo(p) for p in body]
    assert testi == ["Tabella 1 - x", "Tabella 2 - x", "1", "Tabella 3 - x"]


CUR = os.environ.get("IAS19_GAROFALO_XLSX")
DOCX = os.environ.get("IAS19_RELAZIONE_DOCX")


@pytest.mark.skipif(not all(p and os.path.exists(p) for p in (CUR, DOCX)), reason="file Garofalo non disponibili")
def test_relazione_garofalo(tmp_path):
    import docx

    from ias19tfr.engine import valuta
    from ias19tfr.io_excel import dati_da_foglio_storico, ipotesi_da_excel
    from ias19tfr.relazione import DatiRelazione, scrivi_relazione
    from ias19tfr.scenari import griglia_sensitivity, scomposizione_csl

    ip = ipotesi_da_excel(CUR)
    dv = dati_da_foglio_storico(CUR)
    dv.dipendenti = [e for e in dv.dipendenti if e.azienda in (1, 4, 18)]  # Hesperia, Rugani, San Biagio
    res = valuta(dv.dipendenti, ip)
    dr = DatiRelazione(cliente="Garofalo Healthcare S.p.A.", data_valutazione=date(2026, 9, 30), dv=dv, res=res,
                       csl=scomposizione_csl(dv.dipendenti, ip, res), sens=griglia_sensitivity(dv.dipendenti, ip),
                       ip=ip)
    out = tmp_path / "relazione.docx"
    scrivi_relazione(DOCX, str(out), dr)

    d = docx.Document(str(out))
    testo = "\n".join(p.text for p in d.paragraphs)
    assert d.core_properties.comments == "30 settembre 2026"
    assert "i calcoli hanno riguardato n. 491 dipendenti" in testo
    assert re.search(r"San Biagio\s+dipendenti n\. 33", testo)
    assert "Nobili" not in testo.split("ALLEGATI")[0].split("suddivisi nelle seguenti")[1][:2000]
    celle = {c.text for t in d.tables for r in t.rows for c in r.cells}
    assert "628.462" in celle            # DBO Hesperia nella scheda (Tabella 7 GHC: 628.462,25)
    assert "332.875" in celle            # DBO San Biagio al centro della griglia di sensitivity
    cat = sorted({e.carriera for e in dv.dipendenti})
    assert any(c.startswith("MLx,p=") for c in celle)       # tavole dalle ipotesi della valutazione
    assert len([c for c in celle if c.startswith("MLx,p=")]) <= len(cat)
    # didascalie numerate in sequenza
    num = [int(m.group(1)) for p in d.paragraphs for m in [re.match(r"Tabella (\d+) ", p.text.strip())]
           if m and p.style.name.lower() in ("caption", "didascalia")]
    assert num == list(range(1, len(num) + 1))
