"""Riga di comando.

Esempi:
    ias19tfr valuta --dati "TFR Garofalo_2026_09.xlsx" --ipotesi "TFR Garofalo_2026_09.xlsx" \
        --formato storico --uscita risultati.xlsx
    ias19tfr valuta --dati "Schema dati TFR 30.09.2026.xlsx" --ipotesi ipotesi_2026_09.xlsx \
        --aziende-50 1,2,3,10 --uscita risultati.xlsx
"""
from __future__ import annotations

import argparse
import sys

from .engine import valuta
from .io_excel import dati_da_foglio_storico, dati_da_schema_cliente, ipotesi_da_excel
from .output import scrivi_risultati
from .scenari import griglia_sensitivity, scomposizione_csl


def _lista_int(s: str):
    return {int(x) for x in s.split(",") if x.strip()} if s else None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="ias19tfr")
    sub = ap.add_subparsers(dest="cmd", required=True)
    v = sub.add_parser("valuta", help="valuta il TFR e scrive l'Excel dei risultati")
    v.add_argument("--dati", required=True)
    v.add_argument("--ipotesi", required=True, help="cartella con i fogli Tax e Tav")
    v.add_argument("--formato", choices=["cliente", "storico"], default="cliente")
    v.add_argument("--aziende-50", default="", help="codici aziende con almeno 50 dipendenti")
    v.add_argument("--metodo", choices=["financial", "lineare"], default="financial")
    v.add_argument("--uscita", required=True)
    a = ap.parse_args(argv)

    ip = ipotesi_da_excel(a.ipotesi)
    if a.formato == "storico":
        dv = dati_da_foglio_storico(a.dati)
    else:
        dv = dati_da_schema_cliente(a.dati, aziende_almeno_50=_lista_int(a.aziende_50))
    for w in dv.avvisi:
        print("AVVISO:", w, file=sys.stderr)
    res = valuta(dv.dipendenti, ip)
    csl = scomposizione_csl(dv.dipendenti, ip, res)
    sens = griglia_sensitivity(dv.dipendenti, ip, a.metodo)
    scrivi_risultati(a.uscita, dv, res, csl, sens, a.metodo)
    t = res[0]
    print(f"Dipendenti {t.n}  CCP {t.ccp:,.2f}  PSL {t.psl:,.2f}  CSL {t.csl:,.2f}  DBO {t.dbo:,.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
