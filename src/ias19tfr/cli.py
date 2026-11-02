"""Riga di comando.

    ias19tfr esegui valutazione.yaml       valutazione completa da file di configurazione
    ias19tfr valuta --dati ... --ipotesi ... --uscita ...   valutazione rapida senza movimentazione

Esempio di valutazione.yaml (vedi esempi/valutazione_esempio.yaml):

    data_valutazione: 2026-09-30
    metodo: financial                  # oppure: lineare
    dati: {file: "Schema dati TFR 30.09.2026.xlsx", formato: cliente}
    ipotesi: "Ipotesi 2026_09.xlsx"    # fogli Tax e Tav
    aziende_almeno_50: [1, 2, 3, 5, 10, 17, 22, 23, 28, 29, 30]
    aziende_escluse: [27]
    trasferimenti: {32: 921968.00}
    precedente:
      data: 2025-12-31
      file: "TFR Garofalo_2025_12.xlsx"   # foglio storico: dati e ipotesi di apertura
    sensitivity_assoluta: true
    uscita: "Risultati IAS19 2026_09.xlsx"
    relazione:                         # facoltativo: relazione tecnica Word
      modello: "Relazione GHCSPA IAS19 2026_09.docx"   # relazione precedente usata come modello
      uscita: "Relazione IAS19 2026_09.docx"
      cliente: "Garofalo Healthcare S.p.A."            # default: parole chiave del modello
      categorie: {1: "Dirigenti", 2: "Quadri", 3: "Impiegati"}   # nomi dei codici carriera
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from .engine import valuta
from .io_excel import dati_da_foglio_storico, dati_da_schema_cliente, ipotesi_da_excel
from .output import scrivi_risultati
from .riconciliazione import movimentazione
from .scenari import griglia_sensitivity, scomposizione_csl, sensitivity_assoluta


def _data(v) -> date:
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v), "%Y-%m-%d").date()


def _lista_int(s: str):
    return {int(x) for x in s.split(",") if x.strip()} if s else None


def _carica_dati(spec: dict, base: Path, cfg: dict):
    f = str(base / spec["file"])
    if spec.get("formato", "cliente") == "storico":
        return dati_da_foglio_storico(f)
    return dati_da_schema_cliente(
        f,
        aziende_almeno_50=set(cfg.get("aziende_almeno_50") or []) or None,
        aziende_escluse=set(cfg.get("aziende_escluse") or []) or None,
        carriera_default=int(cfg.get("carriera_default", 3)),
    )


def esegui(cfg_path: str) -> int:
    import yaml

    base = Path(cfg_path).resolve().parent
    cfg = yaml.safe_load(Path(cfg_path).read_text(encoding="utf-8"))
    metodo = cfg.get("metodo", "financial")
    ip = ipotesi_da_excel(str(base / cfg["ipotesi"]))
    dv = _carica_dati(cfg["dati"], base, cfg)
    for w in dv.avvisi:
        print("AVVISO:", w, file=sys.stderr)

    res = valuta(dv.dipendenti, ip)
    csl = scomposizione_csl(dv.dipendenti, ip, res)
    sens = griglia_sensitivity(dv.dipendenti, ip, metodo)
    sa = sensitivity_assoluta(dv.dipendenti, ip, metodo) if cfg.get("sensitivity_assoluta") else None

    mov = None
    ip_prev = None
    tr = {}
    prec = cfg.get("precedente")
    if prec:
        fprev = str(base / prec["file"])
        ip_prev = ipotesi_da_excel(fprev)
        if "dbo_apertura" in prec:
            apertura = {int(k): float(v) for k, v in prec["dbo_apertura"].items()}
        else:
            dprev = dati_da_foglio_storico(fprev)
            apertura = {k: r.dbo_metodo(metodo) for k, r in valuta(dprev.dipendenti, ip_prev).items() if k}
        sc = {k: a.accantonamento for k, a in dv.aziende.items()}
        vf = {k: a.utilizzi for k, a in dv.aziende.items()}
        sc.update({int(k): float(v) for k, v in (cfg.get("service_cost") or {}).items()})
        vf.update({int(k): float(v) for k, v in (cfg.get("variazioni_fondo") or {}).items()})
        tr = {int(k): float(v) for k, v in (cfg.get("trasferimenti") or {}).items()}
        mov = movimentazione(dv.dipendenti, ip_prev, ip, apertura, sc, vf,
                             _data(prec["data"]), _data(cfg["data_valutazione"]), metodo, tr)

    uscita = str(base / cfg["uscita"])
    scrivi_risultati(uscita, dv, res, csl, sens, metodo, ip=ip, movimenti=mov, sens_assoluta=sa,
                     data_valutazione=_data(cfg["data_valutazione"]).strftime("%d/%m/%Y"))
    rel = cfg.get("relazione")
    if rel:
        import docx
        from .relazione import DatiRelazione, scrivi_relazione
        modello = str(base / rel["modello"])
        cliente = rel.get("cliente") or (docx.Document(modello).core_properties.keywords or "").strip()
        dr = DatiRelazione(
            cliente=cliente, data_valutazione=_data(cfg["data_valutazione"]), dv=dv, res=res, csl=csl,
            sens=sens, ip=ip, metodo=metodo, movimenti=mov, sens_assoluta=sa,
            data_precedente=_data(prec["data"]) if prec else None, ip_precedente=ip_prev, trasferimenti=tr,
            categorie={int(k): str(v) for k, v in (rel.get("categorie") or {}).items()},
            data_emissione=_data(rel["data_emissione"]) if rel.get("data_emissione") else None)
        scrivi_relazione(modello, str(base / rel["uscita"]), dr)
        print(f"Relazione scritta in {base / rel['uscita']}")

    t = res[0]
    print(f"Dipendenti {t.n}  CCP {t.ccp:,.2f}  DBO {t.dbo_metodo(metodo):,.2f}")
    if mov:
        m = mov[0]
        print(f"DBO apertura {m.dbo_apertura:,.2f}  SC {m.service_cost:,.2f}  NIE {m.nie:,.2f}  "
              f"VF {m.variazioni_fondo:,.2f}  Utile/(perdita) {m.utile:,.2f}")
    print(f"Risultati scritti in {uscita}")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="ias19tfr")
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("esegui", help="valutazione completa da file di configurazione YAML")
    e.add_argument("config")
    v = sub.add_parser("valuta", help="valutazione rapida e Excel dei risultati")
    v.add_argument("--dati", required=True)
    v.add_argument("--ipotesi", required=True, help="cartella con i fogli Tax e Tav")
    v.add_argument("--formato", choices=["cliente", "storico"], default="cliente")
    v.add_argument("--aziende-50", default="", help="codici aziende con almeno 50 dipendenti")
    v.add_argument("--metodo", choices=["financial", "lineare"], default="financial")
    v.add_argument("--uscita", required=True)
    a = ap.parse_args(argv)

    if a.cmd == "esegui":
        return esegui(a.config)

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
    scrivi_risultati(a.uscita, dv, res, csl, sens, a.metodo, ip=ip)
    t = res[0]
    print(f"Dipendenti {t.n}  CCP {t.ccp:,.2f}  PSL {t.psl:,.2f}  CSL {t.csl:,.2f}  DBO {t.dbo:,.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
