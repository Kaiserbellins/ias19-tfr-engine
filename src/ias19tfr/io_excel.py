"""Lettura degli input Excel.

Due formati di dati:
- schema cliente ("Schema dati TFR ... per attuario.xlsx"): fogli `codici` e `DATI TFR*`;
- foglio di calcolo storico (es. "TFR Garofalo_2026_09 ...xlsx"): foglio `InputTFR`.

Le basi tecniche si leggono da una cartella con i fogli `Tax` e `Tav` nel layout del foglio
storico, che diventa il "file di ipotesi" della valutazione.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set

import numpy as np
import openpyxl

from .model import ALIQUOTA_TFR, N_ANNI, Dipendente, Ipotesi, IpotesiFinanziarie, TavolaPermanenza

ETA_MAX = 111
N_CATEGORIE = 10


def _apri(path: str):
    if str(path).lower().endswith(".xls"):
        raise ValueError(f"{path}: formato .xls non supportato, salvarlo come .xlsx")
    return openpyxl.load_workbook(path, data_only=True, read_only=True)


def _num(v, default=0.0) -> float:
    if v is None or v == "":
        return default
    if isinstance(v, str):
        v = v.strip().replace(",", ".")
        if v == "":
            return default
    return float(v)


# ---------------------------------------------------------------- basi tecniche

def ipotesi_da_excel(path: str, aliquota_tfr: float = ALIQUOTA_TFR) -> Ipotesi:
    """Legge `Tax` e `Tav` nel layout del foglio storico.

    Tav: colonne X/Y = l_x RG48 maschi/femmine (righe 2..113 = eta 0..111);
         dimissioni per eta in righe 116..227, anticipazioni in righe 230..341;
         colonne B..K = maschi p=1..10, L..U = femmine p=1..10.
    Tax: righe 2..51 = anni 1..50; B inflazione, C rivalutazione (solo controllo),
         D attualizzazione, F..O incremento salariale reale p=1..10.
    """
    wb = _apri(path)
    tav = wb["Tav"]
    rows = list(tav.iter_rows(min_row=1, max_row=341, max_col=27, values_only=True))

    def cell(r, c):  # r, c 1-based
        return rows[r - 1][c - 1]

    lx_base = {
        1: np.array([_num(cell(2 + x, 24)) for x in range(ETA_MAX + 1)]),
        2: np.array([_num(cell(2 + x, 25)) for x in range(ETA_MAX + 1)]),
    }
    dim, ant = {}, {}
    for sesso in (1, 2):
        for p in range(1, N_CATEGORIE + 1):
            col = 1 + p if sesso == 1 else 11 + p
            dim[(sesso, p)] = np.array([_num(cell(116 + x, col)) for x in range(ETA_MAX + 1)])
            ant[(sesso, p)] = np.array([_num(cell(230 + x, col)) for x in range(ETA_MAX + 1)])
    tavola = TavolaPermanenza.costruisci(lx_base, dim, ant)

    tax = wb["Tax"]
    trows = list(tax.iter_rows(min_row=2, max_row=1 + N_ANNI, max_col=15, values_only=True))
    infl = np.array([_num(r[1]) for r in trows])
    riv_foglio = np.array([_num(r[2]) for r in trows])
    att = np.array([_num(r[3]) for r in trows])
    incr = {p: np.array([_num(r[4 + p]) for r in trows]) for p in range(1, N_CATEGORIE + 1)}
    fin = IpotesiFinanziarie(inflazione=infl, attualizzazione=att, incremento_reale=incr)
    scarto = np.abs(fin.tasso_rivalutazione() - riv_foglio).max()
    if scarto > 1e-5:
        raise ValueError(f"Tax: la rivalutazione del foglio differisce dalla regola di legge di {scarto:.6f}")
    wb.close()
    return Ipotesi(tavola=tavola, finanziarie=fin, aliquota_tfr=aliquota_tfr)


# ---------------------------------------------------------------- dati dipendenti

@dataclass
class Azienda:
    numero: int
    nome: str
    dipendenti: int = 0
    tfr: float = 0.0
    accantonamento: float = 0.0   # TFR maturato nel periodo (service cost contabile)
    utilizzi: float = 0.0         # variazioni/utilizzi del fondo nel periodo (VF)


@dataclass
class DatiValutazione:
    dipendenti: List[Dipendente]
    aziende: Dict[int, Azienda] = field(default_factory=dict)
    avvisi: List[str] = field(default_factory=list)


def _normalizza(codice, azienda, eta, sesso, anz, futuro, carriera, salario, tfr, fip,
                arrotonda: bool = True) -> Dipendente:
    """Stesse regole del foglio `DatiTFR`."""
    if arrotonda:
        return Dipendente(
            codice=int(_num(codice)), azienda=int(_num(azienda)),
            eta=int(math.floor(_num(eta))), sesso=int(_num(sesso)),
            anzianita=max(int(math.floor(_num(anz))), 1),
            futuro=max(int(math.floor(_num(futuro))), 1),
            carriera=int(_num(carriera)),
            salario=float(math.floor(_num(salario))),
            tfr=float(max(math.floor(_num(tfr) + 0.5), 0)),
            fip=_num(fip),
        )
    return Dipendente(int(_num(codice)), int(_num(azienda)), int(_num(eta)), int(_num(sesso)),
                      max(int(_num(anz)), 1), max(int(_num(futuro)), 1), int(_num(carriera)),
                      _num(salario), max(_num(tfr), 0.0), _num(fip))


def dati_da_foglio_storico(path: str, arrotonda: bool = True) -> DatiValutazione:
    """Legge `InputTFR` (FIP espressa in percentuale, es. 6.9074) e i nomi aziende da `Pilota`."""
    wb = _apri(path)
    ws = wb["InputTFR"]
    dip = []
    for r in ws.iter_rows(min_row=3, max_col=13, values_only=True):
        if not isinstance(r[2], (int, float)) and not (isinstance(r[2], str) and r[2].strip().isdigit()):
            continue
        _, _, cod, az, _cc, eta, sesso, anz, fut, car, ral, tfr, fip = r
        dip.append(_normalizza(cod, az, eta, sesso, anz, fut, car, ral, tfr, _num(fip) / 100, arrotonda))
    aziende = {}
    pil = wb["Pilota"]
    for r in pil.iter_rows(min_row=12, max_row=60, max_col=3, values_only=True):
        if isinstance(r[0], (int, float)) and r[1]:
            aziende[int(r[0])] = Azienda(int(r[0]), str(r[1]), int(_num(r[2])))
    wb.close()
    return DatiValutazione(dip, aziende)


def dati_da_schema_cliente(path: str, aziende_almeno_50: Optional[Set[int]] = None,
                           carriera_default: int = 3, arrotonda: bool = True,
                           aziende_escluse: Optional[Set[int]] = None,
                           sesso_default: int = 1) -> DatiValutazione:
    """Legge lo schema dati compilato dal cliente e applica le regole di passaggio all'InputTFR.

    Regole (confermate da Francesco il 02/11/2026, verificate sul caso Garofalo 30/09/2026):
    - esclusi i dipendenti liquidati: righe senza dati anagrafici (eta vuota);
    - escluse le aziende in `aziende_escluse` (es. cedute: Eutonia);
    - FIP vuota = 0; FIP scritta come percentuale (6,9074) o arrotondata (0,0691) = aliquota piena;
      valori intermedi tra 0 e l'aliquota piena sono ammessi;
    - `aziende_almeno_50`: aziende con almeno 50 dipendenti (L. 296/2006), FIP = aliquota piena
      per tutti i dipendenti;
    - carriera vuota o 0 = `carriera_default` (3);
    - futuro servizio 0, anzianita 0 e TFR negativo sono portati a 1, 1 e 0 (`_normalizza`).
    """
    wb = _apri(path)
    avvisi: List[str] = []
    aziende: Dict[int, Azienda] = {}
    cod = wb["codici"]
    for r in cod.iter_rows(min_row=3, max_col=9, values_only=True):
        if isinstance(r[0], (int, float)) and r[1]:
            aziende[int(r[0])] = Azienda(int(r[0]), str(r[1]).strip(), int(_num(r[2])),
                                         _num(r[6]), _num(r[7]), _num(r[8]))
    nome_dati = next(n for n in wb.sheetnames if n.upper().startswith("DATI TFR"))
    ws = wb[nome_dati]
    dip = []
    for i, r in enumerate(ws.iter_rows(min_row=3, max_col=13, values_only=True), start=3):
        if not isinstance(r[1], (int, float)):
            continue
        _, codice, az, eta, sesso, anz, fut, tfr, fip, car, sal = r[:11]
        if eta in (None, ""):
            continue  # liquidato
        if aziende_escluse and int(_num(az)) in aziende_escluse:
            continue
        if car in (None, "", 0):
            car = carriera_default
        if sesso in (None, ""):
            avvisi.append(f"riga {i}: sesso mancante, assunto {sesso_default}")
            sesso = sesso_default
        fip = _num(fip)
        if fip > 1:  # percentuale scritta come 6,9074
            fip /= 100
        if abs(fip - ALIQUOTA_TFR) < 1e-4:  # es. 0,0691 arrotondato
            fip = ALIQUOTA_TFR
        if aziende_almeno_50 is not None and int(_num(az)) in aziende_almeno_50:
            fip = ALIQUOTA_TFR
        try:
            dip.append(_normalizza(codice, az, eta, sesso, anz, fut, car, sal, tfr, fip, arrotonda))
        except (TypeError, ValueError) as exc:
            avvisi.append(f"riga {i}: dato non valido ({exc})")
    wb.close()
    dv = DatiValutazione(dip, aziende, avvisi)
    dv.avvisi.extend(valida(dv))
    return dv


def valida(dv: DatiValutazione) -> List[str]:
    """Controlli di coerenza: restituisce avvisi, non blocca la valutazione."""
    av = []
    conta: Dict[int, int] = {}
    tfr: Dict[int, float] = {}
    for e in dv.dipendenti:
        conta[e.azienda] = conta.get(e.azienda, 0) + 1
        tfr[e.azienda] = tfr.get(e.azienda, 0.0) + e.tfr
        if not 15 <= e.eta <= 75:
            av.append(f"dip. {e.codice}: eta {e.eta} fuori range")
        if e.sesso not in (1, 2):
            av.append(f"dip. {e.codice}: sesso {e.sesso} non valido")
        if not 0 <= e.fip <= ALIQUOTA_TFR + 1e-9:
            av.append(f"dip. {e.codice}: aliquota FIP {e.fip} fuori range")
    for n, a in dv.aziende.items():
        # tolleranza: arrotondamento all'euro di ogni posizione individuale
        if a.tfr and abs(tfr.get(n, 0.0) - a.tfr) > 0.5 * conta.get(n, 1) + 1:
            av.append(f"azienda {n} {a.nome}: TFR dati {tfr.get(n, 0.0):,.2f} vs contabile {a.tfr:,.2f}")
    return av
