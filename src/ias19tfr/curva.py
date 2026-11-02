"""Curva dei tassi di attualizzazione dalla BCE (modello Svensson).

La BCE pubblica ogni giorno lavorativo i parametri Svensson della curva dei rendimenti dell'area
euro (https://www.ecb.europa.eu/stats/financial_markets_and_interest_rates/euro_area_yield_curves/).
Serie: YC.B.U2.EUR.4F.G_N_A.SV_C_YM.<parametro> (curva dei titoli di Stato AAA, scelta dello
studio) con <parametro> in BETA0..BETA3, TAU1, TAU2.

I tassi Svensson sono in percentuale, composti in continuo. Il motore li converte in tassi spot
annui effettivi e poi in forward annui, per gli anni 1..50; oltre la scadenza massima
(default 30 anni) il forward resta costante.
"""
from __future__ import annotations

import csv
import io
import math
from dataclasses import dataclass
from typing import Dict, Optional

import numpy as np

CURVE = {"AAA": "G_N_A"}
PARAMETRI = ("BETA0", "BETA1", "BETA2", "BETA3", "TAU1", "TAU2")
URL = "https://data-api.ecb.europa.eu/service/data/YC/B.U2.EUR.4F.{curva}.SV_C_YM.{par}"


@dataclass(frozen=True)
class Svensson:
    beta0: float
    beta1: float
    beta2: float
    beta3: float
    tau1: float
    tau2: float
    data: str = ""
    curva: str = "AAA"

    def spot_continuo(self, m: float) -> float:
        """Tasso spot (decimale, composto in continuo) per scadenza m anni."""
        x1, x2 = m / self.tau1, m / self.tau2
        f1 = (1 - math.exp(-x1)) / x1
        f2 = f1 - math.exp(-x1)
        f3 = (1 - math.exp(-x2)) / x2 - math.exp(-x2)
        return (self.beta0 + self.beta1 * f1 + self.beta2 * f2 + self.beta3 * f3) / 100.0

    def spot_annuo(self, m: float) -> float:
        return math.exp(self.spot_continuo(m)) - 1.0

    def forward_annui(self, anni: int = 50, scadenza_max: int = 30, spread: float = 0.0) -> np.ndarray:
        """Forward annui i_t (t = 1..anni), con eventuale spread aggiunto agli spot annui."""
        fwd = np.zeros(anni)
        prev = 1.0
        for t in range(1, anni + 1):
            if t <= scadenza_max:
                s = self.spot_annuo(t) + spread
                cur = (1 + s) ** t
                fwd[t - 1] = cur / prev - 1.0
                prev = cur
            else:
                fwd[t - 1] = fwd[scadenza_max - 1]
        return fwd


def scarica_parametri(mese: str, curva: str = "AAA", timeout: float = 30.0) -> Svensson:
    """Parametri Svensson dell'ultimo giorno lavorativo del mese (formato 'AAAA-MM')."""
    import requests  # import locale: il resto del motore non richiede rete

    anno, m = map(int, mese.split("-"))
    inizio = f"{anno:04d}-{m:02d}-01"
    fine = f"{anno:04d}-{m:02d}-31"
    valori: Dict[str, float] = {}
    data_rif: Optional[str] = None
    for par in PARAMETRI:
        url = URL.format(curva=CURVE[curva.upper()], par=par)
        r = requests.get(url, params={"startPeriod": inizio, "endPeriod": fine, "format": "csvdata"},
                         timeout=timeout)
        r.raise_for_status()
        righe = list(csv.DictReader(io.StringIO(r.text)))
        if not righe:
            raise ValueError(f"Nessun dato BCE per {par} in {mese}")
        ultima = max(righe, key=lambda x: x["TIME_PERIOD"])
        valori[par] = float(ultima["OBS_VALUE"])
        data_rif = ultima["TIME_PERIOD"]
    return Svensson(valori["BETA0"], valori["BETA1"], valori["BETA2"], valori["BETA3"],
                    valori["TAU1"], valori["TAU2"], data=data_rif or "", curva=curva.upper())
