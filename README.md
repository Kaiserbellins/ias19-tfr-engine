# ias19-tfr-engine

Motore Python per la valutazione IAS 19 del Trattamento di Fine Rapporto (art. 2120 c.c.).

- **Metodo principale:** Financial proration (DBO = W̃ + (C̃ − C)), con anticipazioni come uscita finanziaria.
- **Metodo alternativo:** linear proration (PUCM classico, pro-rata anzianità pregressa / totale).
- **Output:** CCP, DW, PSL, CSL e la sua scomposizione (DC-R/A, DC-Inf, DC-Sal), DBO, service cost
  pro futuro, duration, erogazioni attese per anno, sensitivity.
- **Curva dei tassi:** parametri Svensson BCE (curva AAA o tutti i titoli) convertiti in forward annui.

Riconciliato al centesimo con il foglio di riferimento Garofalo al 30.09.2026: DBO di gruppo
13.352.807,06 €, DBO per azienda e griglia di sensitivity 3×3, scomposizione del CSL.

## Installazione

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
```

## Uso

```bash
# dati e ipotesi dal foglio di calcolo storico (fogli InputTFR, Tax, Tav)
ias19tfr valuta --dati "TFR Garofalo_2026_09.xlsx" --ipotesi "TFR Garofalo_2026_09.xlsx" \
    --formato storico --uscita risultati.xlsx

# dati dallo schema compilato dal cliente
ias19tfr valuta --dati "Schema dati TFR 30.09.2026.xlsx" --ipotesi ipotesi_2026_09.xlsx \
    --aziende-50 1,2,3 --uscita risultati.xlsx
```

Il file di ipotesi è una cartella Excel con i fogli `Tax` (inflazione, attualizzazione, incrementi
salariali per categoria, 50 anni) e `Tav` (RG48 e frequenze di dimissione e anticipazione per età),
nello stesso layout del foglio storico.

## Test

```bash
pytest                                   # test sintetici
IAS19_GAROFALO_XLSX=/percorso/foglio.xlsx pytest   # anche la riconciliazione con il foglio di riferimento
```

I dati dei clienti non vanno mai nel repository (`.gitignore` esclude `*.xlsx` e `data/`).

## Struttura

| Modulo | Contenuto |
| --- | --- |
| `model.py` | dipendente, tavole, ipotesi finanziarie, risultati aggregati |
| `engine.py` | proiezione individuale (flussi W1, W2, C1, C2) e valutazione |
| `scenari.py` | scomposizione CSL, sensitivity relative e assolute, curtailment |
| `riconciliazione.py` | movimentazione DBO, interest cost, scomposizione utili/perdite |
| `curva.py` | curva Svensson BCE e forward annui |
| `io_excel.py` | lettura schema cliente, foglio storico e file di ipotesi |
| `output.py` | Excel dei risultati |
