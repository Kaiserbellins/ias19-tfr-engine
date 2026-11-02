# ias19-tfr-engine

Motore Python per la valutazione IAS 19 del Trattamento di Fine Rapporto (art. 2120 c.c.).

- **Metodo principale:** Financial proration (DBO = W̃ + (C̃ − C)), con anticipazioni come uscita finanziaria.
- **Metodo alternativo:** linear proration (PUCM classico, pro-rata anzianità pregressa / totale).
- **Output:** CCP, DW, PSL, CSL e la sua scomposizione (DC-R/A, DC-Inf, DC-Sal), DBO, service cost
  pro futuro, duration, erogazioni attese per anno, sensitivity.
- **Curva dei tassi:** parametri Svensson BCE della curva AAA convertiti in forward annui.
- **Movimentazione:** DBO di apertura, service cost contabile, interest cost (forward anno 1 della
  valutazione precedente × giorni/365), variazioni del fondo, utili/perdite scomposti in esperienza,
  ipotesi demografiche e ipotesi finanziarie.

Riconciliato al centesimo con il foglio di riferimento Garofalo al 30.09.2026: DBO di gruppo
13.352.807,06 €, DBO per azienda e griglia di sensitivity 3×3, scomposizione del CSL.

## Installazione

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
```

## Uso

La valutazione completa (con movimentazione rispetto alla valutazione precedente) si descrive in un
file YAML: vedi [`esempi/valutazione_esempio.yaml`](esempi/valutazione_esempio.yaml).

```bash
ias19tfr esegui valutazione.yaml
```

L'Excel dei risultati contiene i fogli Riepilogo (schede per azienda), Movimentazione, Sensitivity
(griglia ±10% e variazioni assolute su tassi, inflazione e turnover), Andamento (dati del grafico
prospettico), Disclosure (duration, service cost pro futuro, erogazioni attese), Statistiche (per età
e anzianità) e Dipendenti (risultati individuali).

Con la sezione `relazione` del YAML viene prodotta anche la relazione tecnica Word, usando come
modello una relazione già emessa (stili, copertina, intestazioni, testi metodologici e formule
restano quelli del modello). Vengono rigenerati dai risultati:

- cliente, data di valutazione, data di emissione, numero di dipendenti ed elenco delle società;
- Tabelle 1-6 (RG48, dimissioni e anticipazioni, tavole di permanenza, tassi annui) e incrementi
  salariali per categoria, dalle ipotesi della valutazione;
- caratteristiche del collettivo (per società, per età e anzianità);
- Tabella 7 della movimentazione e scomposizione degli utili/perdite attuariali per società;
- informativa IAS 19: riconciliazione della DBO, ipotesi principali, sensitivity con variazioni
  assolute, duration, service cost pro futuro, erogazioni attese;
- allegati: scheda di valutazione e grafico prospettico per ogni società, griglie di sensitivity.

Le tabelle che nel modello erano immagini diventano tabelle Word. Sommario e indici si aggiornano
all'apertura del documento in Word (confermare la richiesta di aggiornamento dei campi).

Prospetto delle movimentazioni trimestrali per la transizione all'IFRS 18 (service cost in categoria
operativa, interest cost in categoria di finanziamento, utili/perdite attuariali in OCI), a partire
dalle configurazioni delle valutazioni progressive dello stesso esercizio:

```bash
ias19tfr ifrs18 v2026_03.yaml v2026_06.yaml v2026_09.yaml --uscita "Prospetto IFRS18 2026.xlsx"
```

Valutazione rapida senza movimentazione:

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

Per la movimentazione servono anche `IAS19_GAROFALO_PREV_XLSX` (valutazione precedente, da
convertire in .xlsx se è un .xls) e `IAS19_SCHEMA_XLSX` (schema dati cliente); per la relazione
`IAS19_RELAZIONE_DOCX` (relazione usata come modello). Il caso Finnat (tre trimestri 2026) usa
`IAS19_FINNAT_DIR`, cartella con i fogli storici convertiti in .xlsx.

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
| `relazione.py` | relazione tecnica Word dal modello |
| `ifrs18.py` | movimentazioni trimestrali classificate per l'IFRS 18 |
