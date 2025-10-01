# Analisi Produzione e Consumo Energetico

Questo progetto fornisce una serie di strumenti per consolidare, analizzare e visualizzare i dati di produzione e consumo energetico delle macchine industriali.

Il sistema è composto da tre script principali che coprono l'intero flusso di lavoro: dalla preparazione dei dati alla visualizzazione interattiva e alla reportistica statica.

## Struttura dei File

- **`Dati consumi e costi energetici.xlsx`**: Il file Excel principale. Contiene i fogli di lavoro individuali per ogni macchina (es. `F01`, `ASS1`, ecc.) dove vengono inseriti i dati grezzi.
- **`crea_consolidato.py`**: Il primo script da eseguire. Legge tutti i fogli macchina dal file Excel principale e crea (o aggiorna) un foglio riepilogativo chiamato `Consolidato`.
- **`energy_dashboard.py`**: Il secondo script. Avvia una dashboard web interattiva (basata su Streamlit) che legge i dati dal foglio `Consolidato` per un'analisi visuale.
- **`unisci_dati.py`**: Uno script indipendente per generare report statici in formato Excel, unendo dati da file separati.
- **`prod_quantita.xlsx` / `prod_consumo_macchine.xlsx`**: File di input richiesti solo per lo script `unisci_dati.py`.
- **`report.xlsx` / `report_anomalie.xlsx`**: File di output generati da `unisci_dati.py`.

## Flusso di Lavoro Principale

Segui questi passaggi per analizzare i dati più recenti.

### Passo 1: Aggiornare i Dati Grezzi

Assicurati che tutti i dati di produzione e consumo più recenti siano stati inseriti nei rispettivi fogli di lavoro (es. `F01`, `F02`, ecc.) all'interno del file `Dati consumi e costi energetici.xlsx`.

### Passo 2: Eseguire il Consolidamento dei Dati

Apri un terminale nella directory del progetto ed esegui lo script di consolidamento:

```bash
python crea_consolidato.py
```

Questo script leggerà tutti i fogli delle macchine e aggiornerà il foglio `Consolidato` con i dati più recenti, rendendoli pronti per l'analisi.

### Passo 3: Avviare la Dashboard Interattiva

Dopo aver consolidato i dati, avvia la dashboard per l'analisi visuale con il seguente comando:

```bash
streamlit run energy_dashboard.py
```

Si aprirà una pagina web nel tuo browser che ti permetterà di filtrare i dati per macchina, anno e mese e di visualizzare grafici interattivi su consumi e costi.

## Flusso di Lavoro Alternativo (Report Statici)

Se hai bisogno di generare report Excel statici (come descritto nella versione precedente di questo README), puoi usare lo script `unisci_dati.py`.

1.  Assicurati che i file `prod_quantita.xlsx` e `prod_consumo_macchine.xlsx` siano aggiornati.
2.  Esegui lo script:
    ```bash
    python unisci_dati.py
    ```
3.  Verranno generati i file `report.xlsx` e `report_anomalie.xlsx`.

## Requisiti

Assicurati di avere Python 3.x installato. Le librerie necessarie possono essere installate con un unico comando:

```bash
pip install pandas openpyxl streamlit plotly
```

---
*Documentazione aggiornata il 01/10/2025*