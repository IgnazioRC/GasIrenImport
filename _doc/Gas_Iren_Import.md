# Gas Iren Import

Script Python per l'importazione automatica delle bollette gas Iren (formato PDF) in un file Excel con analisi dettagliate dei consumi e dei parametri contrattuali.

**Autore:** Ignazio Rusconi Clerici — Novembre 2025, aggiornato Marzo 2026

---

## Requisiti

- Python 3.9+
- Librerie (installare nel venv `Python`):

```bash
pip install pdfplumber pandas xlsxwriter openpyxl
```

---

## Struttura cartelle attesa

```
BASE_DIR/
    Confronto consumi gas (Iren).xlsx   ← creato/aggiornato dallo script
    Gignese/
        *.pdf                  ← bollette PDF Gignese
    PuntaAla/
        *.pdf                  ← bollette PDF Punta Ala
```

---

## Avvio

```bash
python3 gas_iren_import.py
```

Si apre una finestra GUI. Non è previsto utilizzo da riga di comando.

---

## Configurazione

La configurazione viene salvata automaticamente in:

```
~/_Config/GasIrenImport/config.json
```

Contiene l'ultimo percorso usato e viene ricaricata ad ogni avvio.

---

## Interfaccia utente

| Elemento | Descrizione |
|---|---|
| **Cartella base** | Percorso della cartella contenente `Gignese/` e `PuntaAla/`. Mostrato in forma abbreviata da `Dropbox/` o `~/`. Sola lettura. |
| **Sfoglia…** | Apre il selettore di cartelle posizionandosi sulla cartella già in uso. |
| **▶ Esegui** | Avvia l'elaborazione in un thread separato. Si disabilita durante l'esecuzione. |
| **Log** | Pannello scrollabile con il dettaglio di ogni operazione. |
| **Esci** | Chiude l'applicazione. |

---

## Modalità di funzionamento

### Modalità incrementale (Excel esistente)
Se `Confronto consumi gas (Iren).xlsx` è già presente nella cartella base:

1. Viene creato un backup: `Confronto consumi gas (Iren).bak.YYYYMMDD.xlsx`
2. I dati esistenti vengono ricaricati dall'Excel
3. Vengono scansionati i PDF nelle sottocartelle
4. Vengono aggiunte **solo le bollette nuove** (chiave: numero fattura; fallback: località + periodo)
5. L'Excel viene riscritto con tutti i dati uniti

### Modalità completa (primo avvio)
Se l'Excel non esiste, tutti i PDF trovati vengono elaborati e l'Excel viene creato da zero.

---

## Dati estratti da ogni bolletta PDF

| Campo | Descrizione |
|---|---|
| Periodo | Etichetta mese/anno (es. `feb-26`) |
| Smc | Consumi fatturati in Standard metri cubi (0.0 se mese senza consumi) |
| PSV | Prezzo Sul Virtuale €/smc (indice contrattuale mensile) |
| Adeguamento PCS | Correzione per potere calorifico superiore €/smc |
| Sconto % PSV+Spread | Percentuale di sconto applicata (normalmente 15%) |
| Sconto unitario | Sconto in €/smc ricavato dal quadro di dettaglio |
| Spread dichiarato | Spread contrattuale €/smc (normalmente 0) |
| Spread calcolato | Spread ricavato dai dati della bolletta |
| Spread OK | `OK` se lo spread calcolato è compatibile con 0 (±0,0005 €/smc), `NO` altrimenti, `ND` se non calcolabile |
| Quota fissa vendita | Costo fisso mensile parte vendita (atteso 6,00 €/mese) |
| Quota consumi vendita | Costo variabile parte vendita |
| Totale materia | Totale spesa materia prima gas |
| Spesa rete+oneri | Trasporto, gestione contatore e oneri di sistema |
| Accise+IVA | Imposte totali |
| Totale bolletta | Importo totale fatturato |

---

## Struttura file Excel generato

| Foglio | Contenuto |
|---|---|
| **Gignese** | Tutte le bollette di Gignese, ordinate per data |
| **PuntaAla** | Tutte le bollette di Punta Ala, ordinate per data |
| **Confronto Mensile** | Smc e totale bolletta per mese, entrambe le località affiancate, ordine cronologico |
| **Confronto Annuale** | Consumo rolling 12 mesi per entrambe le località |
| **Log** | Storico completo di tutte le elaborazioni con esito per ogni PDF |

---

## Formula prezzo contrattuale

```
Prezzo Vendita = (PSV + Spread + Adeguamento PCS) × (1 − 15%)
```

Lo script verifica che lo sconto del 15% venga applicato correttamente da Iren calcolando lo spread implicito dalla bolletta e confrontandolo con il valore dichiarato.

---

## Note

- Le bollette devono essere in formato PDF originale Iren (non scansioni)
- I file PDF che non contengono un periodo di riferimento riconoscibile vengono ignorati (es. contratti)
- Il backup giornaliero sovrascrive il backup dello stesso giorno se lo script viene eseguito più volte
