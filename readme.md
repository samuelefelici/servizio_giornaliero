# Servizio Giornaliero – ExtraUrbano (Python/Streamlit)

Applicazione web Streamlit per l'elaborazione automatica dei turni di servizio giornaliero.

Porting della macro Excel in Python con le seguenti funzionalità:
- **Parsing robusto**: ricerca automatica dell'header (colonna "Cognome e Nome"), supporto per formati Excel multipli (.xls, .xlsx, HTML, CSV)
- **Filtri configurabili**: esclusione di matricole e turni specifici tramite file CSV di configurazione
- **Normalizzazione dati**: rinomina automatica delle residenze secondo mapping predefinito
- **Gestione assenze**: sostituzione automatica delle sigle di assenza (PC, AA, AF, etc.) con "Assente"
- **Inserimento trasferte**: interfaccia per aggiungere manualmente trasferte fuori residenza
- **Evidenziazione trasferte**: evidenzia automaticamente i turni fuori dalla residenza di appartenenza
- **Offset orari**: applica automaticamente +2 ore agli orari (escluse le righe inserite manualmente)
- **Ordinamenti configurabili**: ordina per cognome/nome o per orario di inizio
- **Visualizzazione per deposito**: raggruppa e visualizza i dati per ogni deposito
- **Export multipli**: genera file Excel e PDF con formattazione professionale

## Requisiti

Python 3.11+

## Installazione

```bash
pip install -r requirements.txt
```

## Utilizzo

```bash
streamlit run app.py
```

L'applicazione si aprirà nel browser predefinito all'indirizzo `http://localhost:8501`.

## Struttura del progetto

- `app.py`: interfaccia utente Streamlit
- `src/process.py`: elaborazione e trasformazione dei dati Excel
- `src/pdf_export.py`: generazione del PDF con formattazione avanzata
- `src/constants.py`: costanti e configurazioni (sigle assenze, mapping residenze, etc.)
- `src/utils.py`: funzioni di utilità
- `config/`: file CSV per filtri configurabili
  - `matricole_da_omettere.csv`: matricole da escludere
  - `turni_attivita_da_omettere.csv`: turni da escludere
- `assest/`: risorse (logo per PDF)

## Funzionalità principali

### Upload e elaborazione
1. Carica un file Excel (.xls o .xlsx) con i turni di servizio
2. Il sistema rileva automaticamente l'header e la data
3. Applica filtri, normalizzazioni e trasformazioni configurate

### Inserimento trasferte
- Aggiungi manualmente trasferte fuori residenza
- Compilazione guidata con validazione dei campi obbligatori
- Le trasferte inserite mantengono gli orari originali (senza offset)

### Anteprima
- Visualizzazione raggruppata per deposito
- Evidenziazione automatica delle trasferte in grassetto
- Righe aggiunte manualmente mostrate in blu
- Compattazione delle righe duplicate (stesso nominativo/matricola)

### Export
- **Excel**: file .xlsx con tutti i dati elaborati
- **PDF**: documento formattato con logo, intestazioni e stili professionali
