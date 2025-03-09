# Concatenator

Concatenator è un semplice ma efficace strumento in Python che permette di selezionare immagini dal tuo computer, raggrupparle automaticamente per operatore (TIM, Vodafone, WindTre, Iliad) in base ai nomi dei file, e generare automaticamente un documento Word completo e organizzato.

Questo strumento è particolarmente utile per generare report rapidi, cataloghi di immagini, o raccolte fotografiche suddivise per categorie.

## Funzionalità

- **Selezione multipla di immagini** direttamente dall'interfaccia grafica.
- **Filtraggio automatico delle immagini** in base al nome dei file per gli operatori telefonici più comuni: TIM, Vodafone (VF), WindTre (W3), Iliad.
- **Generazione automatica di un documento Word** con orientamento orizzontale per facilitare la visualizzazione delle immagini.
- **Organizzazione delle immagini** in blocchi personalizzabili con titoli specifici.

## Requisiti

- Python 3.6 o superiore
- Librerie Python:
  - `python-docx`
  - `Pillow`
  - `tkinter`

## Installazione

Installa le dipendenze richieste usando pip:

```bash
pip install python-docx Pillow
```

Tkinter è solitamente incluso nella distribuzione standard di Python. Se necessario, segui la guida ufficiale per la tua piattaforma.

## Come usare Concatenator

1. Avvia lo script con:

```bash
python concatenator.py
```

2. Inserisci il titolo generale per il documento quando richiesto.

3. Segui le istruzioni nella console per aggiungere i blocchi di immagini:
   - Dai un nome a ogni blocco.
   - Seleziona le immagini attraverso la finestra grafica che si aprirà automaticamente.

4. Ripeti il processo per ogni gruppo di immagini che desideri inserire.

5. Una volta conclusa la selezione, lo script genererà automaticamente un documento Word con tutte le immagini organizzate per blocchi e per operatore.

6. Il file Word sarà salvato automaticamente nella cartella corrente, con un nome basato sul titolo del documento inserito.

## Struttura del documento generato

Il documento Word finale avrà una struttura organizzata in titoli (blocchi), e ciascun blocco sarà ulteriormente diviso per operatore, in base alle sigle riconosciute nei nomi delle immagini (TIM, W3, VF, Iliad).

## Esempio di nome file riconosciuto
- `screenshot_TIM_01.png`
- `foto_VF_evento.jpeg`
- `test_W3_image.jpg`
- `promo_iliad_offerta.png`

Assicurati che le immagini contengano chiaramente una delle sigle degli operatori (TIM, VF, W3, Iliad) nel nome del file per una corretta categorizzazione automatica.

## Supporto e contributi

Se hai problemi o suggerimenti, apri una issue nella sezione dedicata di GitHub.

## Licenza

Questo progetto è rilasciato sotto la licenza MIT. Consulta il file [LICENSE](LICENSE) per ulteriori informazioni.
