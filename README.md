# Concatenator.py

## Descrizione
`concatenator.py` è un tool sviluppato per automatizzare la creazione di documenti Word contenenti immagini organizzate per operatore. Il programma consente di selezionare immagini dal file system, raggrupparle in blocchi, e generare documenti Word separati per ciascun operatore riconosciuto nei nomi dei file.

## Funzionalità
- Selezione interattiva delle immagini tramite interfaccia grafica.
- Riconoscimento automatico dell'operatore dai nomi dei file.
- Creazione di documenti Word con orientamento orizzontale.
- Inserimento di immagini ridimensionate mantenendo le proporzioni.
- Raggruppamento delle immagini in blocchi definiti dall'utente.
- Supporto a diversi formati immagine (`.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`).

## Requisiti
- Python 3.x
- Librerie richieste:
  - `python-docx`
  - `Pillow`
  - `tkinter`
  - `tqdm`

Puoi installare le dipendenze eseguendo:
```sh
pip install python-docx pillow tqdm
```

## Utilizzo
1. Avvia il programma eseguendo:
   ```sh
   python concatenator.py
   ```
2. Inserisci il titolo del documento quando richiesto.
3. Aggiungi uno o più blocchi di immagini selezionando i file desiderati.
4. Il programma genererà automaticamente documenti Word separati per ciascun operatore rilevato nei nomi dei file.

## Struttura del Codice
- **`select_images()`**: Apre una finestra di dialogo per la selezione delle immagini.
- **`filter_images_by_operator()`**: Classifica le immagini in base agli operatori (Iliad, TIM, VF, W3).
- **`create_or_update_document()`**: Genera o aggiorna un documento Word con le immagini selezionate.
- **`__main__`**: Gestisce il flusso principale dell'applicazione.

## Esempio di Output
Dopo aver eseguito il programma, verranno generati file Word con nomi del tipo:
```
Nome_Documento_Iliad.docx
Nome_Documento_TIM.docx
Nome_Documento_VF.docx
Nome_Documento_W3.docx
```
Ciascun file conterrà le immagini raggruppate per operatore.

## Autori
Sviluppato da **Alessandro Frullo** in collaborazione con **Selektra Italia Srl**.

## Licenza
Questo progetto è rilasciato sotto la **Licenza Apache 2.0**. Consulta il file `LICENSE` per maggiori dettagli.
