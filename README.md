# Concatenator

Concatenator è un tool Python progettato per la generazione automatizzata di reportistica di progetto in formato Word (.docx). Il programma permette di selezionare immagini, suddividerle per blocchi tematici e raggrupparle in documenti suddivisi per operatore (Iliad, TIM, Vodafone e Wind3). 

## Funzionalità principali
- **Selezione delle immagini**: Gli utenti possono selezionare manualmente le immagini da includere nel report.
- **Filtraggio automatico**: Le immagini vengono automaticamente suddivise per operatore in base al nome del file.
- **Ritaglio bordi bianchi**: Opzionalmente, le immagini possono essere ritagliate per rimuovere eventuali bordi bianchi.
- **Creazione documenti Word**: Viene generato un documento per ciascun operatore, con le immagini organizzate per sezioni e formattate correttamente.
- **Elaborazione multipla**: È possibile inserire più blocchi tematici di immagini all'interno di un unico report.

## Installazione
Per utilizzare Concatenator, è necessario avere Python installato sul proprio sistema insieme alle seguenti librerie:

```sh
pip install python-docx pillow tqdm
```

## Utilizzo
Eseguire il file `concatenator.py` e seguire le istruzioni a schermo:

```sh
python concatenator.py
```

1. Inserire il titolo del documento.
2. Scegliere se ritagliare automaticamente i bordi bianchi dalle immagini.
3. Selezionare la cartella di destinazione per i documenti generati.
4. Aggiungere i blocchi di immagini fornendo un titolo per ciascun gruppo.
5. Il programma genererà i documenti Word per ciascun operatore contenente le immagini organizzate in sezioni.

## Output
Alla fine del processo, il software creerà file Word separati per ciascun operatore, salvandoli nella cartella selezionata dall'utente.

## Autori
Questo software è stato sviluppato da **Alessandro Frullo** in collaborazione con l'azienda **Selektra Italia Srl**.
