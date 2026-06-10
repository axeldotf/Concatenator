# 📑 ReportGenerator - Concatenator v2

**ReportGenerator** è uno strumento standalone in Python progettato per la **generazione automatizzata di report Word (.docx)** a partire da immagini tematiche.

Il tool consente di **raggruppare immagini per blocchi e per operatore telefonico** — **Iliad, TIM, Vodafone/VF e Wind3/W3** — attraverso una **interfaccia grafica intuitiva**, con opzioni avanzate per ritaglio, ordinamento, etichettatura e generazione automatica dei documenti.

La versione **v2** introduce una GUI più completa e ottimizzata, una migliore gestione delle immagini, una visualizzazione delle sottocartelle 4G/5G e una generazione più efficiente dei report.

---

## 🧩 Funzionalità Principali

* **Interfaccia grafica (GUI)**: applicazione user-friendly basata su Tkinter.
* **Layout grafico personalizzato**: interfaccia con palette Selektra, header dedicato, pulsanti stilizzati e footer autore.
* **Suddivisione automatica per operatore**: le immagini vengono assegnate agli operatori in base al nome file.
* **Blocchi tematici**: le immagini possono essere organizzate in più blocchi con titoli personalizzati.
* **Ritaglio intelligente**: opzioni per rimuovere bordi bianchi *lateralmente*, *verticalmente* o *entrambi*.
* **Ottimizzazione del ritaglio immagini**: utilizzo di funzioni native Pillow per rendere il processo più rapido.
* **Etichettatura opzionale**: inserimento automatico del nome operatore e tecnologia sopra ogni immagine.
* **Ordinamento automatico per tecnologia**: le immagini vengono ordinate secondo la sequenza definita in `ORDER`.
* **Generazione Word multi-documento**: viene creato un file `.docx` per ogni operatore.
* **Supporto a documenti già esistenti**: se il file Word dell’operatore è già presente, il tool può aprirlo e aggiungere nuovi blocchi.
* **Formato documento ottimizzato**: layout orizzontale, margini ridotti e immagini scalate in base alla pagina.
* **Elaborazione asincrona**: la generazione avviene in background, mantenendo la GUI responsiva.
* **Barra di avanzamento**: visualizzazione dello stato di completamento della generazione.
* **Visualizzazione sottocartelle output**: elenco delle sottocartelle presenti nella directory selezionata, con evidenza grafica per cartelle che contengono dati 4G e/o 5G.

---

## 🖥️ Requisiti

Assicurati di avere **Python 3** installato e installa i pacchetti richiesti:

```bash
pip install python-docx pillow tqdm
```

---

## ▶️ Utilizzo

Posizionati nella directory contenente il file `concatenator_v2.py` e avvia l'applicazione eseguendo il seguente comando da terminale:

```bash
python concatenator_v2.py
```

---

## 🧭 Procedura guidata nell’interfaccia

1. **Inserisci il titolo del documento**
2. **Scegli la modalità di ritaglio** delle immagini:

   * `none`
   * `sides`
   * `topbottom`
   * `both`
3. **Abilita o disabilita l’etichetta automatica**
4. **Seleziona la cartella di output**
5. **Verifica le sottocartelle disponibili**, con indicazione visiva della presenza di contenuti 4G e/o 5G
6. **Aggiungi uno o più blocchi di immagini** con titolo personalizzato
7. Premi **“Genera Documenti”** per avviare l’elaborazione

✅ Al termine, nella cartella selezionata troverai **un file Word per ciascun operatore** contenente tutte le immagini raggruppate per blocco.

---

## 🖼️ Output

I documenti generati seguono questa struttura:

```text
TitoloDocumento_Iliad.docx
TitoloDocumento_TIM.docx
TitoloDocumento_VF.docx
TitoloDocumento_W3.docx
```

Ogni documento conterrà:

* Immagini scalate e formattate automaticamente
* Eventuale ritaglio dei bordi bianchi
* Eventuale intestazione con operatore e tecnologia
* Suddivisione per blocchi tematici
* Ordinamento per tecnologia secondo la sequenza definita in `ORDER`
* Layout Word in formato orizzontale con margini ridotti

---

## ⚙️ Logica di ordinamento

Le immagini vengono ordinate in base alla sequenza definita nella lista `ORDER`, che include tecnologie e metriche come:

```text
GSM900 RXLEV
LTE800 RSRP
LTE800 QUAL
UMTS900 RSCP
UMTS900 QUAL
LTE1800 RSRP
LTE1800 QUAL
LTE2100 RSRP
LTE2100 QUAL
LTE2600 RSRP
LTE2600 QUAL
RSRP 700
RSRQ 700
RSRP 3500
RSRQ 3500
5G SS-RSRP
5G SS-RSRQ
```

Le immagini che non corrispondono a nessuna voce della lista vengono inserite in fondo.

---

## 🚀 Novità della versione v2

La versione **v2** introduce diverse migliorie rispetto alla prima versione:

* Nuovo nome applicativo: **ReportGenerator - Selektra Italia**
* GUI più ampia e strutturata
* Header grafico personalizzato
* Footer con dicitura **“Creato da Alessandro Frullo”**
* Lista sottocartelle della directory di output
* Evidenza cromatica delle cartelle contenenti sottocartelle 4G e/o 5G
* Barra di avanzamento durante la generazione
* Pulsante principale evidenziato per la generazione dei documenti
* Ritaglio immagini più veloce tramite `ImageChops`
* Gestione font ottimizzata tramite cache
* Elaborazione immagini in memoria senza creare file temporanei
* Supporto all’apertura e aggiornamento di documenti Word già esistenti
* Migliore gestione degli errori durante la generazione
* Interfaccia più coerente con la palette grafica Selektra

---

## 📦 Creazione dell'eseguibile (EXE)

Per distribuire il tool come applicazione standalone su Windows, puoi generare un eseguibile usando `pyinstaller`.

Nel progetto è incluso un file di icona `conc.ico`. Ecco il comando consigliato:

```bash
pyinstaller --onefile --windowed --icon=conc.ico concatenator_v2.py \
  --exclude PyQt5 \
  --exclude PyQt5.sip \
  --exclude PyQt5.QtCore
```

Opzioni utilizzate:

* `--onefile`: raggruppa tutto in un singolo file eseguibile.
* `--windowed`: disabilita la console a terminale e utilizza solo la GUI.
* `--icon=conc.ico`: imposta l'icona dell'applicazione.
* `--exclude ...`: esclude moduli non necessari, riducendo le dimensioni del file finale.

Al termine della procedura, nella cartella `dist` troverai:

```text
concatenator_v2.exe
```

Il file sarà pronto per l'utilizzo su Windows.

---

## 📁 Struttura consigliata del progetto

```text
Concatenator/
│
├── concatenator_v2.py
├── conc.ico
└── README.md
```

---

## ✍️ Autore

Sviluppato da **Alessandro Frullo**
In collaborazione con **Selektra Italia Srl**
