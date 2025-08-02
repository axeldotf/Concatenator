# 📑 Concatenator - Report Generator

**Concatenator** è uno strumento standalone in Python progettato per la **generazione automatizzata di report Word (.docx)** a partire da immagini tematiche. Il tool consente di **raggruppare immagini per blocchi e per operatore telefonico (Iliad, TIM, Vodafone, Wind3)**, con interfaccia grafica intuitiva e opzioni avanzate di personalizzazione.

---

## 🧩 Funzionalità Principali

- **Interfaccia grafica (GUI)**: Applicazione user-friendly basata su Tkinter.
- **Suddivisione automatica per operatore**: Le immagini vengono assegnate agli operatori in base al nome file.
- **Blocchi tematici**: Le immagini possono essere organizzate in più blocchi con titoli personalizzati.
- **Ritaglio intelligente**: Opzioni per rimuovere bordi bianchi *lateralmente*, *verticalmente* o *entrambi*.
- **Etichettatura opzionale**: Inserimento automatico del nome operatore e tecnologia sopra ogni immagine.
- **Generazione Word multi-documento**: Un file `.docx` per ogni operatore, con layout ottimizzato e immagini scalate correttamente.
- **Elaborazione asincrona**: Il tool mantiene responsiva la GUI durante la generazione dei report.

---

## 🖥️ Requisiti

Assicurati di avere Python 3 installato e installa i pacchetti richiesti:

```bash
pip install python-docx pillow tqdm
````

## ▶️ Utilizzo

Posizionati nella directory contenente il file `concatenator.py` e avvia l'applicazione eseguendo il seguente comando da terminale:

```bash
python concatenator.py
````

### 🧭 Procedura guidata nell’interfaccia

1. **Inserisci il titolo del documento**
2. **Scegli la modalità di ritaglio** delle immagini:
   - `none`
   - `sides`
   - `topbottom`
   - `both`
3. **Abilita o disabilita l’etichetta automatica**
4. **Seleziona la cartella di output**
5. **Aggiungi uno o più blocchi di immagini** con titolo personalizzato
6. Premi **“Genera Documenti”** per avviare l’elaborazione

✅ Al termine, nella cartella selezionata troverai **un file Word per ciascun operatore** contenente tutte le immagini raggruppate per blocco.

---

### 🖼️ Output

📁 `TitoloDocumento_Iliad.docx`  
📁 `TitoloDocumento_TIM.docx`  
📁 `TitoloDocumento_VF.docx`  
📁 `TitoloDocumento_W3.docx`

Ogni documento conterrà:
- Immagini scalate e formattate
- Opzionalmente ritagliate e con intestazione
- Ordinamento per tecnologia secondo la sequenza definita in `ORDER`

---

### ✍️ Autore

Sviluppato da **Alessandro Frullo**  
In collaborazione con **Selektra Italia Srl**
