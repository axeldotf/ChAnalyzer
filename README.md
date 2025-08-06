# 📈 ChAnalyzer - Report Generator

**ChAnalyzer** è un'applicazione standalone in Python per la **generazione automatizzata di report Excel (.xlsx)** a partire da file di misura (Excel) contenenti parametri di copertura 5G e legacy. Il tool raggruppa i dati per area misurata, tecnologia e operatore (TIM, Vodafone, Wind3, Iliad), con interfaccia grafica intuitiva e formattazione avanzata.

---

## 🧩 Funzionalità Principali

* **Interfaccia grafica (GUI)**: Applicazione user-friendly basata su Tkinter.
* **Supporto 5G e legacy**: Estrazione e media delle misure RSRP/RSCP/Rx Level per dati legacy e SS-RSRP per 5G.
* **Mappatura automatica**: ARFCN/Group viene tradotto in Operatore e Tecnologia tramite tabelle predefinite.
* **Calcolo medio misure**: Aggregazione e media dei valori per ciascuna combinazione Operatore–Tecnologia.
* **Formattazione condizionale**: Colori (Verde, Giallo, Rosso, Grigio) per evidenziare intervalli di copertura, con legenda integrata.
* **Elaborazione asincrona**: Mantiene la GUI reattiva durante la generazione del report.

---

## 🖥️ Requisiti

* **Python 3.7+**
* **Librerie Python**:

  ```bash
  pip install pandas xlsxwriter tqdm
  ```
* **Tkinter** incluso nella distribuzione standard di Python.

---

## ▶️ Utilizzo

1. Posizionati nella directory contenente `chanalyzer.py`.
2. Esegui il comando:

   ```bash
   python chanalyzer.py
   ```
3. Nell'interfaccia:

   1. Seleziona il file di output `.xlsx` (pulsante "Seleziona Output").
   2. Abilita o disabilita **Includi 5G** a seconda dei dati da processare.
   3. Aggiungi uno o più file di input Excel e assegna un nome al foglio ("Area Misurata").
   4. Premi **Genera Riepilogo** per avviare l'elaborazione.

Al completamento, troverai nel percorso selezionato un file Excel con un foglio `Riepilogo` contenente:

* Colonne:

  * `Area Misurata`
  * `Tecnologia`
  * `Listino Inwit` (VOCE/DATI/DATI PLUS/5G)
  * `Copertura TIM`, `Copertura VF`, `Copertura W3`, `Copertura Iliad`
* Formattazione condizionale per evidenziare la qualità della copertura.
* Legenda colori per interpretare gli intervalli di misura.

---

## 📦 Creazione dell'eseguibile (EXE)

Per distribuire il tool come applicazione standalone su Windows, utilizza PyInstaller. Ecco un esempio di comando:

```bash
pyinstaller --onefile --windowed --icon=chan.ico chanalyzer.py \
  --exclude PyQt5 --exclude PyQt5.sip --exclude PyQt5.QtCore
```

Al termine della procedura, nella cartella `dist` troverai `chanalyzer.exe` pronto per l'uso.

---

### ✍️ Autore

Sviluppato da **Alessandro Frullo**
In collaborazione con **Selektra Italia Srl**
