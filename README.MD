# 📦 FCM_Excel_2 JSON

Convertitore di file Excel (generati dalle funzioni TOP di FCM) in **file JSON** compatibili con **AstaMasterPro**.

✅ **Funzionalità principali**
- Legge file **.xls** e **.xlsx** (foglio fisso: `Tutti i dati`)
- Include **tutte le 34 colonne** senza rinominarle
- Genera:
  - `YYYY_YYYY.json` per ogni stagione
  - `seasons.json` (manifest)
  - `conversion.log` (log dettagliato)
- Opzione **Modalità RAW** (nessuna normalizzazione)
- **Selezione multipla di file** oppure **intera cartella**

> Contenuto della cartella:

FCM_Excel_2_JSON/
├─ app.py                  # Codice sorgente (GUI + logica)
├─ requirements.txt        # Dipendenze Python
├─ build_windows.bat       # Script per generare l'EXE
└─ dist/                   # Output della build

---

## ✅ Requisiti
- **Windows 10/11**
- **Python 3.12** consigliato (3.13 funziona, ma può dare warning con Pandas)
- **Microsoft Visual C++ Build Tools** (per alcune librerie come Pandas/Numpy)  
  👉 [Scarica qui](https://visualstudio.microsoft.com/visual-cpp-build-tools/)

---

## 1. Installare Python
- Per prima cosa, installa Python. Ti consigliamo di usare la versione ufficiale, scaricabile dal sito di Python.
- Vai su https://www.python.org/downloads/.
- Clicca sul pulsante "Download Python 3.x.x" per scaricare l'installer.
- Esegui il file .exe scaricato.

> Molto importante: Assicurati di spuntare la casella "Add python.exe to PATH" prima di cliccare su "Install Now".
> Questa opzione ti permette di eseguire i comandi Python direttamente dal terminale.

---

## 2. Installare Microsoft Visual C++ Build Tools

- Alcune librerie Python, come pandas o numpy, richiedono un compilatore C++ per essere installate su Windows. Se non lo hai, riceverai un errore durante l'installazione delle dipendenze.
- Scarica Visual Studio Build Tools dal sito ufficiale di Microsoft: https://visualstudio.microsoft.com/visual-cpp-build-tools/.
- Esegui l'installer.

- Quando ti viene chiesto di selezionare i "Carichi di lavoro" (workloads), spunta la casella "Sviluppo per desktop con C++".
- Clicca su "Installa" e attendi il completamento del processo.
- Una volta installato tutto, puoi aprire il terminale e procedere con l'installazione delle dipendenze del progetto usando pip install -r requirements.txt.

---

## A) 🔧 Installazione e build manuale (consigliata la prima volta)

### 0) Prerequisiti
- **Windows 10/11**  
- **Python 3.12** consigliato (3.13 va spesso bene, ma 3.12 è più “tranquillo” con Pandas).  
- Nella tua cartella progetto (es. `D:\sandbox\apptojson`) assicurati di avere:
  - `app.py`
  - `requirements.txt` con questo contenuto:

```txt

pandas==2.2.2
openpyxl==3.1.5
xlrd==2.0.1
FreeSimpleGUI==5.2.0.post1
```

> xlrd 2.xsupporta i .xls (non .xlsx), mentre per gli .xlsx usiamo openpyxl—è la combinazione supportata/attuale.

---

> La cartella del progetto è D:\sandbox\apptojson

### 1) Apri `cmd.exe` nella cartella del progetto
```bat
cd D:\sandbox\apptojson
```

### 2) Crea e attiva il virtualenv
```bat
python -m venv .venv
.venv\Scripts\activate.bat
```
Vedrai il prompt con `(.venv)` all’inizio.

> **Se vuoi usare Python 3.12 esplicitamente**:
> ```bat
> py -3.12 -m venv .venv
> .venv\Scripts\activate.bat
> ```

### 3) Installa le dipendenze (FreeSimpleGUI + Pandas + Openpyxl + python-calamine)
```bat
pip install --upgrade pip
pip install -r requirements.txt
```

### 4) Installa PyInstaller
```bat
pip install pyinstaller
```

### 5) Compila l’EXE (stand‑alone, con finestra GUI)
```bat
pyinstaller --noconfirm --clean --onefile --windowed app.py --name "FCM_Excel_2_JSON"
```

### 6) Dove trovo l’EXE?
- Verrà creato in:
  ```
  dist\FCM_Excel_2_JSON.exe
  ```
- Puoi **doppio‑cliccare** per avviare l’app.

``

---

## B) ✅ Script automatico con `build_windows.bat`


Puoi usare `build_windows.bat` per:
- Creare il venv
- Installare le dipendenze
- Compilare l’EXE
- Aprire la cartella `dist`

Esegui:
```bat
build_windows.bat
```

```
---

## ▶️ Uso dell’app
1. Avvia `FCM_Excel_2 JSON.exe`
2. **Scegli uno o più file Excel** (anche mix `.xls`/`.xlsx`) **oppure** una cartella
3. Seleziona la **cartella di output**
4. Premi **Converti**
5. Output:
   - `YYYY_YYYY.json` per ogni file valido
   - `seasons.json`
   - `conversion.log`

---

## 🎯 Test rapido dell’EXE

1) Avvia `dist\FCM_Excel_2_JSON.exe`  
2) **Cartella input**: quella con i file tipo  
   `Nome_Lega 2021_2022_Risultati Tutti i dati.xls`  
3) **Cartella output**: una cartella vuota per i JSON  
4) Premi **Converti**  
5) In output troverai:
   - `2021_2022.json`, `2022_2023.json`, …  
   - `seasons.json`  
   - `conversion.log`

**Dentro a `2021_2022.json`** controlla:
- `season_label`: `2021/2022`
- `season_key`: `2021_2022`
- `columns`: tutte le **34 colonne**
- `players`: ogni riga con le 34 colonne (in **RAW** = come Excel; in **Normalizzato** = numeri puliti)

---

## 🛠 Troubleshooting
- **SmartScreen dice “Autore sconosciuto”** → clic su **Ulteriori informazioni → Esegui comunque**
- **Errore leggendo `.xls`** → assicurati di avere `xlrd==2.0.1`:
  ```bat
  pip uninstall -y xlrd
  pip install xlrd==2.0.1
  ```
- **Problemi con Python 3.13** → usa Python 3.12
- **GUI non parte** → verifica che in `requirements.txt` ci sia:
  ```python
  import FreeSimpleGUI as sg
  ```
