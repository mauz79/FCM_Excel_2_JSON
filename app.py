#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FCM – Excel → JSON (stand-alone, GUI per Windows) – FreeSimpleGUI

Funzionalità:
- Legge file Excel (.xls/.xlsx) con foglio "Tutti i dati" prodotti da FCM (FantaCalcio Manager)
- Verifica le 34 colonne obbligatorie (tutte incluse nel JSON, senza rinomini)
- Estrae stagione dal filename (YYYY_YYYY | YYYY-YYYY | YYYY/YYYY)
- Scrive un file per stagione: "YYYY_YYYY.json" + aggiorna "seasons.json"
- Log a schermo e in "conversion.log" (nella cartella di output)
- Opzione "Modalità RAW" per saltare la normalizzazione
- NUOVO: selezione di singoli file (anche multipli) OPPURE cartella input

Dipendenze:
- pandas==2.2.2, openpyxl==3.1.5, xlrd==2.0.1, FreeSimpleGUI==5.2.0.post1
"""

import re
import json
from pathlib import Path
from datetime import datetime

import pandas as pd
import FreeSimpleGUI as sg


# ====== Costanti e configurazione ======

SHEET_NAME = "Tutti i dati"

REQUIRED_COLUMNS = [
    "Nome","Sq","R","COD","FMld","T","P","Aff%",
    "MVC","MVF","MVT","MVDSt","MVDlt","MVAnd","MVRnd",
    "FMC","FMF","FMT","FMDSt","FMDlt","FMAnd","FMRnd",
    "GF","GFR","GS","GSR","AG","AS","RP","RS","A","E","TIn","ID"
]

# Colonne tipizzate per normalizzazione (se RAW è disattivato)
FLOAT_COLS = {
    "FMld","Aff%","MVC","MVF","MVT","MVDSt","MVDlt","MVAnd","MVRnd",
    "FMC","FMF","FMT","FMDSt","FMDlt","FMAnd","FMRnd"
}
INT_COLS = {"T","P","GF","GFR","GS","GSR","AG","AS","RP","RS","A","E","TIn"}
STR_COLS = {"Nome","Sq","R","COD","ID"}  # ID trattato come stringa


# ====== Utility ======

def extract_season_from_filename(stem: str):
    """
    Estrae la stagione dal nome file con pattern:
      - 2021_2022
      - 2021-2022
      - 2021/2022
    Ritorna: (season_label '2021/2022', season_key '2021_2022')
    """
    m = re.search(r"(20\d{2})\s*[_/\-]\s*(20\d{2})", stem)
    if not m:
        raise ValueError("Impossibile estrarre la stagione dal nome file (atteso pattern YYYY_YYYY)")
    y1, y2 = m.group(1), m.group(2)
    return f"{y1}/{y2}", f"{y1}_{y2}"


def read_excel_with_engine(fp: Path, sheet_name: str):
    """
    Usa openpyxl per .xlsx e xlrd per .xls (strada rapida, senza Rust).
    """
    suffix = fp.suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(fp, sheet_name=sheet_name, engine="openpyxl")
    else:  # .xls (o altro legacy -> proviamo xlrd)
        return pd.read_excel(fp, sheet_name=sheet_name, engine="xlrd")


def normalize_df(df: pd.DataFrame):
    """
    Normalizza tipi:
    - Trim stringhe note
    - Float: converte virgola -> punto, rimuove %, caratteri non numerici; arrotonda a 2 decimali
    - Interi: coerzione numerica, NaN -> 0
    Mantiene tutte le colonne (anche eventuali extra).
    """
    # Trim
    for c in df.columns:
        if c in STR_COLS:
            df[c] = df[c].astype(str).str.strip()

    # Float
    for c in df.columns:
        if c in FLOAT_COLS:
            s = df[c].astype(str).str.strip()
            s = (
                s.str.replace("%", "", regex=False)
                 .str.replace(",", ".", regex=False)
                 .str.replace("–", "", regex=False)
                 .str.replace("-", "", regex=False)
            )
            df[c] = pd.to_numeric(s, errors="coerce").round(2)

    # Interi
    for c in df.columns:
        if c in INT_COLS:
            s = pd.to_numeric(df[c], errors="coerce")
            df[c] = s.fillna(0).astype(int)

    return df


def ensure_required_columns(df: pd.DataFrame):
    """Ritorna la lista di colonne mancanti rispetto a REQUIRED_COLUMNS."""
    return [c for c in REQUIRED_COLUMNS if c not in df.columns]


def write_json(path: Path, data: dict):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def append_logfile(logfile: Path, msg: str):
    try:
        with logfile.open("a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def _process_files(files, output_dir: Path, window, raw_mode=False):
    """
    Core: elabora la lista di file Excel passati.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    seasons = []
    now_iso = datetime.utcnow().isoformat(timespec="seconds") + "Z"

    # Filtra solo .xls/.xlsx esistenti
    files = [Path(p) for p in files]
    files = [p for p in files if p.is_file() and p.suffix.lower() in (".xls", ".xlsx")]
    files = sorted(files, key=lambda p: p.name.lower())

    logfile = output_dir / "conversion.log"

    def log(*args):
        line = " ".join(str(a) for a in args)
        window["-LOG-"].print(line)
        append_logfile(logfile, line)

    if not files:
        log("[INFO] Nessun file .xls/.xlsx valido selezionato.")
        return

    window["-PBAR-"].update(0, max=len(files))
    seen_seasons = set()

    for i, fp in enumerate(files, start=1):
        window["-PBAR-"].update(i)
        log(f"Leggo: {fp.name}")

        # Estrazione stagione da filename
        try:
            season_label, season_key = extract_season_from_filename(fp.stem)
        except Exception as e:
            log(f"  [WARN] {e} -> file saltato")
            continue

        if season_key in seen_seasons:
            log(f"  [WARN] Stagione duplicata '{season_key}' (verrà sovrascritta con questo file).")
        seen_seasons.add(season_key)

        # Lettura Excel
        try:
            df = read_excel_with_engine(fp, SHEET_NAME)
        except Exception as e:
            log(f"  [ERRORE] Impossibile leggere il foglio '{SHEET_NAME}': {e}")
            continue

        # Validazione colonne
        missing = ensure_required_columns(df)
        if missing:
            log(f"  [ERRORE] Colonne mancanti: {missing} -> file saltato")
            continue

        # Normalizzazione (se RAW disattivato)
        if not raw_mode:
            df = normalize_df(df)

        # Mantieni tutte le colonne presenti (ordine del DataFrame)
        cols = list(df.columns)

        out = {
            "schema_version": 1,
            "season_label": season_label,  # es. "2021/2022"
            "season_key": season_key,      # es. "2021_2022" (safe per filename/URL)
            "generated_at": now_iso,
            "columns": cols,
            "players": df.to_dict(orient="records"),
        }

        out_path = output_dir / f"{season_key}.json"
        try:
            write_json(out_path, out)
        except Exception as e:
            log(f"  [ERRORE] Scrittura JSON {out_path.name}: {e}")
            continue

        seasons.append({
            "label": season_label,
            "key": season_key,
            "file": out_path.name,
            "n_players": int(len(df)),
            "last_updated": now_iso
        })
        log(f"  [OK] Generato {out_path.name} ({len(df)} righe)")

    # Aggiorna manifest
    if seasons:
        dedup = {s["key"]: s for s in seasons}
        seasons_sorted = [dedup[k] for k in sorted(dedup.keys())]
        manifest = {"schema_version": 1, "seasons": seasons_sorted}
        try:
            write_json(output_dir / "seasons.json", manifest)
            log(f"[OK] Aggiornato seasons.json ({len(seasons_sorted)} stagioni)")
        except Exception as e:
            log(f"[ERRORE] Scrittura seasons.json: {e}")
    else:
        log("[FINE] Nessun JSON generato (nessun file valido).")


# ====== GUI ======

def main():
    # Tema (compat)
    try:
        if hasattr(sg, "theme"):
            sg.theme("DarkBlue3")
        elif hasattr(sg, "change_look_and_feel"):
            sg.change_look_and_feel("DarkBlue3")
    except Exception:
        pass

    layout = [
        [sg.Text("File Excel (uno o più)"),
         sg.Input(key="-FILES-"),
         sg.FilesBrowse("Scegli file…", key="-FILESBR-",
                        file_types=(("Excel", "*.xls;*.xlsx"),),
                        files_delimiter=";")],
        [sg.Text("OPPURE: Cartella input (Excel)"),
         sg.Input(key="-IN-"),
         sg.FolderBrowse()],
        [sg.Text("Cartella output (JSON)"),
         sg.Input(key="-OUT-"),
         sg.FolderBrowse()],
        [sg.Checkbox("Modalità RAW (non convertire numeri/percentuali)", default=False, key="-RAW-")],
        [sg.Text("Foglio (fisso):"),
         sg.Input(SHEET_NAME, key="-SHEET-", size=(30, 1), disabled=True)],
        [sg.ProgressBar(max_value=100, orientation='h', size=(50, 20), key='-PBAR-')],
        [sg.Button("Converti", key="-RUN-", button_color=("white", "#2563eb")),
         sg.Button("Apri output"),
         sg.Button("Chiudi")],
        [sg.Multiline(size=(110, 22), key="-LOG-", autoscroll=True, disabled=True, write_only=True)]
    ]

    window = sg.Window("FCM_Excel_2_JSON (FreeSimpleGUI)", layout, finalize=True)
    window["-PBAR-"].update(0, max=100)

    while True:
        ev, vals = window.read()
        if ev in (sg.WINDOW_CLOSED, "Chiudi"):
            break

        if ev == "-RUN-":
            # Priorità: se l'utente ha selezionato file specifici, usiamo quelli
            files_str = (vals.get("-FILES-") or "").strip()
            in_dir = Path(vals["-IN-"]) if vals.get("-IN-") else None
            out_dir = Path(vals["-OUT-"]) if vals.get("-OUT-") else None
            raw_mode = bool(vals.get("-RAW-"))
            window["-LOG-"].update("")

            if not out_dir:
                sg.popup_error("Seleziona una cartella output")
                continue

            if files_str:
                # Split robusto: ';' (Windows), virgola, newline
                parts = re.split(r"[;\n,]+", files_str)
                file_paths = [Path(p.strip().strip('"')) for p in parts if p.strip()]
                _process_files(file_paths, out_dir, window, raw_mode=raw_mode)
            else:
                if not in_dir or not in_dir.exists():
                    sg.popup_error("Seleziona una cartella input valida oppure scegli uno o più file in alto")
                    continue
                # Leggi tutti i .xls/.xlsx nella cartella
                dir_files = sorted([p for p in in_dir.glob("*.xls*") if p.is_file()], key=lambda p: p.name.lower())
                _process_files(dir_files, out_dir, window, raw_mode=raw_mode)

        if ev == "Apri output":
            path = vals.get("-OUT-")
            if path and Path(path).exists():
                try:
                    import os
                    os.startfile(path)  # Windows-only
                except Exception:
                    sg.popup_error("Impossibile aprire la cartella output.")
            else:
                sg.popup_error("Seleziona una cartella output valida.")

    window.close()


if __name__ == "__main__":
    main()
