# Prompt‑Serie  
**Projekt:** Excel‑Produktrechner → Python‑Produktrechner  

---

## GLOBALER KONTEXT

### Rolle und Ziel  
Du bist ein **Senior‑Python‑Engineer**.  
Ziel: Aus der Excel‑Datei `TARIFRECHNER_KLV.xlsm` einen **modularen, reinen‑Python‑Produkt­rechner** zu erzeugen, der identische Ergebnisse liefert.  
Die Lösung besteht aus **sieben Scripts** und mehreren **.csv‑Eingabedateien** (siehe Tabelle unten).  
Arbeite **strikt schrittweise**: Erfülle jeden Task, warte auf mein **„✅“**, erst dann fahre fort.

### Deliverables
| Kürzel | Datei | Inhalt | Prüfkriterium |
|--------|-------|--------|---------------|
| EXCEL_TO_TEXT | `excel_to_text.py` | Extrahiert Zellen & Bereiche → `excelzell.csv`, `excelber.csv` | Beide CSVs existieren & haben ≥ 1 Zeile |
| VBA_TO_TEXT | `vba_to_text.py` | Extrahiert alle VBA‑Module → `Mod_*.txt` | Alle Modul‑Dateien vorhanden |
| DATA_EXTRACT | `data_extract.py` | Erzeugt `var.csv`, `tarif.csv`, `grenzen.csv`, `tafeln.csv`, `tarif.py` | Alle Dateien existieren & haben ≥ 1 Datenzeile |
| BASISFUNCT | `basfunct.py` | 1‑zu‑1‑Port der VBA‑Basisfunktionen | pytest‐Suite besteht |
| AUSFUNCT_T1 | `ausfunct.py` | Enthält `Bxt()` und abhängige Funktionen | `Bxt()`‑Test < 1 e‑6 Differenz |
| AUSFUNCT_T2 | `ausfunct2.py` | Weitere Ausgabefunktionen | Alle Funktions‑Tests bestehen |
| CLI_RUNNER | `run_calc.py` | Kommandozeilen‑Interface | `python run_calc.py --help` läuft |

### Allgemeine Regeln  
- **Sprache:** Deutsch in Doku, Variablennamen englisch (`present_value`).  
- **Qualität:** Black‑Format, Ruff‑Lint = 0 Warnungen.  
- **Antwortformat:** Jeder Task liefert **genau einen** ausführbaren Code‑Block.  
- **Fortschritt:** Warte jeweils auf mein **✅**.


---

## TASK 1 – Kontext­export (Excel → CSV)

1. **Input**: `TARIFRECHNER_KLV.xlsm` (im root)  
2. **Output**  
   - `excelzell.csv` (Spalten: Blatt, Adresse, Formel, Wert)  
   - `excelber.csv` (Spalten: Blatt, Name, Adresse)  
3. **Vorgehen**  
   - Nutze `xlwings` für Array‑Formeln.  
   - Ignoriere leere Zellen.  
   - Speichere CSVs im gleichen Verzeichnis wie die Excel.  
   - Eine Robuste Lösung für Bereiche mit fehlerhaften Bezügen nötig  
   - Voraussetzung: `xlwings` benötigt eine lokale Excel-Installation 
4. **Erfolgs‑Check**  
```python
assert Path("excelzell.csv").stat().st_size > 10_000
assert "Kalkulation" in pd.read_csv("excelzell.csv")["Blatt"].unique()
```  
5. **Lieferformat**: Vollständiger, ausführbarer Code‑Block.

---

## TASK 2 – VBA‑Export (VBA → TXT)

1. **Input**: `TARIFRECHNER_KLV.xlsm` (im root)  
2. **Output**: Je VBA‑Modul eine `Mod_*.txt`‑Datei  
3. **Vorgehen**  
   - Verwende `oletools.olevba` oder `vb2py` zum Modul‑Dump.  
   - Dateinamensschema: `Mod_<Name>.txt`.  
   - Verarbeite alle nichtleeren Code-Module, auch ohne `Sub` (z. B. mit Konsanten)    
   - Ignoriere leere Module oder Code‑Objekte (z. B. Excel‑Blatt ohne Code).  
4. **Erfolgs‑Check**  
   - Anzahl `.txt`‑Dateien ≥ Anzahl nicht‑leerer Module im VBA‑Editor.  
   - Jede Datei enthält mindestens eine `Sub` oder `Function`.

---

## TASK 3 – Daten aus Excel extrahieren

1. **Input**: Excel‑Datei `excelzell.csv`, `excelber.csv`  
2. **Output**  
   - `var.csv`  – Variablen (Blatt *Kalkulation*, A4:B9), pro Vertrag unterschiedlich  
   - `tarif.csv` – Tarifdaten (Blatt *Kalkulation*, D4:E11), für mehrere Verträge gleich  
   - `grenzen.csv` – Grenzen (Blatt *Kalkulation*, G4:H5)  
   - `tafeln.csv` – Sterbe­tabelle (Blatt *Tafeln*, Spalten A–E, Daten ab Zeile 4)  
   - `tarif.py`  – Funktion **`raten_zuschlag(zw)`** (Excel‑Formel E12)  
3. **Vorgehen**  
   - Lies jede genannte Zellgruppe aus den Input-Dateien.  
   - Speichere CSVs exakt in den genannten Spalten­formaten.  
   - Bei CSVs immer eine Spalte mit *Name* und eine mit *Wert*
   - Implementiere `raten_zuschlag(zw)` exakt gemäß Formel in Zelle E12.  
4. **Erfolgs‑Check**  
   - Alle Dateien existieren & haben ≥ 1 Datenzeile (bei Tafeln ≥ 100).  
   - `import tarif; tarif.raten_zuschlag(12)` liefert denselben Wert wie Excel‑Zelle E12.

---

## TASK 4 – Test‑Fixtures generieren

1. **Input**: Bisherige CSV‑Dateien (`var.csv`, `tarif.csv`, `grenzen.csv`, `tafeln.csv`)  
2. **Output**: Ordner `tests/` mit PyTest‑Fixtures  
3. **Vorgehen**  
   - `conftest.py` richtet Temp‑Verzeichnis & Mini‑CSV‑Samples ein.  
   - Erstelle Smoke‑Test `test_data_roundtrip.py`, der jede CSV liest & Spalten zählt.  
4. **Erfolgs‑Check**  
   - `pytest -q` läuft grün (0 errors, 0 failures).

---

## TASK 5A – Basisfunktionen übersetzen

1. **Input**: Alle `Mod_*.txt` aus TASK 2  
2. **Output**: `basfunct.py`  
3. **Vorgehen**  
   - Jede VBA‑Function/Procedure wird 1‑zu‑1 als Python‑`def` abgebildet.  
   - Nutze `pandas` für Tabellen‑/CSV‑Zugriffe.  
   - Kein Funktionsrumpf darf mit `pass` enden.  
   - Verfügbare Datenquellen: `excelzell.csv`, `excelber.csv`, `var.csv`, `tarif.csv`, `grenzen.csv`, `tafeln.csv`.

---

## TASK 5B – Funktions­paritäts‑Test

**Erfolgs‑Check**: LLM erstellt `tests/test_func_parity.py`, das  
- alle **öffentlichen** VBA‑Namen (Function/Sub ohne `Private`) einsammelt,  
- Python‑`def`‑Namen in `basfunct.py` scannt (Helper dürfen ignoriert werden),  
- und prüft, dass pro VBA‑Name genau eine Python‑Funktion existiert.  
Bestanden = `pytest -q` läuft vollständig grün.

---

## TASK 6A – Bxt()  (Beitrags­berechnung 1 / 4)

1. **Input**  
   • basfunct.py  
   • CSVs: var.csv, tarif.csv, grenzen.csv, tafeln.csv  
   • excelzell.csv & excelber.csv (für Zell-/Namens­referenzen)

2. **Output**  
   • Funktion `Bxt(vs, age, sex, n, t, zw, tarif)` in ausfunct.py

3. **Vorgehen**  
   • Formel exakt wie in Kalkulation!K5 („Bxt“).  
   • Abhängigkeiten: - Variablen → var.csv - Tarif/Grenzen → tarif.csv, grenzen.csv - Basis­funktionen → basfunct.py.  
   • Keine Platzhalter (`pass`) hinterlassen.

4. **Erfolgs-Check**  
   **Referenz­eingabe**

   | vs | age | sex | n | t | zw | tarif |
   |----|-----|-----|---|---|----|-------|
   | 100 000 | 40 | "M" | 30 | 20 | 12 | "KLV" |

   **Sollwert**

   | Funktion | Erwartet | Toleranz |
   |----------|----------|-----------|
   | Bxt() | **0.04226001** | ± 1 × 10⁻⁸ |

   *LLM erzeugt `tests/test_bxt.py`, der diesen einen Fall prüft.  
   Bestanden = `pytest -q` zeigt grünen Test.*

---

## TASK 6B – Weitere Ausgabefunktionen  (Beitrags­berechnung 2 – 4)

1. **Input** wie 6A  
2. **Output** ausfunct.py (erweitert) **oder** neues ausfunct2.py  
   – implementiere Funktionen für Zellen **K6 (BJB), K7 (BZB), K9 (Pxt)**.  

3. **Erfolgs-Check**  
   **Referenz­eingabe** – identisch zur Tabelle oben.  

   **Sollwerte**

   | Funktion | Excel-Zelle | Erwartet  | Toleranz |
   |----------|------------|-----------|----------|
   | BJB() | K6 | **4 226.00** € | ± 1 × 10⁻² |
   | BZB() | K7 | **371.88** €  | ± 1 × 10⁻² |
   | Pxt() | K9 | **0.04001217** | ± 1 × 10⁻⁸ |

   *LLM legt `tests/test_outputs.py` an und prüft jede Funktion mit diesen Sollwerten.  
   Bestanden = `pytest -q` komplett grün.*

---

## TASK 6C – Verlaufswerte  (11 Funktionen)

1. **Input**  
   • Fertige Beitragsfunktionen (6A, 6B)  
   • basfunct.py + alle CSVs

2. **Output**  
   • Funktion `verlaufswerte(vs, age, sex, n, t, zw, tarif)` in ausfunct.py  
     – DataFrame mit Spalten **B15:L15** des Blatts „Kalkulation“.

3. **Vorgehen**  
   • Berechne für jedes *k = 0 … n* alle Verlaufsgrößen (`Axn`, `axn`, `axt`, `kVx_bpfl`, …).  
   • Spaltennamen & Reihenfolge exakt wie in Excel.

4. **Erfolgs-Check**  
   **Referenz­eingabe** – gleich wie oben.  

   **Sollwerte (Auszug)**

   | k | Axn       | axn        | axt        | Toleranz |
   |---|-----------|------------|------------|----------|
   | 0 | **0.6315923** | **21.4202775** | **16.3130941** | ± 1 × 10⁻⁶ |
   | 1 | **0.6417247** | **20.8311476** | **15.6212042** | ± 1 × 10⁻⁶ |

   *LLM generiert `tests/test_verlauf.py`:  
   – ruft `verlaufswerte()` mit den Referenz­parametern auf,  
   – prüft Spalten `Axn`, `axn`, `axt` für k = 0 und 1.  
   Bestanden = `pytest -q` ohne Fehler.*

---

## TASK 7 – CLI-Runner (`run_calc.py`)

> *Ziel:* Kommandozeilen-Tool, das alle oder ausgewählte Ausgabefunktionen
> (derzeit funktionsfähig nur **Bxt**, die restlichen Platzhalter) berechnet.
> Standard-Inputs werden aus den CSV-Dateien geladen; per Argument lassen sie sich übersteuern.

1. **Input**
   - Python-Module: `ausfunct.py` (Bxt funktionsfähig, 6B/6C Platzhalter)
   - CSV-Dateien: `var.csv`, `tarif.csv`, `grenzen.csv`, `tafeln.csv`

2. **Output**
   - Script **`run_calc.py`**
   - Vollständige `--help`-Ausgabe via `argparse`

3. **Vorgehen**
   - **argparse-Parameter**

     | Typ | Parameter | Beschreibung |
     |-----|-----------|--------------|
     | Datei | `--var-file` *(default `var.csv`)* | alternative Variablen-Datei |
     | Datei | `--tarif-file` *(default `tarif.csv`)* | alternative Tarif-Datei |
     | Liste | `--funcs` *(z. B. `Bxt,BJB`)* | nur diese Funktionen ausführen |
     | Flag  | `--all` *(default)* | alle verfügbaren Funktionen berechnen |

   - Lade `var.csv`, `tarif.csv` (bzw. alternative Pfade) per `pandas`.
     * benötigte Variablen: **vs, age, sex, n, t, zw, tarif**  
       → lies sie case-insensitiv aus den Spalten `Variable`, `Wert`.
   - Bestimme `funcs_to_run = ALL_FUNCS if args.all or not args.funcs else args.funcs.split(",")`
   - **Dynamischer Import**  
     ```python
     from importlib import import_module
     funcs = {name: getattr(import_module("ausfunct"), name, None) for name in funcs_to_run}
     ```
   - Für Platzhalter oder fehlende Funktionen:
     ```python
     if func is None or func.__doc__ == "PLACEHOLDER":
         results[name] = "not yet implemented"
     else:
         results[name] = func(**input_args)
     ```
   - Gib `results` als **kompaktes JSON** (`json.dumps(..., ensure_ascii=False)`) auf STDOUT.

4. **Erfolgs-Check**
   - **Default-Call** (liest `var.csv`, `tarif.csv`):  
     ```bash
     python run_calc.py
     # → {"Bxt": 0.04226001, "BJB": "not yet implemented", ...}
     ```
   - **Teilmenge-Call**  
     ```bash
     python run_calc.py --funcs Bxt
     # → {"Bxt": 0.04226001}
     ```
   - **Alternative Variablen-Datei**  
     ```bash
     python run_calc.py --var-file my_vars.csv --funcs Bxt
     ```
     → verwendet Werte aus `my_vars.csv`.
   - `python run_calc.py --help` listet alle Optionen mit Beschreibung & Beispiel.

*Bestanden* = alle drei Aufrufe funktionieren ohne Traceback;  
Bxt-Wert entspricht Referenz (0.04226001 ± 1e-8); nicht implementierte Funktionen melden `"not yet implemented"`.
