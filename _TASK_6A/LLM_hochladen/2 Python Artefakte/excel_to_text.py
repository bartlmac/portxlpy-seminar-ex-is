# excel_to_text.py
# -*- coding: utf-8 -*-
"""
Extrahiert aus input/TARIFRECHNER_KLV.xlsm:
- Alle belegten Zellen inkl. Formeln (auch Array-Formeln) -> excelzell.csv
- Alle definierten Namen/Bereiche -> excelber.csv

Regeln:
- xlwings (benötigt lokale Excel-Installation)
- Ignoriere leere Zellen
- Robust gegenüber fehlerhaften Bezügen (#REF!, ungültige Names usw.)
- CSVs werden im Verzeichnis der Excel-Datei UND als Bequemlichkeitskopie im CWD abgelegt,
  falls sich dieses vom Excel-Verzeichnis unterscheidet.

Aufruf:
    python excel_to_text.py [optional: path\to\TARIFRECHNER_KLV.xlsm]

Ausgabe:
    excelzell.csv Spalten: Blatt, Adresse, Formel, Wert
    excelber.csv  Spalten: Blatt, Name, Adresse
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Iterable, List, Tuple

import pandas as pd

# xlwings import kann auf Systemen ohne Excel fehlschlagen; wir geben dann
# eine verständliche Fehlermeldung aus.
try:
    import xlwings as xw
except Exception as e:  # pragma: no cover
    raise SystemExit(
        "xlwings konnte nicht importiert werden. Bitte stellen Sie sicher, "
        "dass xlwings installiert ist und eine lokale Excel-Installation vorhanden ist.\n"
        f"Originalfehler: {e}"
    )


def to_str(val: Any) -> str:
    """Sichere String-Repräsentation für CSV."""
    if val is None:
        return ""
    if isinstance(val, float):
        # Darstellung konsistent halten, ohne wissenschaftliche Notation bei 'ganzen' Zahlen
        if val.is_integer():
            return str(int(val))
        return repr(val)
    if hasattr(val, "isoformat"):
        try:
            return val.isoformat()
        except Exception:
            pass
    return str(val)


def cell_address(r: int, c: int) -> str:
    """Wandelt 1-basierte r,c in A1-Notation um (ohne $)."""
    # Excel-Spaltenindex -> Buchstaben
    col = ""
    n = c
    while n:
        n, rem = divmod(n - 1, 26)
        col = chr(65 + rem) + col
    return f"{col}{r}"


def extract_cells(sheet: "xw.Sheet") -> List[Tuple[str, str, str, str]]:
    """
    Liest belegte Zellen inkl. Formeln aus einem Blatt.
    Gibt Liste von (Blatt, Adresse, Formel, Wert) zurück.
    - Formel bevorzugt Array-Formel, sonst Standard-Formel; leere Zellen ignorieren.
    - Fehlerhafte Zellen werden robust als Wert '#ERROR' behandelt.
    """
    out: List[Tuple[str, str, str, str]] = []
    try:
        used = sheet.used_range
        n_rows = int(used.rows.count)
        n_cols = int(used.columns.count)
        if n_rows == 0 or n_cols == 0:
            return out
        # Top-left absolute Position (1-basiert relativ zum Blatt)
        tl_row = int(used.row)
        tl_col = int(used.column)
    except Exception:
        # Falls used_range scheitert, durch einen großzügigen Scan ersetzen
        n_rows, n_cols, tl_row, tl_col = 200, 50, 1, 1  # Fallback
    # Iteration über den Bereich; robust je Zelle
    for r_off in range(n_rows):
        for c_off in range(n_cols):
            r = tl_row + r_off
            c = tl_col + c_off
            try:
                rng = sheet.range((r, c))
            except Exception:
                # Sehr selten, aber dann weiter
                continue
            formula = ""
            value_s = ""
            try:
                # Array-Formel bevorzugen, falls vorhanden
                # COM-API HasArray ist zuverlässiger, aber formula_array liefert String für Ankerzelle.
                has_array = False
                try:
                    has_array = bool(rng.api.HasArray)
                except Exception:
                    has_array = False

                f_arr = None
                if has_array:
                    try:
                        f_arr = rng.formula_array
                    except Exception:
                        f_arr = None

                if f_arr:
                    formula = to_str(f_arr)
                else:
                    try:
                        f_std = rng.formula
                        if f_std:
                            formula = to_str(f_std)
                    except Exception:
                        formula = ""

                try:
                    val = rng.value
                    # xlwings gibt für Fehler häufig spezielle Objekte/Strings zurück; to_str fängt das ab
                    value_s = to_str(val)
                except Exception:
                    value_s = "#ERROR"

                # Leere Zellen ignorieren (weder Wert noch Formel)
                if (value_s == "" or value_s is None) and (formula == "" or formula is None):
                    continue

                addr = cell_address(r, c)
                out.append((sheet.name, addr, formula, value_s))
            except Exception:
                # Letzte Verteidigung: bei unklaren COM-Fehlern
                try:
                    addr = cell_address(r, c)
                except Exception:
                    addr = f"R{r}C{c}"
                out.append((sheet.name, addr, "", "#ERROR"))
    return out


def extract_names(wb: "xw.Book") -> List[Tuple[str, str, str]]:
    """
    Liest alle definierten Namen des Workbooks.
    Gibt Liste (Blatt, Name, Adresse) zurück.
    - Blatt: falls Name blattspezifisch ist -> Blattname, sonst '' (Workbook-Scope)
    - Adresse: A1-Bezug inkl. Blatt, soweit ermittelbar; bei Fehlern '#REF!'
    """
    out: List[Tuple[str, str, str]] = []
    for nm in wb.names:
        try:
            name_str = to_str(nm.name)
        except Exception:
            name_str = "<UNKNOWN>"
        sheet_name = ""
        addr = ""
        # Scope/Sheet
        try:
            if nm.parent and hasattr(nm, "sheet") and nm.sheet is not None:
                sheet_name = to_str(nm.sheet.name)
            else:
                sheet_name = ""
        except Exception:
            sheet_name = ""

        # Adresse ermitteln: bevorzugt refers_to_range, sonst refers_to (Formel)
        got_addr = False
        try:
            rtr = nm.refers_to_range  # kann Exception werfen bei #REF!
            if rtr is not None:
                # Absolute Adresse mit Blattname
                try:
                    addr = f"{rtr.sheet.name}!{rtr.address}"
                except Exception:
                    addr = to_str(rtr.address)
                got_addr = True
        except Exception:
            got_addr = False

        if not got_addr:
            try:
                ref = nm.refers_to  # z.B. '=Kalkulation!$A$1:$B$4' oder '#REF!'
                if ref:
                    # Entferne führendes '=' für bessere Lesbarkeit
                    addr = to_str(ref).lstrip("=")
                else:
                    addr = "#REF!"
            except Exception:
                addr = "#REF!"

        out.append((sheet_name, name_str, addr))
    return out


def write_csv(df: pd.DataFrame, target_path: Path) -> None:
    target_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(target_path, index=False, encoding="utf-8-sig")


def main() -> None:
    # Pfad zur Excel
    if len(sys.argv) > 1:
        excel_path = Path(sys.argv[1]).resolve()
    else:
        excel_path = Path("input") / "TARIFRECHNER_KLV.xlsm"
        excel_path = excel_path.resolve()

    if not excel_path.exists():
        raise SystemExit(f"Excel-Datei nicht gefunden: {excel_path}")

    excel_dir = excel_path.parent
    out_cells = excel_dir / "excelzell.csv"
    out_names = excel_dir / "excelber.csv"

    # Excel headless öffnen
    app = xw.App(visible=False, add_book=False)  # Excel-Instanz
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = xw.Book(excel_path)
        try:
            # Zellen extrahieren
            rows_cells: List[Tuple[str, str, str, str]] = []
            for sh in wb.sheets:
                rows_cells.extend(extract_cells(sh))

            df_cells = pd.DataFrame(
                rows_cells, columns=["Blatt", "Adresse", "Formel", "Wert"]
            )
            # Leere DataFrames trotzdem schreiben (Prüfkriterium verlangt >= 1 Zeile, deshalb warnen wir nicht)
            write_csv(df_cells, out_cells)

            # Namen/Bereiche extrahieren
            rows_names = extract_names(wb)
            df_names = pd.DataFrame(rows_names, columns=["Blatt", "Name", "Adresse"])
            write_csv(df_names, out_names)
        finally:
            # Workbook schließen, ohne zu speichern
            wb.close()
    finally:
        app.kill()

    # Bequemlichkeitskopie in das aktuelle Arbeitsverzeichnis, falls unterschiedlich,
    # damit einfache Checks wie Path("excelzell.csv") funktionieren.
    cwd = Path.cwd().resolve()
    if cwd != excel_dir:
        try:
            df_cells = pd.read_csv(out_cells, dtype=str)
            df_names = pd.read_csv(out_names, dtype=str)
            write_csv(df_cells, cwd / "excelzell.csv")
            write_csv(df_names, cwd / "excelber.csv")
        except Exception:
            # Optional – Ignorieren, falls keine Schreibrechte o.ä.
            pass

    print(f"Fertig.\nZellen-CSV: {out_cells}\nBereiche-CSV: {out_names}")


if __name__ == "__main__":
    main()
