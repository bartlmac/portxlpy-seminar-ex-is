# vba_to_text.py
# -*- coding: utf-8 -*-
"""
Exportiert alle VBA-Module aus einer Excel-Arbeitsmappe (xlsm) in einzelne Textdateien.

Input:
    input/TARIFRECHNER_KLV.xlsm  (oder per CLI-Argument)

Output:
    Für jedes nicht-leere VBA-Modul eine Datei nach Schema:
        Mod_<Name>.txt
    Ablage im selben Verzeichnis wie die Excel-Datei.

Vorgehen:
    - Nutzung von oletools.olevba zum Extrahieren der Module.
    - Alle nichtleeren Module werden exportiert (auch solche ohne Prozeduren, z. B. nur Konstanten).
    - Leere Module/Code-Objekte werden ignoriert.
    - Robuste Dateinamen (sanitisiert), Duplikate werden durchnummeriert.

Aufruf:
    python vba_to_text.py [optional: path\\to\\TARIFRECHNER_KLV.xlsm]
"""

from __future__ import annotations
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

try:
    from oletools.olevba import VBA_Parser  # type: ignore
except Exception as e:
    raise SystemExit(
        "Fehler: 'oletools' ist nicht installiert oder inkompatibel.\n"
        "Bitte installieren mit: pip install oletools\n"
        f"Originalfehler: {e}"
    )

PROCRE = re.compile(r"\b(Sub|Function)\b", flags=re.IGNORECASE)


def sanitize_module_name(name: str) -> str:
    name = name.strip() or "Unbenannt"
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)
    if name.upper() in {
        "CON",
        "PRN",
        "AUX",
        "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10)),
    }:
        name = f"_{name}_"
    return name


def is_nonempty(code: str) -> bool:
    if not code:
        return False
    stripped = []
    for ln in code.splitlines():
        s = ln.strip()
        if not s or s.startswith("'"):
            continue
        if "'" in s:
            s = s.split("'", 1)[0].strip()
        if s:
            stripped.append(s)
    return bool(stripped)


def collect_modules(xlsm_path: Path) -> Dict[str, List[str]]:
    """Liest alle VBA-Module aus der Arbeitsmappe."""
    modules: Dict[str, List[str]] = defaultdict(list)
    vp = VBA_Parser(str(xlsm_path))
    try:
        if not vp.detect_vba_macros():
            return {}
        for (_subfilename, _stream_path, vba_filename, vba_code) in vp.extract_all_macros():
            try:
                mod_name = sanitize_module_name(vba_filename or "Unbenannt")
                if vba_code and is_nonempty(vba_code):
                    modules[mod_name].append(vba_code)
            except Exception:
                continue
    finally:
        try:
            vp.close()
        except Exception:
            pass
    return modules


def write_modules(modules: Dict[str, List[str]], out_dir: Path) -> List[Tuple[str, Path, bool]]:
    """Schreibt Module als Textdateien."""
    results: List[Tuple[str, Path, bool]] = []
    used_names: Dict[str, int] = defaultdict(int)

    for raw_name, chunks in sorted(modules.items()):
        base = f"Mod_{raw_name}.txt"
        cnt = used_names[base]
        used_names[base] += 1
        filename = base if cnt == 0 else f"Mod_{raw_name}_{cnt+1}.txt"

        path = out_dir / filename
        path.parent.mkdir(parents=True, exist_ok=True)

        code = "\n\n' --------- Modul-Teilung ---------\n\n".join(chunks)
        has_proc = bool(PROCRE.search(code))

        if not has_proc:
            header = (
                "' Hinweis: Dieses Modul enthält keine 'Sub' oder 'Function'-Definitionen; "
                "z. B. nur Konstanten/Attribute.\n"
            )
            code = header + code

        path.write_text(code, encoding="utf-8-sig")
        results.append((raw_name, path, has_proc))
    return results


def main() -> None:
    if len(sys.argv) > 1:
        xlsm_path = Path(sys.argv[1]).resolve()
    else:
        xlsm_path = (Path("input") / "TARIFRECHNER_KLV.xlsm").resolve()

    if not xlsm_path.exists():
        raise SystemExit(f"Excel-Datei nicht gefunden: {xlsm_path}")

    out_dir = xlsm_path.parent
    modules = collect_modules(xlsm_path)

    if not modules:
        print("Keine nicht-leeren VBA-Module gefunden.")
        return

    results = write_modules(modules, out_dir)
    total = len(results)
    with_proc = sum(1 for _n, _p, hp in results if hp)
    print(
        f"Export abgeschlossen: {total} Modul-Datei(en) geschrieben in {out_dir}\n"
        f"Mit Prozeduren (Sub/Function): {with_proc} | Ohne Prozeduren: {total - with_proc}"
    )


if __name__ == "__main__":
    main()
