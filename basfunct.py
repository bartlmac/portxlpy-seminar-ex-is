# basfunct.py
# -*- coding: utf-8 -*-
"""
1:1-Port der VBA-Basisfunktionen aus den Modulen:
- mGWerte  :contentReference[oaicite:0]{index=0}
- mBarwerte  :contentReference[oaicite:1]{index=1}
- mConstants  :contentReference[oaicite:2]{index=2}

Vorgaben:
- Pandas wird für Tabellen-/CSV-Zugriffe genutzt (insb. tafeln.csv).
- Kein Funktionsrumpf endet mit 'pass'.
- Datenquellen (bei Bedarf): excelzell.csv, excelber.csv, var.csv, tarif.csv, grenzen.csv, tafeln.csv.

Hinweise:
- Excel-Rundungen (WorksheetFunction.Round) werden mit Decimal(ROUND_HALF_UP) nachgebildet.
- Der VBA-Cache (Scripting.Dictionary) wird als Python- dict implementiert.
- Act_qx liest aus 'tafeln.csv' (Long-Format "Name|Wert", wobei Name = "<Spaltenüberschrift>|<Zeilen-Schlüssel aus Spalte A>").
  Erwartete Spaltenüberschriften sind z. B. "DAV1994_T_M", "DAV2008_T_F" analog zum VBA-Code.  :contentReference[oaicite:3]{index=3}
"""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP, InvalidOperation, getcontext
from typing import Any, Dict, Tuple, Optional

import pandas as pd

# ------------------------------------------------------------
# Konstanten (aus mConstants)  :contentReference[oaicite:4]{index=4}
# ------------------------------------------------------------
rund_lx: int = 16
rund_tx: int = 16
rund_Dx: int = 16
rund_Cx: int = 16
rund_Nx: int = 16
rund_Mx: int = 16
rund_Rx: int = 16
max_Alter: int = 123


# ------------------------------------------------------------
# Hilfen: Excel-kompatibles Runden & Zahl-Conversion
# ------------------------------------------------------------
getcontext().prec = 40  # reichlich Präzision für finanzmathematische Zwischenschritte


def _xl_round(x: float | int, ndigits: int) -> float:
    """
    Excel ROUND-Nachbildung (kaufmännisch: .5 -> vom Null weg).
    VBA: WorksheetFunction.Round(...).  :contentReference[oaicite:5]{index=5}
    """
    try:
        q = Decimal(str(x))
        if ndigits >= 0:
            exp = Decimal("1").scaleb(-ndigits)  # = 10**(-ndigits)
        else:
            # negative Stellen: runden auf 10er/100er etc.
            exp = Decimal("1").scaleb(-ndigits)  # funktioniert auch für negative
        return float(q.quantize(exp, rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError, TypeError):
        # Fallback: Python round (nur als Notnagel)
        return float(round(x, ndigits))


def _to_float(val: Any) -> float:
    try:
        return float(val)
    except Exception:
        return float("nan")


# ------------------------------------------------------------
# Datenlade-Logik für Tafeln (ersetzt Worksheet.Range-Lookups)  :contentReference[oaicite:6]{index=6}
# ------------------------------------------------------------
class _TafelnRepo:
    """
    Lädt 'tafeln.csv' (Spalten: Name, Wert) und stellt Lookup nach Spalten-Header & Zeilenschlüssel bereit.
    Erwartetes 'Name'-Schema: "<Header>|<Key>", z. B. "DAV2008_T_M|65".
    """

    def __init__(self) -> None:
        self._loaded = False
        self._map: Dict[Tuple[str, str], float] = {}

    def _load(self) -> None:
        if self._loaded:
            return
        # Suchreihenfolge: CWD, ./input
        candidates = [pd.Path.cwd() / "tafeln.csv"] if hasattr(pd, "Path") else []
        # Fallback: pathlib
        from pathlib import Path

        candidates = [Path("tafeln.csv"), Path("input") / "tafeln.csv"]
        for p in candidates:
            if p.exists():
                df = pd.read_csv(p, dtype={"Name": str})
                for _, row in df.iterrows():
                    name = str(row.get("Name", "")).strip()
                    if not name or "|" not in name:
                        continue
                    header, key = name.split("|", 1)
                    val = _to_float(row.get("Wert", ""))
                    self._map[(header.strip().upper(), key.strip())] = val
                self._loaded = True
                return
        # Wenn nichts gefunden wurde, als leer markieren; Act_qx wird dann Fehler werfen.
        self._loaded = True
        self._map = {}

    def qx(self, header: str, age_key: str) -> float:
        self._load()
        k = (header.strip().upper(), age_key.strip())
        if k not in self._map:
            raise KeyError(
                f"Sterbewert nicht gefunden für Header='{header}', Key='{age_key}'. "
                f"Stelle sicher, dass 'tafeln.csv' mit passendem Schema vorhanden ist."
            )
        return self._map[k]


# Singleton-Repo
_tafeln_repo = _TafelnRepo()

# ------------------------------------------------------------
# Cache-Mechanismus (aus mGWerte)  :contentReference[oaicite:7]{index=7}
# ------------------------------------------------------------
_cache: Dict[str, float] | None = None


def InitializeCache() -> None:
    """Erstellt/initialisiert den Cache (entspricht CreateObject('Scripting.Dictionary')).  :contentReference[oaicite:8]{index=8}"""
    global _cache
    _cache = {}  # leerer dict


def CreateCacheKey(
    Art: str,
    Alter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int,
    Rentenbeginnalter: int,
    Schicht: int,
) -> str:
    """Bildet den Schlüssel wie im VBA-Original.  :contentReference[oaicite:9]{index=9}"""
    return f"{Art}_{Alter}_{Sex}_{Tafel}_{Zins}_{GebJahr}_{Rentenbeginnalter}_{Schicht}"


def _ensure_cache() -> Dict[str, float]:
    global _cache
    if _cache is None:
        InitializeCache()
    assert _cache is not None
    return _cache


# ------------------------------------------------------------
# Grundwerte/Tabellenfunktionen (aus mGWerte)  :contentReference[oaicite:10]{index=10}
# ------------------------------------------------------------
def Act_qx(
    Alter: int,
    Sex: str,
    Tafel: str,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """
    Liefert q_x aus 'Tafeln' in Abhängigkeit von Alter, Geschlecht und Tafelkürzel.
    VBA-Logik:
      - Sex != "M" -> "F"
      - unterstützte Tafeln: "DAV1994_T", "DAV2008_T" (sonst ERROR 1)  :contentReference[oaicite:11]{index=11}
      - Spaltenwahl via sTafelvektor = f"{Tafel}_{Sex}"
      - Indexierung Alter+1 in Matrize; hier nutzen wir den Schlüssel (Alter) aus Spalte A in tafeln.csv.
    """
    sex = Sex.upper()
    if sex != "M":
        sex = "F"
    tafel_u = Tafel.upper()
    if tafel_u not in {"DAV1994_T", "DAV2008_T"}:
        # VBA: Act_qx = 1 : Error(1) – wir lösen in Python eine Exception aus:
        raise ValueError(f"Nicht unterstützte Tafel: {Tafel}")

    sTafelvektor = f"{tafel_u}_{sex}"
    # In tafeln.csv wird die erste Spalte (Keys) als Bestandteil des 'Name'-Feldes kodiert.
    # Wir erwarten, dass der Schlüssel dem Alter entspricht (z. B. "65").
    age_key = str(int(Alter))  # robust gegen floats
    return float(_tafeln_repo.qx(sTafelvektor, age_key))


def v_lx(
    Endalter: int,
    Sex: str,
    Tafel: str,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> list[float]:
    """
    Vektor der lx; Startwert 1_000_000; Rundung je Schritt mit rund_lx.  :contentReference[oaicite:12]{index=12}
    Endalter = -1 -> bis max_Alter.
    """
    grenze = max_Alter if Endalter == -1 else Endalter
    vek = [0.0] * (grenze + 1)
    vek[0] = 1_000_000.0
    for i in range(1, grenze + 1):
        q_prev = Act_qx(i - 1, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
        vek[i] = _xl_round(vek[i - 1] * (1.0 - q_prev), rund_lx)
    return vek


def Act_lx(
    Alter: int,
    Sex: str,
    Tafel: str,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """lx an Position Alter.  :contentReference[oaicite:13]{index=13}"""
    vek = v_lx(Alter, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    return float(vek[Alter])


def v_tx(
    Endalter: int,
    Sex: str,
    Tafel: str,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> list[float]:
    """Vektor der tx (#Tote), Rundung rund_tx.  :contentReference[oaicite:14]{index=14}"""
    grenze = max_Alter if Endalter == -1 else Endalter
    vek = [0.0] * (grenze + 1)
    v_temp_lx = v_lx(grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    for i in range(0, grenze):
        vek[i] = _xl_round(v_temp_lx[i] - v_temp_lx[i + 1], rund_tx)
    return vek


def Act_tx(
    Alter: int,
    Sex: str,
    Tafel: str,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """tx an Position Alter.  :contentReference[oaicite:15]{index=15}"""
    vek = v_tx(Alter, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    return float(vek[Alter])


def v_Dx(
    Endalter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> list[float]:
    """Vektor der Dx; Rundung rund_Dx.  :contentReference[oaicite:16]{index=16}"""
    grenze = max_Alter if Endalter == -1 else Endalter
    vek = [0.0] * (grenze + 1)
    v_ = 1.0 / (1.0 + float(Zins))
    v_temp_lx = v_lx(grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    for i in range(0, grenze + 1):
        vek[i] = _xl_round(v_temp_lx[i] * (v_ ** i), rund_Dx)
    return vek


def Act_Dx(
    Alter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """Dx(Alter) mit Cache.  :contentReference[oaicite:17]{index=17}"""
    cache = _ensure_cache()
    key = CreateCacheKey("Dx", Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    if key in cache:
        return cache[key]
    vek = v_Dx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    res = float(vek[Alter])
    cache[key] = res
    return res


def v_Cx(
    Endalter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> list[float]:
    """Vektor der Cx; Rundung rund_Cx.  :contentReference[oaicite:18]{index=18}"""
    grenze = max_Alter if Endalter == -1 else Endalter
    vek = [0.0] * (grenze + 1)
    v_ = 1.0 / (1.0 + float(Zins))
    v_temp_tx = v_tx(grenze, Sex, Tafel, GebJahr, Rentenbeginnalter, Schicht)
    for i in range(0, grenze):
        vek[i] = _xl_round(v_temp_tx[i] * (v_ ** (i + 1)), rund_Cx)
    return vek


def Act_Cx(
    Alter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """Cx(Alter) mit Cache.  :contentReference[oaicite:19]{index=19}"""
    cache = _ensure_cache()
    key = CreateCacheKey("Cx", Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    if key in cache:
        return cache[key]
    vek = v_Cx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    res = float(vek[Alter])
    cache[key] = res
    return res


def v_Nx(
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> list[float]:
    """Vektor der Nx; rückwärts kumulierte Summe der Dx; Rundung rund_Dx.  :contentReference[oaicite:20]{index=20}"""
    vek = [0.0] * (max_Alter + 1)
    v_temp_Dx = v_Dx(-1, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    vek[max_Alter] = v_temp_Dx[max_Alter]
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = _xl_round(vek[i + 1] + v_temp_Dx[i], rund_Dx)
    return vek


def Act_Nx(
    Alter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """Nx(Alter) mit Cache.  :contentReference[oaicite:21]{index=21}"""
    cache = _ensure_cache()
    key = CreateCacheKey("Nx", Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    if key in cache:
        return cache[key]
    vek = v_Nx(Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    res = float(vek[Alter])
    cache[key] = res
    return res


def v_Mx(
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> list[float]:
    """Vektor der Mx; rückwärts kumulierte Summe der Cx; Rundung rund_Mx.  :contentReference[oaicite:22]{index=22}"""
    vek = [0.0] * (max_Alter + 1)
    v_temp_Cx = v_Cx(-1, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    vek[max_Alter] = v_temp_Cx[max_Alter]
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = _xl_round(vek[i + 1] + v_temp_Cx[i], rund_Mx)
    return vek


def Act_Mx(
    Alter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """Mx(Alter) mit Cache.  :contentReference[oaicite:23]{index=23}"""
    cache = _ensure_cache()
    key = CreateCacheKey("Mx", Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    if key in cache:
        return cache[key]
    vek = v_Mx(Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    res = float(vek[Alter])
    cache[key] = res
    return res


def v_Rx(
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> list[float]:
    """Vektor der Rx; rückwärts kumulierte Summe der Mx; Rundung rund_Rx.  :contentReference[oaicite:24]{index=24}"""
    vek = [0.0] * (max_Alter + 1)
    v_temp_Mx = v_Mx(Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    vek[max_Alter] = v_temp_Mx[max_Alter]
    for i in range(max_Alter - 1, -1, -1):
        vek[i] = _xl_round(vek[i + 1] + v_temp_Mx[i], rund_Rx)
    return vek


def Act_Rx(
    Alter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """Rx(Alter) mit Cache.  :contentReference[oaicite:25]{index=25}"""
    cache = _ensure_cache()
    key = CreateCacheKey("Rx", Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    if key in cache:
        return cache[key]
    vek = v_Rx(Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    res = float(vek[Alter])
    cache[key] = res
    return res


def Act_Altersberechnung(GebDat: pd.Timestamp | str, BerDat: pd.Timestamp | str, Methode: str) -> int:
    """
    Altersberechnung nach Kalenderjahresmethode ('K') bzw. Halbjahresmethode ('H').  :contentReference[oaicite:26]{index=26}
    """
    # Normalisieren auf pandas.Timestamp
    gd = pd.to_datetime(GebDat)
    bd = pd.to_datetime(BerDat)
    meth = "H" if Methode != "K" else "K"

    J_GD = gd.year
    J_BD = bd.year
    M_GD = gd.month
    M_BD = bd.month

    if meth == "K":
        return int(J_BD - J_GD)
    else:
        # Int(J_BD - J_GD + 1/12 * (M_BD - M_GD + 5))
        return int((J_BD - J_GD) + (1.0 / 12.0) * (M_BD - M_GD + 5))


# ------------------------------------------------------------
# Barwerte (aus mBarwerte)  :contentReference[oaicite:27]{index=27}
# ------------------------------------------------------------
def Act_Abzugsglied(k: int, Zins: float) -> float:
    """
    Abzugsglied gemäß VBA-Schleife.  :contentReference[oaicite:28]{index=28}
    """
    if k <= 0:
        return 0.0
    acc = 0.0
    for l in range(0, k):
        acc += (l / k) / (1.0 + (l / k) * float(Zins))
    return acc * (1.0 + float(Zins)) / k


def Act_ag_k(g: int, Zins: float, k: int) -> float:
    """Barwert einer vorschüssigen Rentenzahlung mit k-Zahlungen p.a. über g Perioden.  :contentReference[oaicite:29]{index=29}"""
    v = 1.0 / (1.0 + float(Zins))
    if k <= 0:
        return 0.0
    if Zins > 0:
        # (1 - v^g) / (1 - v) - Abzugsglied * (1 - v^g)
        return (1.0 - (v ** g)) / (1.0 - v) - Act_Abzugsglied(k, float(Zins)) * (1.0 - (v ** g))
    else:
        return float(g)


def Act_ax_k(
    Alter: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    k: int,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """
    äx_k = Nx/Dx - Abzugsglied(k,Zins); nur falls k>0, sonst 0.  :contentReference[oaicite:30]{index=30}
    """
    if k <= 0:
        return 0.0
    return Act_Nx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht) / Act_Dx(
        Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht
    ) - Act_Abzugsglied(k, float(Zins))


def Act_axn_k(
    Alter: int,
    n: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    k: int,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """
    ax:n_k gemäß VBA.  :contentReference[oaicite:31]{index=31}
    """
    if k <= 0:
        return 0.0
    part1 = (
        Act_Nx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
        - Act_Nx(Alter + n, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    ) / Act_Dx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    part2 = Act_Abzugsglied(k, float(Zins)) * (
        1.0
        - Act_Dx(Alter + n, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
        / Act_Dx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    )
    return part1 - part2


def Act_nax_k(
    Alter: int,
    n: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    k: int,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """
    n|ax_k gemäß VBA.  :contentReference[oaicite:32]{index=32}
    """
    if k <= 0:
        return 0.0
    return (
        Act_Dx(Alter + n, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
        / Act_Dx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
        * Act_ax_k(Alter + n, Sex, Tafel, float(Zins), k, GebJahr, Rentenbeginnalter, Schicht)
    )


def Act_nGrAx(
    Alter: int,
    n: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """
    n-Graduationswert Ax gemäß VBA: (Mx(x) - Mx(x+n)) / Dx(x).  :contentReference[oaicite:33]{index=33}
    """
    return (
        Act_Mx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
        - Act_Mx(Alter + n, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)
    ) / Act_Dx(Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht)


def Act_nGrEx(
    Alter: int,
    n: int,
    Sex: str,
    Tafel: str,
    Zins: float,
    GebJahr: int = 0,
    Rentenbeginnalter: int = 0,
    Schicht: int = 1,
) -> float:
    """
    n-Graduationswert Ex gemäß VBA: Dx(x+n) / Dx(x).  :contentReference[oaicite:34]{index=34}
    """
    return Act_Dx(Alter + n, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht) / Act_Dx(
        Alter, Sex, Tafel, float(Zins), GebJahr, Rentenbeginnalter, Schicht
    )
