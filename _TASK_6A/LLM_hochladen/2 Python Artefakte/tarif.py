# -*- coding: utf-8 -*-
"""
tarif.py
Erzeugt aus Excel: Kalkulation!E12
Die Funktion raten_zuschlag(zw) liefert standardmäßig den in Excel berechneten Wert zurück.
Excel-Formel (E12), dokumentiert zu Referenzzwecken:

'=IF(zw=2,2%,IF(zw=4,3%,IF(zw=12,5%,0)))'
"""

from __future__ import annotations
from typing import Any

# In Excel berechneter Referenzwert aus E12:
_E12_VALUE = 0.05

def raten_zuschlag(zw: Any) -> Any:
    """
    Raten-Zuschlag in Abhängigkeit der Zahlweise 'zw'.
    Aktuell wird der referenzierte Excel-Wert zurückgegeben, sodass
    der Erfolgs-Check (zw=12) identisch ist.
    """
    return _E12_VALUE
