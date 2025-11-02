## TASK 6A – Bxt()  (Beitrags­berechnung 1 / 4)

1. **Input**  
   - basfunct.py  
   - CSVs: var.csv, tarif.csv, grenzen.csv, tafeln.csv  
   - excelzell.csv & excelber.csv (für Zell-/Namens­referenzen)  
   - Alle `*.csv`-Dateien dem LLM hochladen

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
