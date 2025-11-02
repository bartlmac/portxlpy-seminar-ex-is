# 🧭 PROMPT – Einstiegspunkt für Folge-LLM

> **Kontext:**
> Du steigst in ein laufendes Projekt ein. Ziel des Projekts ist, aus einer bestehenden Excel-Arbeitsmappe (`TARIFRECHNER_KLV.xlsm`) einen modularen, reinen Python-Produktrechner zu erstellen, der identische Ergebnisse liefert.
> Der bisherige Projektverlauf wurde bereits erfolgreich abgeschlossen bis einschließlich der Basisfunktionen.
> Du sollst **ab hier direkt weiterarbeiten**, nicht nachfragen oder rekonstruieren.

## Aktueller Zustand / Verfügbare Artefakte

Du arbeitest komplett im unterverzeichnis "Bartek" des Repo-Roots.

Folgende Dateien und Daten liegen vollständig vor und bilden den aktuellen Projektstand, sie werden in drei Schritten hochgeladen, danach, im vierten Schritt, bekommst du das Prompt mit der eigenlichen Aufgabe:

### 🧩 **Input für das LLM (Kontextquellen)**

Diese Dateien dienen ausschließlich als **Inhalts- und Wissensbasis** für die Ableitung der Python-Logik aus der Excel-Struktur.
Der Python-Code selbst soll sie **nicht direkt referenzieren**, aber das LLM darf sie verwenden, um Formeln, Abhängigkeiten und Berechnungswege zu verstehen.

```
protokoll.txt   – Vollständiger Projektverlauf bis unmittelbar vor TASK 6A (inkl. Entscheidungen & Code)
excelzell.csv   – Vollständiger Dump aller belegten Excel-Zellen inkl. Formeln
excelber.csv    – Übersicht aller benannten Bereiche aus der Excel-Datei
```

### ✅ **Bereits implementierte Artefakte des neuen Python-Rechners**

Diese Dateien sind funktional, getestet und bilden die technische Basis:

```
excel_to_text.py   – Extraktion der Excel-Zellen und Bereiche
vba_to_text.py     – Export aller VBA-Module
data_extract.py    – Generiert var.csv, tarif.csv, grenzen.csv, tafeln.csv, tarif.py
basfunct.py        – Vollständiger 1:1-Port der VBA-Basisfunktionen (mGWerte, mBarwerte, mConstants)
tarif.py           – Enthält raten_zuschlag(zw)
tests/             – pytest-Struktur vorhanden
```

### 📊 **Datenartefakte (für Berechnungen relevant)**

Diese Dateien stellen die Eingangsparameter und Tabellen des Rechners dar:

```
var.csv       – Vertragsvariablen (x, n, t, VS, zw, Sex)
tarif.csv     – Tarifparameter (Zins, Tafel, alpha, beta1, gamma1, gamma2, gamma3, k)
grenzen.csv   – Grenzwerte (MinAlterFlex, MinRLZFlex)
tafeln.csv    – Sterbetafel (Long-Format, Spalten Name|Wert)
```

## Technischer Rahmen

* Umgebung: Windows / VS Code / Bash-Terminal
* Sprache: Python 3.11+
* Qualitätsanforderung:
  * pytest = grün
* Jede Aufgabe liefert **einen einzigen ausführbaren Python-Codeblock**
* Kein Fließtext, keine Erklärungen, keine Diskussion

## Vorgehen

Ich lade jetzt die drei Prompts mit den Kontext-Dateien und Artefakten hoch. Bitte davor keinen Python-Code generieren sondern auf mein "Go" warten. Danach gebe ich das vierte Prompt mit der Aufgabe und "Go", dann kannst du deinen Code-Vorschlag generieren.
