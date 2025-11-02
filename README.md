# Excel-Tarifrechner → Python-Rechner (LLM Proof of Concept)

*ver. 0.05 (2025-11-28)*

Dieses Repository begleitet das Video / Webinar der DAV-Arbeitsgruppe.

Excel-Tarifrechner sind in der täglichen Aktuarpraxis allgegenwärtig – aber komplexe Formeln, verstreute VBA-Makros und eingeschränkte Teamarbeit bremsen Innovation und Wartbarkeit aus.

Python bietet dank leistungsstarker Bibliotheken eine skalierbare und leicht wartbare Alternative mit klar strukturiertem Code und nahtloser Integration in moderne Workflows. In diesem Video wird gezeigt, wie man unter Einsatz eines Large-Language-Models (LLM) – hier ChatGPT – einen typischen Excel-Tarifrechner nach Python übersetzt. Dazu werden zwei unterschiedliche Ansätze vorgestellt.

**„Portierung von Referenzrechnern mit Large-Language-Models“**  
Ziel ist es, einen klassischen Excel-Tarifrechner der Lebensversicherung reproduzierbar in **reinen Python-Code** zu überführen – in zwei unterschiedlichen Workflows („handwerklich“ vs. „industriell“).

---

## Inhaltsverzeichnis
1. [Projektüberblick](#projektüberblick)  
2. [Repository-Struktur](#repository-struktur)  
3. [Workflows](#workflows)  
   * [Arnos „handwerklicher“ Ansatz](#arnos-handwerklicher-ansatz)  
   * [Barteks „industrieller“ Ansatz](#barteks-industrieller-ansatz)  
4. [Erste Schritte](#erste-schritte)  
5. [Benutzung](#benutzung)  
6. [Tests & Berichte](#tests--berichte)  
7. [Mitwirken](#mitwirken)  
8. [Lizenz](#lizenz)

---

## Projektüberblick

* **Problem:** Excel-Tarifrechner sind schnell gebaut, aber schwer wartbar und kaum CI-fähig.  
* **Lösung:** Einsatz von Large-Language-Models, um Excel-Formeln, VBA-Module und Tabellendaten automatisiert in Python-Code zu migrieren.  
* **Mehrwert:**  
  * nachvollziehbarer, modularer Source-Code  
  * automatisierte Tests & Continuous Integration  
  * Basis für künftige Produkt- und Bestandsmigrationen innerhalb der LV-IT

Die beiden Ansätze unterscheiden sich in **Automatisierungsgrad** und **Tool-Stack**:

| Merkmal                  | Handwerklich (Arno) | Industriell (Bartek) |
|--------------------------|---------------------|----------------------|
| Input für LLM            | VBA-Quelltext & Screenshot | Vollständiger Excel-Dump als Text |
| Manuelle Schritte        | Screenshot, Copy-&-Paste der Formeln | keine |
| Zielsetzung              | schneller PoC       | vollautomatisierbarer Workflow |
| Haupterkenntnis          | LLM erkennt Zellen überraschend gut | Kontext-Limit aktuell Engpass |

---

## Repository-Struktur

```text
dev/
├─ Arno/                 # Handwerklicher Workflow
│  ├─ input/             # Chat-Verlauf (nur Prompts), Screenshot, Original-Excel
│  └─ output/            # Von ChatGPT generierter Python-Code
├─ Bartek/               # Industrieller Workflow
│  ├─ input/             # Optimierte Prompts, Original-Excel
│  └─ output/            # root: i/o und python-Module
│     └─ tests/          # PyTest-Fixtures & Smoke-Tests
└─ README.md             # *this file*
```

*(Bei neuen Files bitte die gleiche Tiefenstruktur beibehalten.)*

---

## Workflows

### Arnos „handwerklicher“ Ansatz

*Ziel:* **Rapid Prototyping** – bei möglichst wenig Prompts und Nutzung eines Reasoning-Modells.

Idee: Da das Modell keine Excel-Datei verarbeiten kann, werden die Bestandteile der Eingabedatei `Tarifrechner_KLV.xlsm` separat behandelt. Die Aufgabe wird in drei Schritte (plus einen 4. Schritt für einen Werteabgleich) zerlegt.

| Schritt | Beschreibung | Chatprotokoll | Erzeugte Dateien | 
| ------- | ------------ | ------------- | ---------------- |
| 1       | Tafeln aus Excel in eine XML-Datei überführen (ganzer Inhalt des Blattes `Tafeln` per Copy&Paste an ChatGPT). | Chat 1 - Excel_nach_XML_konvertieren | `Tafeln.xml` |
| 2       | VBA-Module (`mConstants`, `mBarwerte`, `mGwerte`) nach Python übersetzen; Excel-Rundungsregeln beibehalten. | Chat 2 - VBA_nach_Python_übersetzen | `constants.py` `barwerte.py` `gwerte.py` |
| 3       | Tabellenblatt `Kalkulation` als CLI-Programm abbilden (Screenshot + Formeln als Text). | Chat 3 - Excel-Tarifrechner_nach_Python_mit_QS (Prompts 5–7) | `beitrag_und_verlaufswerte.py` `tarifrechner.py` |
| 4       | Wertevergleich Excel ↔ Python. |  Chat 3 - Excel-Tarifrechner_nach_Python_mit_QS (Prompt 8) | `vergleich.py` |

---

### Barteks „industrieller“ Ansatz

*Workflow-Ziel:* **100 % script-gesteuerte Migration** – keine händischen Zwischenschritte.

``` mermaid
flowchart TD

subgraph Excel-Dump
    A1[excel_to_text.py<br>Zellen & Bereiche → CSV]
    A2[vba_to_text.py<br>VBA-Module → TXT]
    A3[data_extract.py<br>var.csv • tarif.csv • tafeln.csv]
    B1[tests PyTest<br>Smoke + Funktionsparitaet]
    A1 --> B1
    A2 --> B1
    A3 --> B1
end

subgraph Code-Portierung
    C1[basfunct.py<br>VBA-Basis → Python]
    D1[Bxt in ausfunct.py]
    D2[BJB • BZB • Pxt<br>offen]
    D3[verlaufswerte<br>offen]
    R1[tests PyTest<br>Referenzparitaet]
    C1 --> D1
    C1 --> D2
    C1 --> D3
    D1 --> R1
    D2 --> R1
    D3 --> R1
end

subgraph Ausführungsebene
    E1[run_calc.py<br>CLI-Runner]
end

B1 --> C1
R1 --> E1

```

*Workflow-Phasen:*

| Abschnitt | Bedeutung |
|-----------|-----------|
| **Excel-Dump & Preprocessing** | Automatisierte Extract-Skripte (Schritte 1–4) |
| **Code-Portierung** | Übersetzung der Logik nach Python (Schritte 5–6C) |
| **Ausführungsebene** | End-User-Interface via CLI (Schritt 7) |

*Status:* ✅ erledigt – Schritte 1–5, 6A, 7 • ⏳ offen – Schritte 6B & 6C

---

## Erste Schritte

### Variante B – GitHub Codespaces (Browser-IDE)

1. Repository öffnen → **Branch `docker-seminar-setup`** wählen → **Code ▸ Codespaces ▸ Create**.  
2. Nach dem Start öffnet sich die VS-Code-Web-IDE.  
3. Terminal öffnen und z. B. ausführen:
   ```bash
   pytest -q
   python Bartek/output/run_calc.py --help
   ```

### Variante C – Lokal mit Docker Desktop + VS Code

**Voraussetzungen**  
- Docker Desktop installiert (Windows: **Linux-Container** aktiv – wenn im Menü „Switch to Windows containers…“ steht, ist alles korrekt).  
- Visual Studio Code + Extension **Dev Containers** (`ms-vscode-remote.remote-containers`).  
- Git CLI.

**Setup (Windows-Beispiel)**

```powershell
# Projektordner anlegen & betreten
cd C:\dev\LLM_seminar

# Repo klonen (Seminar-Branch), öffnen
git clone -b docker-seminar-setup --single-branch https://github.com/bartlmac/portxlpy.git
cd portxlpy
code .
```

**In VS Code:** `F1` → **Dev Containers: Reopen in Container**  
**Smoke-Test (im Container):**
```powershell
pytest -q    # Erwartet: 4 passed
```

> **Hinweis:** `postCreateCommand` läuft nur beim **Neuaufbau** – bei Bedarf `F1` → **Dev Containers: Rebuild and Reopen in Container**.

**Troubleshooting (lokal):**
```powershell
# Projekt sauber stoppen und Volumes löschen
docker compose down -v

# Optional mehr Platz schaffen – unbenutzte Images entfernen (Vorsicht!)
docker image prune -a
```

> **Warnung:** `docker image prune -a` löscht **alle unbenutzten** Images. Verwende nur, wenn du sicher bist, dass diese Images nicht mehr gebraucht werden.

---

## Benutzung

### CLI-Runner von Arno
```bash
# Hauptberechnung
python Arno/output/tarifrechner.py

# Werte-Gegenprobe Excel ↔ Python (optional über pytest)
python Arno/output/vergleich.py
```

### CLI-Runner von Bartek
```bash
# Funktionsweise wählbar mit --funcs
python Bartek/output/run_calc.py --funcs Bxt
```

---

## Tests & Berichte

### Lokal ausführen
```bash
# Arno
cd Arno/output && pytest -q

# Bartek
cd Bartek/output && pytest -q
```
*Terminal bleibt dank `-q` aufgeräumt (nur „passed/failed“).*

Optional: JUnit-XML/HTML-Report erzeugen (siehe `pytest`-Plugins).

---

## Mitwirken

Pull Requests sind willkommen! Bitte beachte:

1. Erstelle einen Issue für größere Änderungen.  
2. Schreibe (oder aktualisiere) Tests für neue Features.  

---

## Lizenz
*DAV*

---

**Kontakt:**  
*Bartlomiej Maciaga* – <bartlomiej.maciaga@hotmail.com>  
*Dr. Arno Rasch* – <arno.rasch@vtmw.de>
