# Klassenplanung

Erzeugt aus Untis-Stundenplan-PDFs pro Klasse eine Planungs-Excel-Datei
(nach Kalenderwoche ausgerichtete Lektionszeilen, Ferien/Feiertage,
Spezialwochen und stunden-genaue Ausfälle durch Spezialtage).

## Pipeline

```
input_files/*.pdf                 (Untis-Stundenplan-Export als PDF)
        │  pdf_timetable_to_xlsx.py
        ▼
output_files/Stundenplan_*.xlsx   (1:1-Excel-Abbild der PDF-Tabellen)
        │  stundenplan_parsing.py + rendering.py
        ▼
output_files/klassen/Planung_*.xlsx   (1 Datei pro Klasse)
```

`main.py` führt beide Schritte nacheinander aus. Die drei Skripte können
aber auch einzeln aufgerufen werden (siehe unten).

## Voraussetzungen

- Python 3.10+
- Abhängigkeiten installieren:
  ```bash
  pip install -r requirements.txt
  ```

## Verwendung

**Komplette Pipeline** (PDFs konvertieren + Klassen-Planungen erzeugen):

```bash
python3 main.py
```

**Nur PDFs konvertieren** (einzelnes PDF, z. B. zum Testen):

```bash
python3 pdf_timetable_to_xlsx.py input_files/irgendein.pdf output_files/Stundenplan_XY_Phasen_1_3.xlsx
```

**Nur Klassen-Planungen neu erzeugen** (wenn sich nur `config.py` geändert
hat und die Stundenplan-Exporte in `output_files/` noch aktuell sind):

```bash
python3 klassenplanung.py
```

## Eingabedateien: Namensschema

Alle PDFs in `input_files/` müssen dem Schema

```
..._<Kürzel>_phasen_<von>_<bis>.pdf
```

folgen, z. B. `leh a4q_KÜ_phasen_1_3.pdf`. Daraus leitet `main.py` den
Namen des Exports ab: `Stundenplan_KÜ_Phasen_1_3.xlsx`. Passt ein
Dateiname nicht ins Schema, bricht `main.py` mit einer klaren
Fehlermeldung ab, statt still einen falschen Namen zu vergeben.

## Projektstruktur

| Datei                     | Zweck                                                                 |
|----------------------------|------------------------------------------------------------------------|
| `main.py`                  | Einstiegspunkt: führt die komplette Pipeline aus                      |
| `pdf_timetable_to_xlsx.py` | Wandelt ein Stundenplan-PDF 1:1 in ein Excel-Raster um                |
| `stundenplan_parsing.py`   | Liest die Excel-Raster aus und extrahiert Lektionen je Klasse/Fach    |
| `rendering.py`             | Baut das Excel-Layout der Planungsdateien (Kopfzeile, Wochen, Blöcke) |
| `klassenplanung.py`        | Orchestriert: prüft pro Klasse, ob der Stundenplan übers Jahr konstant bleibt, und erzeugt die Dateien |
| `config.py`                | Schulkalender + Gestaltungswerte (siehe unten)                        |

## Jährliche Wartung: `config.py`

Vor jedem neuen Schuljahr in `config.py` aktualisieren:

- `BASIS_JAHR` – Kalenderjahr, in dem Phase 1 beginnt
- `FERIEN` – Ferienblöcke (Start-/Enddatum)
- `UNTERRICHTSFREIETAGE` – einzelne unterrichtsfreie Tage
- `SPEZIALTAGE` – Termine mit Uhrzeiten (z. B. pädagogische Halbtage), die
  nur einzelne Lektionen ausfallen lassen, nicht den ganzen Tag
- `SPEZIALWOCHEN` – Klassenstufen-spezifische Wochen (Skilager, Projektwochen, ...)

`STIL` (Schrift, Farben, Zeilen-/Spaltenmasse) muss nur angepasst werden,
wenn sich am Aussehen der Excel-Dateien etwas ändern soll.

## Ausgabe

- `output_files/Stundenplan_*.xlsx` – ein Export pro PDF (Zwischenschritt)
- `output_files/klassen/Planung_<Klasse>.xlsx` – die eigentlichen
  Planungsdateien, ein Arbeitsblatt pro Jahr oder pro Phase (je nachdem,
  ob sich der Stundenplan der Klasse innerhalb des Jahres ändert)
