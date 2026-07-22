"""
Zentrale Konfiguration für klassenplanung.py: Schulkalender (Ferien,
unterrichtsfreie Tage, Spezialtage mit Uhrzeiten, Spezialwochen) und die
Excel-Gestaltung (Schrift, Farben, Zeilen-/Spaltenmasse).

Diese Datei enthält nur rohe Werte (Daten, Hex-Farben, Zahlen) - keine
openpyxl-Objekte. Die Umwandlung in Font/PatternFill/etc. passiert in
klassenplanung.py, da das Rendering-Logik ist, keine Konfiguration.

Muss jedes Schuljahr aktualisiert werden: BASIS_JAHR sowie die vier
Kalender-Dictionaries unten.
"""

from datetime import date, datetime

from farbtoene import tint

# Schuljahr, in dem Phase 1 beginnt (die Stundenplan-Exporte enthalten keine
# Jahreszahlen, nur Tag.Monat).
BASIS_JAHR = 2026

FERIEN = {
    "Herbstferien": [date(2026, 9, 26), date(2026, 10, 11)],
    "Weihnachtsferien": [date(2026, 12, 19), date(2027, 1, 3)],
    "Fasnachtsferien": [date(2027, 2, 6), date(2027, 2, 21)],
    "Fruehjahrsferien": [date(2027, 3, 20), date(2027, 4, 4)],
}

UNTERRICHTSFREIETAGE = {
    "GeKo": date(2027, 3, 10),
    "Tag der Arbeit": date(2027, 5, 1),
    "Auffahrt": date(2027, 5, 6),
    "Freitag nach Auffahrt": date(2027, 5, 7),
    "Pfingstmontag": date(2027, 5, 17),
}

SPEZIALTAGE = {
    "Alle Klassen Sporttag": [datetime(2026, 8, 31, 8, 0), datetime(2026, 8, 31, 18, 0)],
    "Alle Klassen 1. Pädagogischer Halbtag": [datetime(2026, 9, 2, 14, 0), datetime(2026, 9, 2, 18, 0)],
    "Alle Klassen 1.5 Pädagogischer Halbtag": [datetime(2026, 10, 13, 14, 0), datetime(2026, 10, 13, 18, 0)],
    "Alle Klassen Tec-Day": [datetime(2026, 10, 15, 8, 0), datetime(2026, 10, 15, 18, 0)],
    "Alle Klassen 2. Pädagogischer Halbtag": [datetime(2026, 10, 23, 14, 0), datetime(2026, 10, 23, 15, 45)],
    "4. Klassen Infotag Uni-Basel": [datetime(2026, 11, 20, 0, 0), datetime(2026, 11, 20, 23, 59)],
    "Alle Klassen Debattiertag": [datetime(2026, 11, 25, 0, 0), datetime(2026, 11, 25, 23, 59)],
    "4. Klassen mündliche Präsentationen Maturaaarbeit": [datetime(2027, 1, 8, 14, 0), datetime(2027, 1, 8, 23, 59)],
    "1. Klassen Skiabholen Sportamt": [datetime(2027, 1, 14, 8, 0), datetime(2027, 1, 14, 9, 30)],
    "1. Klassen Skirückgabe Sportamt": [datetime(2027, 1, 25, 8, 0), datetime(2027, 1, 25, 9, 30)],
    "Alle Klassen 3. Pädagogischer Halbtag": [datetime(2027, 2, 1, 14, 0), datetime(2027, 2, 1, 18, 0)],
    "Alle Klassen Klassenkonferenzen": [datetime(2027, 2, 2, 14, 0), datetime(2027, 2, 2, 18, 0)],
    "Alle Klassen 4. Pädagogischer Halbtag": [datetime(2027, 4, 22, 14, 0), datetime(2027, 4, 22, 18, 0)],
    "Alle Klassen Zeugnisklassenkonferenzen": [datetime(2027, 6, 24, 12, 30), datetime(2027, 6, 24, 18, 0)],
    "Alle Klassen Planungskonferenz": [datetime(2027, 6, 25, 12, 30), datetime(2027, 6, 25, 18, 0)],
}

SPEZIALWOCHEN = {
    "1. Klassen Regiowoche": [date(2026, 9, 21), date(2026, 9, 25)],
    "2. Klassen Theaterwoche": [date(2026, 9, 21), date(2026, 9, 25)],
    "3. Klassen SPF-Woche": [date(2026, 9, 21), date(2026, 9, 25)],
    "4. Klassen Bildungsreise": [date(2026, 9, 21), date(2026, 9, 25)],
    "1. Klassen Skilager": [date(2027, 1, 18), date(2027, 1, 22)],
    "2. Klassen Politik + Kulturwoche / Musikprojekt": [date(2027, 1, 18), date(2027, 1, 22)],
    "3. Klassen Politik + Kulturwoche / Musikprojekt": [date(2027, 1, 18), date(2027, 1, 22)],
    "4. Klassen Politik + Kulturwoche / Musikprojekt": [date(2027, 1, 18), date(2027, 1, 22)],
    "1. Klassen NaWi-Woche": [date(2027, 6, 28), date(2027, 7, 2)],
    "2. Klassen Sommersportlager": [date(2027, 6, 28), date(2027, 7, 2)],
    "3. Klassen Maturaarbeitswoche": [date(2027, 6, 28), date(2027, 7, 2)],
}

# Rohe Gestaltungswerte fürs Excel-Layout (Farben als Hex ohne '#').
# Die vier Schrift-Rollen entsprechen den vier Textarten im Layout: Fach-Titel
# in der Kopfzeile (header), Datum links/rechts (datum), normale Lektionszelle
# (normal), Ferien-/Frei-/Ausfall-Beschriftung (frei).
#
# Pro Rolle können "font_name", "font_size" und "farbe" (Schriftfarbe, Hex
# ohne '#') optional gesetzt werden und überschreiben dann nur für diese
# Rolle die Werte "font_name"/"font_size" unten bzw. die Standard-Schwarz-
# Schriftfarbe. Wird ein Schlüssel weggelassen, gilt der Default:
#   font_name -> STIL["font_name"], font_size -> STIL["font_size"],
#   farbe -> automatisch/Schwarz, bold/italic -> False.

TINT_FAKTOR = 0.55
STIL = {
    "font_name": "Aptos Narrow (Body)",
    "font_size": 18,
    "schriften": {
        "header": {"bold": True},
        "datum": {},
        "normal": {},
        "frei": {},
    },
    "farbe_ferien": tint("A9DEF9", TINT_FAKTOR),
    "farbe_frei": tint("A9DEF9", TINT_FAKTOR),
    "farbe_teilausfall": tint("A9DEF9", TINT_FAKTOR),
    "farbe_header": tint("FF99C8", TINT_FAKTOR),
    "farbe_datum": tint("E4C1F9", TINT_FAKTOR),
    "farbe_daten": tint("FCF6BD", TINT_FAKTOR),
    "farbe_prüfungen": tint("D0F4DE", TINT_FAKTOR),
    "farbe_leer": "F2F2F2",
    "farbe_rand": "000000",
    "zeilenhoehe": 128,
    "zeilenhoehe_header": 64,
    "spaltenbreite_daten": 48,
    "spaltenbreite_datum": 16,
}
