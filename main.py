"""
Führt die komplette Pipeline aus:
  1. Jedes Untis-Stundenplan-PDF in input_files/ wird zu einem
     Stundenplan-Export in output_files/ konvertiert (siehe
     pdf_timetable_to_xlsx.py).
  2. Aus diesen Exporten wird pro Klasse eine Planungs-Excel-Datei erzeugt
     (siehe klassenplanung.py).

Der Name eines Exports wird aus dem PDF-Dateinamen abgeleitet:
  'leh a4q_KÜ_phasen_1_3.pdf' -> 'Stundenplan_KÜ_Phasen_1_3.xlsx'
Setzt voraus, dass alle PDFs in input_files/ diesem Namensschema
'..._<Kürzel>_phasen_<von>_<bis>.pdf' folgen.
"""

import re
from pathlib import Path

from jahresuebersicht import erstelle_jahresuebersicht_datei
from klassenplanung import erstelle_alle_planungen
from pdf_timetable_to_xlsx import build_workbook

PDF_NAME_RE = re.compile(r"_(?P<kuerzel>[^_]+)_phasen_(?P<von>\d+)_(?P<bis>\d+)\.pdf$", re.IGNORECASE)


def ausgabename_fuer_pdf(pdf_pfad):
    """
    Leitet aus einem Eingabe-PDF wie 'leh a4q_KÜ_phasen_1_3.pdf' den Namen
    des Stundenplan-Exports ab: 'Stundenplan_KÜ_Phasen_1_3.xlsx'.
    """
    m = PDF_NAME_RE.search(pdf_pfad.name)
    if not m:
        raise ValueError(
            f"PDF-Name '{pdf_pfad.name}' hat nicht das erwartete Format "
            f"'..._<Kürzel>_phasen_<von>_<bis>.pdf'."
        )
    return f"Stundenplan_{m['kuerzel']}_Phasen_{m['von']}_{m['bis']}.xlsx"


def main():
    basis_dir = Path(__file__).parent
    input_dir = basis_dir / "input_files"
    output_dir = basis_dir / "output_files"
    ausgabe_ordner = output_dir / "klassen"

    output_dir.mkdir(parents=True, exist_ok=True)

    for pdf_pfad in sorted(input_dir.glob("*.pdf")):
        ziel_name = ausgabename_fuer_pdf(pdf_pfad)
        wb = build_workbook(pdf_pfad)
        wb.save(output_dir / ziel_name)
        print(f"Konvertiert: {pdf_pfad.name} -> {ziel_name}")

    erstelle_alle_planungen(output_dir, ausgabe_ordner)
    erstelle_jahresuebersicht_datei(output_dir, output_dir / "Jahresuebersicht.xlsx")


if __name__ == "__main__":
    main()
