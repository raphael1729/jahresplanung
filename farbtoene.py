"""
Berechnet aufgehellte Varianten einer Hex-Farbe (Tint = Mischung mit Weiss),
z. B. um einen Wert für STIL in config.py zu bestimmen.

Aufruf: python farbtoene.py <hex> <alpha> [<alpha> ...]
Beispiel: python farbtoene.py FF006E 0.3 0.2
"""

import sys


def tint(hex_farbe, alpha):
    """Mischt hex_farbe (ohne '#') mit Weiss. alpha=1.0 -> Originalfarbe,
    alpha=0.0 -> Weiss."""
    r, g, b = int(hex_farbe[0:2], 16), int(hex_farbe[2:4], 16), int(hex_farbe[4:6], 16)
    r2 = round(r * alpha + 255 * (1 - alpha))
    g2 = round(g * alpha + 255 * (1 - alpha))
    b2 = round(b * alpha + 255 * (1 - alpha))
    return f"{r2:02X}{g2:02X}{b2:02X}"


if __name__ == "__main__":
    hex_farbe = sys.argv[1].lstrip("#")
    for wert in sys.argv[2:]:
        alpha = float(wert)
        print(f"alpha={alpha:g}: {tint(hex_farbe, alpha)}")
