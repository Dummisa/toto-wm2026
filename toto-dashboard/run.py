"""Einstiegspunkt: Rangliste berechnen und Dashboard erzeugen.

In PyCharm einfach diese Datei ausfuehren (Rechtsklick -> Run 'run').
Danach docs/index.html committen und pushen -> GitHub Pages aktualisiert sich.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
from generate import main  # noqa: E402

if __name__ == '__main__':
    main()
