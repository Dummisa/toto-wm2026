"""Einstiegspunkt: Daten holen (lokal oder OneDrive), Scoreboard erzeugen.

Lokal in PyCharm: einfach ausfuehren -> nutzt data/TOTO_WM2026.xlsx.
In der Cloud: GitHub Actions setzt TOTO_DATA_URL -> Datei wird heruntergeladen.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
from generate import main  # noqa: E402

if __name__ == '__main__':
    main()
