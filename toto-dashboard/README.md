# TOTO WM 2026 — Live-Rangliste

Statisches Dashboard, das die Tipp-Rangliste der Gruppenphase aus einem
Excel-Workbook berechnet und über GitHub Pages live anzeigt.

## Wie es funktioniert

1. Alle Tipps liegen im Workbook `data/TOTO_WM2026.xlsx` (ein Blatt pro
   Teilnehmer + ein Blatt `Realität` für die echten Ergebnisse).
2. `python run.py` liest die **Roh-Ergebnisse** (die eingetippten Tore) und
   berechnet alles in Python neu — Gruppentabellen, Sechzehntelfinalisten,
   Punkte und Rangliste. Die Excel-Formeln (SORTBY etc.) werden dafür *nicht*
   gebraucht.
3. Das Ergebnis wird als `docs/index.html` geschrieben.
4. GitHub Pages serviert `docs/index.html` als Website.

## Punktesystem

| Kategorie | Punkte |
|---|---|
| Richtiger Sieger / Unentschieden | 5 pro Spiel |
| Richtige Toranzahl | bis 5 pro Spiel (−1 pro Tor Abweichung) |
| Richtiger Sechzehntelfinalist | 5 pro Team |
| Playoff (später) | folgt |

**Sechzehntelfinalisten** zählen erst, sobald **alle 72 Gruppenspiele** im
Blatt `Realität` eingetragen sind. Solange nicht alle Resultate feststehen,
wäre ein Vergleich der Qualifikanten unsinnig (die Tabellen stehen noch nicht
fest). Bis dahin zeigt das Dashboard den Hinweis „Gruppenphase läuft".

Ein Spieler bekommt R32-Punkte nur, wenn er **selbst alle 72 Spiele** getippt
hat — ein leerer Tippzettel bekommt keine geschenkten Punkte.

## Lokal ausführen (PyCharm)

1. Projekt in PyCharm öffnen.
2. Interpreter einrichten und Abhängigkeiten installieren:
   ```
   pip install -r requirements.txt
   ```
3. `run.py` ausführen (Rechtsklick → *Run 'run'*).
4. `docs/index.html` im Browser öffnen, um das Ergebnis zu prüfen.
5. Änderungen committen und pushen:
   ```
   git add data/TOTO_WM2026.xlsx docs/index.html
   git commit -m "Resultate aktualisiert"
   git push
   ```

## GitHub Pages aktivieren (einmalig)

1. Repo auf GitHub anlegen und den Projektordner hochladen.
2. Im Repo: **Settings → Pages**.
3. Unter *Build and deployment*:
   - **Source**: „Deploy from a branch"
   - **Branch**: `main`, Ordner `/docs`
4. Speichern. Nach ein paar Minuten ist die Seite unter
   `https://<dein-username>.github.io/<repo-name>/` erreichbar.

## Automatische Aktualisierung (optional)

Die Datei `.github/workflows/update.yml` regeneriert das Dashboard automatisch,
sobald du eine neue Version von `data/TOTO_WM2026.xlsx` pushst. Du musst dann
`docs/index.html` nicht selbst erzeugen — es reicht, die Excel-Datei zu pushen.

Für stündliche Updates die `schedule`-Zeilen im Workflow einkommentieren.

## Playoff-Punkte später ergänzen

Die Funktion `compute_playoff_points()` in `src/compute.py` ist als Platzhalter
vorbereitet und gibt aktuell 0 zurück. Sobald das Playoff-Workbook steht:

1. Logik in `compute_playoff_points()` ergänzen (zweites Workbook laden,
   Sieger pro KO-Runde mit der Realität vergleichen, Punkte je Runde vergeben).
2. `run.py` mit dem Playoff-Pfad als zweitem Argument aufrufen, z.B.
   `python src/generate.py data/TOTO_WM2026.xlsx data/TOTO_Playoffs.xlsx`.

Die Playoff-Spalte erscheint im Dashboard automatisch, sobald ein Playoff-Pfad
übergeben wird (`playoff_active`). Die übrige Pipeline und das Layout sind
bereits darauf vorbereitet.

## Projektstruktur

```
toto-dashboard/
├── data/
│   └── TOTO_WM2026.xlsx        # Tipps + Realität (hier pflegst du die Daten)
├── src/
│   ├── compute.py              # Berechnungslogik (Punkte, Tabellen, Quali)
│   └── generate.py             # erzeugt docs/index.html
├── docs/
│   └── index.html              # generiert — von GitHub Pages serviert
├── .github/workflows/
│   └── update.yml              # optionale Automatik
├── run.py                      # Einstiegspunkt (in PyCharm ausführen)
├── requirements.txt
└── README.md
```
