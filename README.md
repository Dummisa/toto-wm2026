# TOTO WM 2026 — Live-Rangliste (Cloud-Version)

Das Scoreboard wird jetzt **in der Cloud** erzeugt (GitHub Actions) und über
GitHub Pages veröffentlicht. Du brauchst kein Python und kein PyCharm mehr auf
einem bestimmten Gerät — du bearbeitest nur die OneDrive-Excel und löst die
Aktualisierung von überall aus (auch per GitHub-App auf dem Handy).

## So funktioniert es

1. Die Tipps liegen in **einer Excel-Datei auf OneDrive** (ein Blatt pro
   Teilnehmer + ein Blatt `Realität`).
2. GitHub Actions lädt diese Datei über einen OneDrive-Link herunter, berechnet
   die Rangliste in Python neu und schreibt `docs/index.html`.
3. GitHub Pages zeigt die Seite an.

Du musst also nur zwei Dinge tun, wenn neue Resultate da sind:
**(a)** Ergebnisse in der OneDrive-Excel eintragen, **(b)** den Workflow starten.

## Einmalige Einrichtung

### 1. OneDrive-Freigabelink holen
- In OneDrive die Datei `TOTO_WM2026.xlsx` auswählen → **Teilen** →
  Berechtigung auf **„Jeder mit dem Link kann anzeigen"** stellen → **Link
  kopieren**. (Wichtig: Ansicht reicht, kein Bearbeitungslink nötig.)

### 2. Link in GitHub hinterlegen
- Im Repo: **Settings → Secrets and variables → Actions → Tab „Variables"
  → New repository variable**.
- Name: `TOTO_DATA_URL`
- Value: der kopierte OneDrive-Link → **Add variable**.

### 3. GitHub Pages auf „Actions" stellen
- **Settings → Pages → Source: „GitHub Actions"**.

Das war's. Ab jetzt nie mehr nötig.

## Aktualisieren — von jedem Gerät

1. Resultate in der OneDrive-Excel eintragen (Excel-App auf Laptop, Tablet
   oder Handy — egal welches Gerät, Hauptsache OneDrive synchronisiert).
2. Auf github.com (oder in der **GitHub-Handy-App**): Tab **Actions** →
   Workflow **„Update Scoreboard"** → **„Run workflow"**.
3. Nach 1–2 Minuten ist das Scoreboard aktualisiert unter
   `https://<dein-benutzername>.github.io/<repo>/`.

### Noch bequemer: automatisch
Im Workflow `.github/workflows/scoreboard.yml` die zwei `schedule`-Zeilen
einkommentieren — dann aktualisiert sich das Scoreboard z.B. alle 30 Minuten
von selbst, ganz ohne Tippen.

## Lokal läuft es weiterhin

Wenn du doch mal lokal arbeitest: Datei unter `data/TOTO_WM2026.xlsx` ablegen
und `python run.py` ausführen. Ist die Umgebungsvariable `TOTO_DATA_URL` nicht
gesetzt, nimmt das Script automatisch die lokale Datei.

## Hinweise / Stolpersteine

- Der OneDrive-Link muss auf **„Jeder mit dem Link"** stehen, sonst kann die
  Cloud die Datei nicht laden.
- Diese Methode ist für **persönliches OneDrive** gebaut (1drv.ms /
  onedrive.live.com). Bei **OneDrive for Business / SharePoint** funktioniert
  der automatische Direktlink evtl. nicht — dann an den Freigabelink `?download=1`
  anhängen und diesen als `TOTO_DATA_URL` verwenden.
- Die Berechnung liest nur die **eingetippten Tore** und rechnet alles selbst
  neu; die Excel-Formeln (SORTBY) müssen also nicht funktionieren.

## Playoff-Punkte später

In `src/compute.py` ist `compute_playoff_points()` als Platzhalter vorbereitet.
Sobald das Playoff-Sheet steht: Logik dort ergänzen, eine zweite Variable
`TOTO_PLAYOFF_URL` im Repo anlegen und im Workflow die auskommentierte Zeile
aktivieren. Die Playoff-Spalte erscheint dann automatisch im Dashboard.

## Struktur

```
toto-dashboard/
├── data/TOTO_WM2026.xlsx          # nur für lokale Läufe (optional)
├── src/
│   ├── compute.py                 # Berechnung (Punkte, Tabellen, Quali)
│   ├── datasource.py              # holt Daten lokal ODER vom OneDrive-Link
│   └── generate.py                # erzeugt docs/index.html
├── docs/index.html                # generiert — von GitHub Pages serviert
├── .github/workflows/scoreboard.yml  # Cloud: erzeugen + veröffentlichen
├── run.py
├── requirements.txt
└── README.md
```
