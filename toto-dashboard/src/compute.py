"""
TOTO WM 2026 — Berechnungslogik.

Liest die Roh-Ergebnisse (eingetippte Tore) aus dem Excel-Workbook und
berechnet die gesamte Rangliste neu in Python. Verlaesst sich NICHT auf
die Excel-Formeln (insb. nicht auf das Excel-exklusive SORTBY), damit das
Dashboard auch dann funktioniert, wenn die Datei nie in Excel geoeffnet wurde.

Punktesystem (Gruppenphase):
  - Richtiger Sieger / Unentschieden: 5 Punkte
  - Richtige Toranzahl: max. 5 Punkte (minus 1 pro Tor Abweichung, min. 0)
  - Richtiger Sechzehntelfinalist: 5 Punkte je korrekt getipptes Team
    -> zaehlt ERST, wenn alle Gruppenspiele in 'Realitaet' eingetragen sind.

Playoff-Punkte sind als Hook vorbereitet (compute_playoff_points), damit sie
spaeter aus einem zweiten Workbook live dazukommen koennen.
"""

from openpyxl import load_workbook

# ── Layout des Gruppenphasen-Sheets ─────────────────────────────
# Spaltenpaare (Heim, Gast) fuer die 6 Gruppen pro Haelfte
COL_PAIRS = [('B', 'C'), ('I', 'J'), ('P', 'Q'),
             ('W', 'X'), ('AD', 'AE'), ('AK', 'AL')]
GRP_TOP = ['A', 'B', 'C', 'D', 'E', 'F']
GRP_BOT = ['G', 'H', 'I', 'J', 'K', 'L']
# (Zeile der Teamnamen, Zeile der Tore) fuer die 6 Spiele
TOP_ROWS = [(4, 5), (9, 10), (15, 16), (20, 21), (26, 27), (31, 32)]
BOT_ROWS = [(51, 52), (56, 57), (62, 63), (67, 68), (73, 74), (78, 79)]

# Sheets, die keine Teilnehmer sind
NON_PLAYER_SHEETS = {'Auswertung', 'Realität', 'Realitaet', 'Anleitung'}


def build_group_layout():
    """Gibt dict zurueck: Gruppe -> Liste von (heim_col, gast_col, team_row, score_row)."""
    layout = {}
    for gi, (c1, c2) in enumerate(COL_PAIRS):
        layout[GRP_TOP[gi]] = [(c1, c2, tr, sr) for (tr, sr) in TOP_ROWS]
    for gi, (c1, c2) in enumerate(COL_PAIRS):
        layout[GRP_BOT[gi]] = [(c1, c2, tr, sr) for (tr, sr) in BOT_ROWS]
    return layout


def _num(v):
    """Wandelt einen Zellwert in int um, oder None wenn leer/ungueltig."""
    if v is None or v == '':
        return None
    try:
        return int(v)
    except (TypeError, ValueError):
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return None


def read_sheet_matches(ws, layout):
    """Liest alle Spiele eines Sheets.

    Gibt dict zurueck: Gruppe -> Liste von (heim_team, gast_team, heim_tore, gast_tore).
    Tore sind int oder None (noch nicht eingetragen).
    """
    result = {}
    for grp, matches in layout.items():
        ms = []
        for (c1, c2, tr, sr) in matches:
            ht = ws[f'{c1}{tr}'].value
            at = ws[f'{c2}{tr}'].value
            hg = _num(ws[f'{c1}{sr}'].value)
            ag = _num(ws[f'{c2}{sr}'].value)
            ht = ht.strip() if isinstance(ht, str) else ht
            at = at.strip() if isinstance(at, str) else at
            ms.append((ht, at, hg, ag))
        result[grp] = ms
    return result


def compute_standings(matches):
    """Berechnet die Gruppentabelle.

    matches: Liste von (heim_team, gast_team, heim_tore, gast_tore).
    Gibt (ranking, table) zurueck.
      ranking: Liste der Teams sortiert nach Punkte, Tordiff, Tore (absteigend)
      table:   dict Team -> {'pts','gf','ga','played'}
    """
    table = {}

    def ensure(t):
        if t and t not in table:
            table[t] = {'pts': 0, 'gf': 0, 'ga': 0, 'played': 0}

    for ht, at, hg, ag in matches:
        ensure(ht)
        ensure(at)
        if hg is None or ag is None:
            continue
        table[ht]['gf'] += hg
        table[ht]['ga'] += ag
        table[at]['gf'] += ag
        table[at]['ga'] += hg
        table[ht]['played'] += 1
        table[at]['played'] += 1
        if hg > ag:
            table[ht]['pts'] += 3
        elif hg < ag:
            table[at]['pts'] += 3
        else:
            table[ht]['pts'] += 1
            table[at]['pts'] += 1

    # Sortierung: Punkte, Tordifferenz, erzielte Tore (alle absteigend),
    # bei Gleichstand stabil nach Teamname (deterministisch)
    ranking = sorted(
        table.keys(),
        key=lambda t: (table[t]['pts'],
                       table[t]['gf'] - table[t]['ga'],
                       table[t]['gf'],
                       t),
        reverse=True)
    return ranking, table


def compute_qualifiers(all_standings):
    """Bestimmt die 32 Sechzehntelfinalisten.

    Top 2 jeder Gruppe (24) + die 8 besten Gruppendritten.
    all_standings: dict Gruppe -> (ranking, table).
    Gibt ein set von Teamnamen zurueck.
    """
    qualifiers = set()
    thirds = []
    for grp, (ranking, table) in all_standings.items():
        if len(ranking) >= 1:
            qualifiers.add(ranking[0])
        if len(ranking) >= 2:
            qualifiers.add(ranking[1])
        if len(ranking) >= 3:
            t = ranking[2]
            thirds.append((table[t]['pts'],
                           table[t]['gf'] - table[t]['ga'],
                           table[t]['gf'],
                           t))
    thirds.sort(reverse=True)
    for _, _, _, t in thirds[:8]:
        qualifiers.add(t)
    return qualifiers


def match_points(p_matches, r_matches):
    """Punkte eines Spielers fuer alle Spiele einer Gruppe.

    Gibt (sieger_punkte, tore_punkte) zurueck.
    Punkte nur, wenn sowohl Tipp als auch Realitaet eingetragen sind.
    """
    sieger = 0
    tore = 0
    for (pht, pat, phg, pag), (rht, rat, rhg, rag) in zip(p_matches, r_matches):
        if None in (phg, pag, rhg, rag):
            continue
        ps = (phg > pag) - (phg < pag)   # 1 / 0 / -1
        rs = (rhg > rag) - (rhg < rag)
        if ps == rs:
            sieger += 5
        deviation = abs(phg - rhg) + abs(pag - rag)
        tore += max(0, 5 - deviation)
    return sieger, tore


def group_stage_complete(real_matches):
    """True, wenn alle 72 Gruppenspiele in 'Realitaet' eingetragen sind."""
    for grp, ms in real_matches.items():
        for ht, at, hg, ag in ms:
            if hg is None or ag is None:
                return False
    return True


def player_predictions_complete(p_matches):
    """True, wenn ein Spieler alle 72 Spiele getippt hat.

    Nur dann werden Sechzehntelfinalisten-Punkte vergeben, denn die
    Qualifikanten-Vorhersage ergibt nur Sinn, wenn der ganze Tippzettel
    ausgefuellt ist (ein leerer Zettel wuerde sonst zufaellig Punkte
    erhalten, weil bei 0:0 ueberall die Sortierung Treffer produziert).
    """
    for grp, ms in p_matches.items():
        for ht, at, hg, ag in ms:
            if hg is None or ag is None:
                return False
    return True


def compute_playoff_points(player_name, playoff_path=None):
    """Hook fuer spaetere Playoff-Punkte.

    Aktuell 0. Sobald ein Playoff-Workbook existiert, kann hier die Logik
    ergaenzt werden (z.B. zweites Workbook laden, Sieger pro KO-Runde
    vergleichen, Punkte je Runde vergeben). Die uebrige Pipeline und das
    Dashboard zeigen die Spalte bereits an.
    """
    if not playoff_path:
        return 0
    # TODO: Playoff-Auswertung implementieren, sobald das Sheet steht.
    return 0


def find_players(wb):
    """Alle Teilnehmer-Sheets (alles ausser Auswertung/Realitaet)."""
    return [s for s in wb.sheetnames if s not in NON_PLAYER_SHEETS]


def compute_all(xlsx_path, playoff_path=None):
    """Hauptfunktion: berechnet die komplette Rangliste.

    Gibt dict zurueck:
      {
        'group_stage_complete': bool,
        'r32_counts': bool,            # ob R32-Punkte aktuell zaehlen
        'playoff_active': bool,        # ob Playoff-Punkte aktuell zaehlen
        'players': [ {rank,name,sieger,tore,r32,playoff,total,...}, ... ]
      }
    """
    wb = load_workbook(xlsx_path, data_only=False)
    layout = build_group_layout()

    if 'Realität' in wb.sheetnames:
        real_ws = wb['Realität']
    elif 'Realitaet' in wb.sheetnames:
        real_ws = wb['Realitaet']
    else:
        raise ValueError("Kein 'Realität'-Sheet gefunden.")

    real_matches = read_sheet_matches(real_ws, layout)
    real_standings = {g: compute_standings(ms) for g, ms in real_matches.items()}
    complete = group_stage_complete(real_matches)
    actual_quali = compute_qualifiers(real_standings) if complete else set()

    playoff_active = bool(playoff_path)

    players = []
    for name in find_players(wb):
        ws = wb[name]
        p_matches = read_sheet_matches(ws, layout)

        sieger_total = 0
        tore_total = 0
        for g in layout:
            s, t = match_points(p_matches[g], real_matches[g])
            sieger_total += s
            tore_total += t

        r32 = 0
        if complete:
            p_standings = {g: compute_standings(ms) for g, ms in p_matches.items()}
            if player_predictions_complete(p_matches):
                p_quali = compute_qualifiers(p_standings)
                r32 = 5 * len(p_quali & actual_quali)

        playoff = compute_playoff_points(name, playoff_path)

        players.append({
            'name': name,
            'sieger': sieger_total,
            'tore': tore_total,
            'r32': r32,
            'playoff': playoff,
            'gruppenphase': sieger_total + tore_total + r32,
            'total': sieger_total + tore_total + r32 + playoff,
        })

    # Sortieren + Raenge (gleiche Punkte -> gleicher Rang)
    players.sort(key=lambda x: x['total'], reverse=True)
    last_total = None
    last_rank = 0
    for i, p in enumerate(players):
        if p['total'] != last_total:
            last_rank = i + 1
            last_total = p['total']
        p['rank'] = last_rank

    return {
        'group_stage_complete': complete,
        'r32_counts': complete,
        'playoff_active': playoff_active,
        'players': players,
    }


if __name__ == '__main__':
    import sys
    import json
    path = sys.argv[1] if len(sys.argv) > 1 else 'data/TOTO_WM2026.xlsx'
    result = compute_all(path)
    print(json.dumps(result, indent=2, ensure_ascii=False))
