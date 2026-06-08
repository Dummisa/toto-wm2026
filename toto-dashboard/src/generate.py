"""
Erzeugt aus dem Excel-Workbook die statische Dashboard-Seite (docs/index.html)
fuer GitHub Pages. Die Daten werden direkt in die HTML-Datei eingebettet,
damit es keine fetch/CORS-Probleme gibt.

Aufruf:
    python src/generate.py            # nutzt data/TOTO_WM2026.xlsx
    python src/generate.py pfad.xlsx  # eigener Pfad
"""

import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(__file__))
from compute import compute_all  # noqa: E402

HERE = os.path.dirname(__file__)
ROOT = os.path.dirname(HERE)
DEFAULT_XLSX = os.path.join(ROOT, 'data', 'TOTO_WM2026.xlsx')
OUTPUT = os.path.join(ROOT, 'docs', 'index.html')


def render_html(result):
    """Baut die komplette HTML-Seite mit eingebetteten Daten."""
    updated = datetime.now(timezone.utc).astimezone().strftime('%d.%m.%Y %H:%M')
    payload = json.dumps(result, ensure_ascii=False)

    if result['group_stage_complete']:
        status_text = 'Gruppenphase abgeschlossen — Sechzehntelfinalisten zählen'
        status_class = 'done'
    else:
        status_text = 'Gruppenphase läuft — Sechzehntelfinalisten zählen noch nicht'
        status_class = 'running'

    return f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>TOTO WM 2026 — Rangliste</title>
<style>
  :root {{
    --bg: #0f1117;
    --card: #181b24;
    --row: #1e222e;
    --row-alt: #20242f;
    --text: #e8eaf0;
    --muted: #8b90a0;
    --accent: #3b82f6;
    --gold: #ffd700;
    --silver: #c0c0c0;
    --bronze: #cd7f32;
    --green: #22c55e;
    --border: #2a2f3d;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    padding: 24px 16px 60px;
    line-height: 1.5;
  }}
  .wrap {{ max-width: 860px; margin: 0 auto; }}
  h1 {{
    font-size: 1.7rem;
    font-weight: 800;
    letter-spacing: -0.02em;
    margin-bottom: 4px;
  }}
  .sub {{ color: var(--muted); font-size: 0.85rem; margin-bottom: 20px; }}
  .status {{
    display: inline-flex; align-items: center; gap: 8px;
    padding: 8px 14px; border-radius: 10px; font-size: 0.85rem;
    font-weight: 600; margin-bottom: 22px;
  }}
  .status.running {{ background: rgba(251,191,36,0.12); color: #fbbf24; }}
  .status.done {{ background: rgba(34,197,94,0.12); color: var(--green); }}
  .status .dot {{ width: 8px; height: 8px; border-radius: 50%; background: currentColor; }}
  table {{ width: 100%; border-collapse: collapse; background: var(--card);
           border-radius: 14px; overflow: hidden; }}
  thead th {{
    text-align: right; padding: 13px 12px; font-size: 0.72rem;
    text-transform: uppercase; letter-spacing: 0.05em; color: var(--muted);
    background: #14171f; font-weight: 700; border-bottom: 1px solid var(--border);
  }}
  thead th.l {{ text-align: left; }}
  tbody td {{ padding: 13px 12px; text-align: right; font-variant-numeric: tabular-nums;
              border-bottom: 1px solid var(--border); }}
  tbody td.l {{ text-align: left; }}
  tbody tr:nth-child(even) {{ background: var(--row-alt); }}
  tbody tr:last-child td {{ border-bottom: none; }}
  .rank {{ font-weight: 800; width: 42px; }}
  .rank.r1 {{ color: var(--gold); }}
  .rank.r2 {{ color: var(--silver); }}
  .rank.r3 {{ color: var(--bronze); }}
  .name {{ font-weight: 600; }}
  .total {{ font-weight: 800; font-size: 1.05rem; }}
  .muted-col {{ color: var(--muted); }}
  .po-hidden {{ display: none; }}
  footer {{ margin-top: 18px; color: var(--muted); font-size: 0.78rem; text-align: center; }}
  @media (max-width: 600px) {{
    .hide-sm {{ display: none; }}
    h1 {{ font-size: 1.4rem; }}
    tbody td, thead th {{ padding: 11px 8px; }}
  }}
</style>
</head>
<body>
<div class="wrap">
  <h1>TOTO WM 2026</h1>
  <div class="sub">Live-Rangliste Gruppenphase · zuletzt aktualisiert: {updated}</div>
  <div class="status {status_class}"><span class="dot"></span>{status_text}</div>
  <table id="board">
    <thead>
      <tr>
        <th class="l">#</th>
        <th class="l">Name</th>
        <th class="hide-sm">Spieltipps</th>
        <th class="hide-sm">Tortipps</th>
        <th>16tel</th>
        <th class="po-col po-hidden">Playoff</th>
        <th>Total</th>
      </tr>
    </thead>
    <tbody></tbody>
  </table>
  <footer>Automatisch generiert aus dem Tipp-Workbook · {len(result['players'])} Teilnehmer</footer>
</div>
<script>
  const DATA = {payload};

  function rankClass(r) {{ return r <= 3 ? 'rank r' + r : 'rank'; }}

  function render() {{
    const tbody = document.querySelector('#board tbody');
    tbody.innerHTML = '';
    // Playoff-Spalte nur zeigen, wenn aktiv
    if (DATA.playoff_active) {{
      document.querySelectorAll('.po-col').forEach(e => e.classList.remove('po-hidden'));
    }}
    DATA.players.forEach(p => {{
      const tr = document.createElement('tr');
      const poCell = DATA.playoff_active
        ? `<td class="po-col">${{p.playoff}}</td>` : '';
      tr.innerHTML = `
        <td class="l ${{rankClass(p.rank)}}">${{p.rank}}</td>
        <td class="l name">${{p.name}}</td>
        <td class="hide-sm muted-col">${{p.sieger}}</td>
        <td class="hide-sm muted-col">${{p.tore}}</td>
        <td class="muted-col">${{p.r32}}</td>
        ${{poCell}}
        <td class="total">${{p.total}}</td>`;
      tbody.appendChild(tr);
    }});
  }}
  render();
</script>
</body>
</html>
"""


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSX
    playoff = sys.argv[2] if len(sys.argv) > 2 else None
    result = compute_all(xlsx, playoff_path=playoff)
    html = render_html(result)
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Dashboard erzeugt: {OUTPUT}")
    print(f"  Teilnehmer: {len(result['players'])}")
    print(f"  Gruppenphase abgeschlossen: {result['group_stage_complete']}")
    print(f"  Playoff aktiv: {result['playoff_active']}")
    if result['players']:
        top = result['players'][0]
        print(f"  Aktuell Erster: {top['name']} ({top['total']} Pkt)")


if __name__ == '__main__':
    main()
