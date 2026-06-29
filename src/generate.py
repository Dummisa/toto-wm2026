"""Erzeugt das kombinierte Scoreboard (Gruppenphase + Playoffs) als docs/index.html."""

import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(__file__))
from compute import compute_group                       # noqa: E402
from compute_playoffs import compute_playoffs           # noqa: E402
from datasource import resolve_data_path, resolve_playoff_path  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(__file__))
OUTPUT = os.path.join(ROOT, 'docs', 'index.html')


def build_result():
    group_path = resolve_data_path()
    playoff_path = resolve_playoff_path()

    group, group_complete = ({}, False)
    if group_path:
        group, group_complete = compute_group(group_path)

    playoff = {}
    if playoff_path:
        playoff = compute_playoffs(playoff_path)

    playoff_active = bool(playoff) and any(p['total'] > 0 for p in playoff.values())

    names = sorted(set(group) | set(playoff))
    players = []
    for n in names:
        g = group.get(n, {}).get('total', 0)
        po = playoff.get(n, {}).get('total', 0)
        players.append({'name': n, 'gruppe': g, 'playoff': po, 'total': g + po})

    players.sort(key=lambda x: x['total'], reverse=True)
    last, rank = None, 0
    for i, p in enumerate(players):
        if p['total'] != last:
            rank, last = i + 1, p['total']
        p['rank'] = rank

    return {
        'group_stage_complete': group_complete,
        'playoff_active': playoff_active,
        'players': players,
    }


def render_html(result):
    updated = datetime.now(timezone.utc).astimezone().strftime('%d.%m.%Y %H:%M')
    payload = json.dumps(result, ensure_ascii=False)
    if result['playoff_active']:
        status_text, status_class = 'Playoffs laufen', 'done'
    elif result['group_stage_complete']:
        status_text, status_class = 'Gruppenphase abgeschlossen', 'done'
    else:
        status_text, status_class = 'Gruppenphase läuft', 'running'

    return f"""<!DOCTYPE html>
<html lang="de"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>TOTO WM 2026 — Rangliste</title>
<style>
  :root {{ --bg:#0f1117; --card:#181b24; --row-alt:#20242f; --text:#e8eaf0; --muted:#8b90a0;
    --gold:#ffd700; --silver:#c0c0c0; --bronze:#cd7f32; --green:#22c55e; --border:#2a2f3d; }}
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{ font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
    background:var(--bg); color:var(--text); padding:24px 16px 60px; line-height:1.5; }}
  .wrap {{ max-width:760px; margin:0 auto; }}
  h1 {{ font-size:1.7rem; font-weight:800; letter-spacing:-.02em; margin-bottom:4px; }}
  .sub {{ color:var(--muted); font-size:.85rem; margin-bottom:20px; }}
  .status {{ display:inline-flex; align-items:center; gap:8px; padding:8px 14px;
    border-radius:10px; font-size:.85rem; font-weight:600; margin-bottom:22px; }}
  .status.running {{ background:rgba(251,191,36,.12); color:#fbbf24; }}
  .status.done {{ background:rgba(34,197,94,.12); color:var(--green); }}
  .status .dot {{ width:8px; height:8px; border-radius:50%; background:currentColor; }}
  table {{ width:100%; border-collapse:collapse; background:var(--card);
    border-radius:14px; overflow:hidden; }}
  thead th {{ text-align:right; padding:13px 12px; font-size:.72rem; text-transform:uppercase;
    letter-spacing:.05em; color:var(--muted); background:#14171f; font-weight:700;
    border-bottom:1px solid var(--border); }}
  thead th.l {{ text-align:left; }}
  tbody td {{ padding:13px 12px; text-align:right; font-variant-numeric:tabular-nums;
    border-bottom:1px solid var(--border); }}
  tbody td.l {{ text-align:left; }}
  tbody tr:nth-child(even) {{ background:var(--row-alt); }}
  tbody tr:last-child td {{ border-bottom:none; }}
  .rank {{ font-weight:800; width:42px; }}
  .rank.r1 {{ color:var(--gold); }} .rank.r2 {{ color:var(--silver); }} .rank.r3 {{ color:var(--bronze); }}
  .name {{ font-weight:600; }} .total {{ font-weight:800; font-size:1.05rem; }}
  .muted-col {{ color:var(--muted); }} .po-hidden {{ display:none; }}
  footer {{ margin-top:18px; color:var(--muted); font-size:.78rem; text-align:center; }}
  @media (max-width:600px) {{ h1 {{ font-size:1.4rem; }} tbody td,thead th {{ padding:11px 8px; }} }}
</style></head><body>
<div class="wrap">
  <h1>TOTO WM 2026</h1>
  <div class="sub">Live-Rangliste · zuletzt aktualisiert: {updated}</div>
  <div class="status {status_class}"><span class="dot"></span>{status_text}</div>
  <table id="board"><thead><tr>
    <th class="l">#</th><th class="l">Name</th>
    <th>Gruppenphase</th><th class="po-col po-hidden">Playoffs</th><th>Total</th>
  </tr></thead><tbody></tbody></table>
  <footer>Automatisch generiert · {len(result['players'])} Teilnehmer</footer>
</div>
<script>
  const DATA = {payload};
  const rc = r => r <= 3 ? 'rank r' + r : 'rank';
  (function(){{
    const tb = document.querySelector('#board tbody');
    if (DATA.playoff_active) document.querySelectorAll('.po-col').forEach(e=>e.classList.remove('po-hidden'));
    DATA.players.forEach(p => {{
      const tr = document.createElement('tr');
      const po = DATA.playoff_active ? `<td class="po-col muted-col">${{p.playoff}}</td>` : '';
      tr.innerHTML = `<td class="l ${{rc(p.rank)}}">${{p.rank}}</td>`
        + `<td class="l name">${{p.name}}</td>`
        + `<td class="muted-col">${{p.gruppe}}</td>` + po
        + `<td class="total">${{p.total}}</td>`;
      tb.appendChild(tr);
    }});
  }})();
</script></body></html>
"""


def main():
    result = build_result()
    html = render_html(result)
    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)
    with open(OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Dashboard erzeugt: {OUTPUT}")
    print(f"  Teilnehmer: {len(result['players'])}")
    print(f"  Playoffs aktiv: {result['playoff_active']}")
    if result['players']:
        t = result['players'][0]
        print(f"  Erster: {t['name']} (Gruppe {t['gruppe']} + Playoff {t['playoff']} = {t['total']})")


if __name__ == '__main__':
    main()
