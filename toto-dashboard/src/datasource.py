"""
Bestimmt, woher die Excel-Daten kommen.

- Wenn die Umgebungsvariable TOTO_DATA_URL gesetzt ist (z.B. ein
  OneDrive-Freigabelink), wird die Datei heruntergeladen.
- Sonst wird die lokale Datei data/TOTO_WM2026.xlsx verwendet.

So laeuft dasselbe Projekt sowohl lokal (mit Datei im Ordner) als auch in
der Cloud (GitHub Actions zieht die Datei vom OneDrive-Link) — ohne
Code-Aenderung. Es wird nur die Standardbibliothek genutzt (kein 'requests').
"""

import base64
import os
import tempfile
import urllib.request

HERE = os.path.dirname(__file__)
ROOT = os.path.dirname(HERE)
LOCAL_DATA = os.path.join(ROOT, 'data', 'TOTO_WM2026.xlsx')


def onedrive_to_direct(url):
    """Wandelt einen OneDrive-Freigabelink in einen Direkt-Download-Link.

    Funktioniert fuer persoenliche OneDrive-Links (1drv.ms / onedrive.live.com)
    ueber die oeffentliche 'shares'-API. Andere URLs werden unveraendert
    zurueckgegeben (z.B. wenn schon ein Direktlink uebergeben wird).
    """
    u = url.strip()
    is_onedrive = ('1drv.ms' in u or 'onedrive.live.com' in u)
    if not is_onedrive:
        return u
    b64 = base64.urlsafe_b64encode(u.encode('utf-8')).decode('ascii')
    token = 'u!' + b64.rstrip('=')
    return f'https://api.onedrive.com/v1.0/shares/{token}/root/content'


def _download(url, dest):
    direct = onedrive_to_direct(url)
    req = urllib.request.Request(direct, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = resp.read()
    if len(data) < 1000:
        raise RuntimeError(
            "Heruntergeladene Datei ist verdaechtig klein — stimmt der "
            "OneDrive-Link und ist die Freigabe auf 'Jeder mit dem Link'?")
    with open(dest, 'wb') as f:
        f.write(data)
    return dest


def resolve_data_path(env_var='TOTO_DATA_URL', local_default=LOCAL_DATA):
    """Gibt den Pfad zur Excel-Datei zurueck (lokal oder heruntergeladen)."""
    url = os.environ.get(env_var, '').strip()
    if url:
        tmp = os.path.join(tempfile.gettempdir(), 'toto_data.xlsx')
        print(f"Lade Daten von URL ({env_var}) ...")
        return _download(url, tmp)
    print(f"Verwende lokale Datei: {local_default}")
    return local_default


def resolve_playoff_path(env_var='TOTO_PLAYOFF_URL'):
    """Optionaler Playoff-Datenpfad (fuer spaeter). None wenn nicht gesetzt."""
    url = os.environ.get(env_var, '').strip()
    if not url:
        return None
    tmp = os.path.join(tempfile.gettempdir(), 'toto_playoff.xlsx')
    print(f"Lade Playoff-Daten von URL ({env_var}) ...")
    return _download(url, tmp)
