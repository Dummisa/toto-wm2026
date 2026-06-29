"""Bestimmt die Datenquellen (lokal oder per URL) fuer Gruppenphase + Playoffs.

- TOTO_DATA_URL    -> Gruppenphasen-Workbook (sonst data/TOTO_WM2026.xlsx)
- TOTO_PLAYOFF_URL -> Playoff-Workbook       (sonst data/TOTO_Playoffs.xlsx)

Fehlt eine Datei (kein URL, lokal nicht vorhanden), wird None zurueckgegeben.
Nur Standardbibliothek (kein 'requests').
"""

import base64
import os
import tempfile
import urllib.request

HERE = os.path.dirname(__file__)
ROOT = os.path.dirname(HERE)
LOCAL_GROUP = os.path.join(ROOT, 'data', 'TOTO_WM2026.xlsx')
LOCAL_PLAYOFF = os.path.join(ROOT, 'data', 'TOTO_Playoffs.xlsx')


def onedrive_to_direct(url):
    u = url.strip()
    if '1drv.ms' in u or 'onedrive.live.com' in u:
        b64 = base64.urlsafe_b64encode(u.encode('utf-8')).decode('ascii')
        return f'https://api.onedrive.com/v1.0/shares/u!{b64.rstrip("=")}/root/content'
    return u


def _download(url, dest):
    direct = onedrive_to_direct(url)
    req = urllib.request.Request(direct, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = resp.read()
    if len(data) < 1000:
        raise RuntimeError("Heruntergeladene Datei verdaechtig klein — Link/Freigabe pruefen.")
    with open(dest, 'wb') as f:
        f.write(data)
    return dest


def _resolve(env_var, local_default, tmp_name):
    url = os.environ.get(env_var, '').strip()
    if url:
        print(f"Lade Daten von URL ({env_var}) ...")
        return _download(url, os.path.join(tempfile.gettempdir(), tmp_name))
    if os.path.exists(local_default):
        print(f"Verwende lokale Datei: {local_default}")
        return local_default
    print(f"Keine Datei fuer {env_var} gefunden (uebersprungen).")
    return None


def resolve_data_path():
    return _resolve('TOTO_DATA_URL', LOCAL_GROUP, 'toto_group.xlsx')


def resolve_playoff_path():
    return _resolve('TOTO_PLAYOFF_URL', LOCAL_PLAYOFF, 'toto_playoff.xlsx')
