---
name: outreach-status
description: Affiche le statut complet de la campagne outreach Trees Engineering — combien de contacts contactés, en attente, relancés, et qui est à relancer dans les 7 prochains jours. Utilise ce skill quand l'utilisateur demande "statut", "status", "combien de contacts", "qui relancer", "pipeline", "vue d'ensemble", "rapport", ou veut savoir où en est la campagne.
---

# Outreach Status — Trees Engineering

Donne une vue complète de la campagne : état actuel + prochaines échéances.

## Exécution

Lance ce script Python inline :

```python
import pandas as pd
from datetime import date
import sys
sys.path.insert(0, r'C:\Users\33661\AppData\Local\Python\pythoncore-3.14-64\Lib\site-packages')

TODAY = date.today()
EXCEL = r'C:\Users\33661\Downloads\Trees_Engineering_Expanded_Prospects (6).xlsx'

df = pd.read_excel(EXCEL, sheet_name='All Contacts', header=1, dtype=str)
df.columns = df.columns.str.strip()

import re

def parse_date(status):
    if not isinstance(status, str): return None
    m = re.search(r'\((\d{1,2}/\d{1,2})\)', status)
    if not m: return None
    try:
        d, mo = m.group(1).split('/')
        return date(2025, int(mo), int(d))
    except: return None

def days_since(s):
    d = parse_date(s)
    return (TODAY - d).days if d else None

def get_window(days):
    if days is None: return None
    if 5  <= days <= 7:  return 'Follow-up 1 DUE'
    if 12 <= days <= 14: return 'Follow-up 2 DUE'
    if 20 <= days <= 25: return 'Final Email DUE'
    if days < 5:         return f'J+{days} (pas encore)'
    return 'Hors fenetre'

# Stats globales
total = len(df)
not_started = df['Status'].isna().sum() + (df['Status'] == 'Not Started').sum()
contacted = df['Status'].str.contains('Contacted', na=False).sum()
draft = df['Status'].str.contains('Draft created', na=False).sum() if 'Follow-up Status' in df.columns else 0
sent = df['Status'].str.contains('Sent', na=False).sum() if 'Follow-up Status' in df.columns else 0

print(f"=== STATUT CAMPAGNE TREES ENGINEERING — {TODAY} ===\n")
print(f"Total contacts  : {total}")
print(f"Non contactes   : {not_started}")
print(f"Contactes       : {contacted}")
print(f"Drafts crees    : {draft}")
print(f"Envoyes         : {sent}")

# Contacts avec relances dues ou a venir (7 prochains jours)
print(f"\n=== RELANCES (actives + 7 prochains jours) ===")
upcoming = []
for _, row in df.iterrows():
    status = row.get('Status', '')
    if not isinstance(status, str) or 'Contacted' not in status: continue
    if 'Whatsapp' in status: continue
    cd = parse_date(status)
    if not cd: continue
    days = (TODAY - cd).days
    # Show due now or due in next 7 days
    future_days = days - 25  # days past final window
    if days <= 25 + 7:
        window = get_window(days)
        email = str(row.get('Work Email') or row.get('Personal Email') or '—')
        upcoming.append({
            'Name': row.get('Full Name', ''),
            'Company': row.get('Company', ''),
            'J+': days,
            'Statut': window,
            'Email': 'OK' if email not in ['nan', '—', ''] else 'MANQUANT'
        })

upcoming.sort(key=lambda x: x['J+'], reverse=True)
if upcoming:
    print(f"{'Nom':<25} {'Entreprise':<20} {'J+':<5} {'Statut':<22} {'Email'}")
    print('-' * 85)
    for u in upcoming:
        print(f"{u['Name']:<25} {u['Company']:<20} {u['J+']:<5} {u['Statut']:<22} {u['Email']}")
else:
    print("Aucune relance en cours ou a venir.")
```

## Présentation du résultat

Après avoir exécuté le script, présente les résultats de façon visuelle :

1. **Tableau de bord** — métriques globales en une ligne
2. **Tableau des relances** — trié par urgence (DUE en premier)
3. **Recommandation** — que faire maintenant ? (ex: "3 Follow-up 1 à créer → lance `/outreach-drafts`")
