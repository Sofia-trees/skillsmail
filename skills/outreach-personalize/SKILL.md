---
name: outreach-personalize
description: Aide a personnaliser les emails de relance Trees Engineering avant de créer les drafts. Pour chaque contact qui entre en fenetre de relance, genere une ligne d'accroche personnalisee basee sur le titre, l'entreprise, le secteur et les notes disponibles. Met a jour la colonne "Personalization Note" dans le tracker Excel. Utilise ce skill quand l'utilisateur dit "personnaliser", "personnalise les mails", "ajouter une touche perso", "preparer les relances", ou avant de lancer /outreach-drafts.
---

# Outreach Personalize

Genere des lignes d'accroche personnalisees pour chaque contact en fenetre de relance, puis met a jour le tracker.

## Execution

Lance ce script Python inline :

```python
import pandas as pd, re, sys
from datetime import date
sys.path.insert(0, r'C:\Users\33661\AppData\Local\Python\pythoncore-3.14-64\Lib\site-packages')

TODAY = date.today()
EXCEL = r'C:\Users\33661\Downloads\Trees_Engineering_Expanded_Prospects (6).xlsx'

df = pd.read_excel(EXCEL, sheet_name='All Contacts', header=1, dtype=str)
df.columns = df.columns.str.strip()

def parse_date(status):
    if not isinstance(status, str): return None
    m = re.search(r'\((\d{1,2}/\d{1,2})\)', status)
    if not m: return None
    try:
        d, mo = m.group(1).split('/')
        return date(date.today().year, int(mo), int(d))
    except: return None

targets = []
for idx, row in df.iterrows():
    status = str(row.get('Status', '') or '')
    if 'Contacted' not in status or 'Whatsapp' in status: continue
    cd = parse_date(status)
    if not cd: continue
    days = (TODAY - cd).days
    if not (5 <= days <= 25): continue
    seq = 'Follow-up 1' if 5 <= days <= 7 else 'Follow-up 2' if 12 <= days <= 14 else 'Final Email'
    existing_note = str(row.get('Personalization Note', '') or '').strip()
    targets.append({
        'idx': idx,
        'name': str(row.get('Full Name', '')),
        'title': str(row.get('Title', '')),
        'company': str(row.get('Company', '')),
        'days': days,
        'seq': seq,
        'note': existing_note if existing_note not in ['nan', 'None', ''] else ''
    })

print(f"Contacts en fenetre de relance: {len(targets)}\n")
for t in targets:
    flag = "[note existante]" if t['note'] else "[a personnaliser]"
    print(f"- {t['name']} | {t['title']} | {t['company']} | {t['seq']} | {flag}")
    if t['note']:
        print(f"  Note actuelle: {t['note']}")
```

## Presentation et action

Apres avoir execute le script :

1. Affiche la liste des contacts en fenetre avec leur titre et entreprise
2. Pour chaque contact SANS note existante, propose une ligne d'accroche personnalisee basee sur :
   - Le titre exact (Head of..., Director..., AVP...)
   - Le secteur (Oil & Gas, Healthcare, Infrastructure...)
   - L'entreprise (PETRONAS, YTL, IHH...)
   - Le contexte metier specifique a leur role

**Format des lignes proposees :**
Courtes (1 phrase max), naturelles, feminines et directes. Pas de formules generiques.

Exemples par role :
- Head Well Engineering PETRONAS : "J'ai suivi de pres les projets d'exploration PETRONAS cette annee, votre role a l'interface technique est exactement la ou nos ingenieurs font la difference"
- AVP IHH Healthcare : "Le developpement accelere d'IHH en Asie du Sud-Est cree des besoins en competences techniques qui evoluent vite"
- Executive Director YTL : "YTL a des projets ambitieux en infrastructure, et trouver les bons profils techniques au bon moment est souvent ce qui fait tenir les delais"
- Programme Director : "En tant que Programme Director, vous savez mieux que quiconque ce que ca coute d'avoir le mauvais ingenieur au mauvais moment"

3. Demande confirmation : "Je valide ces lignes et je mets a jour le tracker ?"
4. Si oui, ecris les lignes validees dans la colonne "Personalization Note" du tracker via openpyxl
5. Dis : "Tracker mis a jour. Lance /outreach-drafts pour creer les brouillons."

## Ton

- Feminin et direct, pas corporatif
- Montre que tu connais leur industrie et leur role specifique
- Une seule phrase, pas de "j'espere que vous allez bien"
- En anglais (les emails sont en anglais)
