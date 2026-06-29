---
name: outreach-bounce
description: Scanne Gmail pour les emails non delivres (bounces automatiques) et met a jour le tracker Excel en remplacant "Contacted" par "LinkedIn request sent + mail (undelivered)". Utilise ce skill quand l'utilisateur dit "j'ai des bounces", "mails non delivres", "undelivered", "messages automatiques", "enleve contacted", ou veut signaler qu'un email n'a pas abouti.
---

# Outreach Bounce Handler

Detecte les emails non delivres dans Gmail et met a jour le statut dans le tracker.

## Execution

Lance ce script Python inline :

```python
import imaplib, email as email_lib, re, sys, openpyxl
from datetime import date
from dotenv import load_dotenv
import os
sys.path.insert(0, r'C:\Users\33661\AppData\Local\Python\pythoncore-3.14-64\Lib\site-packages')

load_dotenv(r'C:\Users\33661\.env')
GMAIL = os.getenv('GMAIL_ADDRESS')
PWD   = os.getenv('GMAIL_APP_PASSWORD')
EXCEL = r'C:\Users\33661\Downloads\Trees_Engineering_Expanded_Prospects (6).xlsx'

# Connect to Gmail
imap = imaplib.IMAP4_SSL('imap.gmail.com')
imap.login(GMAIL, PWD)
imap.select('INBOX')

# Search for bounce/undelivered notifications
bounced_emails = set()
search_terms = [
    'FROM "mailer-daemon"',
    'FROM "postmaster"',
    'SUBJECT "Undelivered Mail"',
    'SUBJECT "Delivery Status Notification"',
    'SUBJECT "Mail delivery failed"',
    'SUBJECT "Delivery Failure"',
    'SUBJECT "returned mail"',
]

for term in search_terms:
    _, data = imap.search(None, term)
    if data[0]:
        for num in data[0].split():
            _, msg_data = imap.fetch(num, '(RFC822)')
            raw = msg_data[0][1]
            msg = email_lib.message_from_bytes(raw)
            # Extract original recipient from bounce body
            body = ''
            for part in msg.walk():
                ct = part.get_content_type()
                if ct in ('text/plain', 'message/delivery-status'):
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            body += payload.decode('utf-8', errors='ignore')
                    except: pass
            # Look for email addresses in the bounce
            found = re.findall(r'[\w.\-+]+@[\w.\-]+\.[a-z]{2,}', body)
            for addr in found:
                addr = addr.lower().strip()
                if GMAIL.lower() not in addr and 'mailer' not in addr and 'postmaster' not in addr:
                    bounced_emails.add(addr)

imap.logout()

print(f"Adresses bouncees trouvees: {len(bounced_emails)}")
for e in bounced_emails:
    print(f"  - {e}")

# Update tracker
if bounced_emails:
    wb = openpyxl.load_workbook(EXCEL)
    ws = wb['All Contacts']
    headers = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    
    work_email_col   = headers.index('Work Email') + 1 if 'Work Email' in headers else None
    status_col       = headers.index('Status') + 1 if 'Status' in headers else None
    name_col         = headers.index('Full Name') + 1 if 'Full Name' in headers else None
    
    updated = []
    for row in range(3, ws.max_row + 1):
        we = ws.cell(row=row, column=work_email_col).value if work_email_col else None
        if we and str(we).strip().lower() in bounced_emails:
            current_status = str(ws.cell(row=row, column=status_col).value or '')
            if 'Contacted' in current_status and 'undelivered' not in current_status:
                # Extract date from current status if present
                m = re.search(r'\((\d{1,2}/\d{1,2})\)', current_status)
                date_str = m.group(1) if m else date.today().strftime('%d/%m')
                ws.cell(row=row, column=status_col).value = f"LinkedIn request sent + mail (undelivered) ({date_str})"
                name = ws.cell(row=row, column=name_col).value
                updated.append(str(name))
    
    wb.save(EXCEL)
    if updated:
        print(f"\nTracker mis a jour pour:")
        for n in updated:
            print(f"  - {n} -> LinkedIn request sent + mail (undelivered)")
    else:
        print("\nAucun contact correspondant trouve dans le tracker.")
```

## Presentation

Apres execution :
- Affiche les adresses detectees comme bouncees
- Confirme quels contacts ont ete mis a jour dans le tracker
- Si aucun bounce detecte, indique "Aucun email non delivre trouve dans ta boite."
- Rappelle que ces contacts ne recevront plus de relances automatiques (statut ne contient plus "Contacted" seul)
