---
name: outreach-send
description: Envoie directement les emails de relance Trees Engineering via SMTP (sans passer par les brouillons). Utilise ce skill quand l'utilisateur dit "envoie les emails", "send", "envoyer directement", "lancer les relances", ou veut que les emails partent automatiquement sans validation manuelle dans Gmail.
---

# Outreach Send — Trees Engineering

Envoie directement les emails de relance via SMTP avec 45 secondes d'attente entre chaque envoi.

## ⚠️ Important

Ce mode envoie les emails **immédiatement** — pas de validation dans Gmail. Confirme toujours avec l'utilisateur avant de lancer.

Affiche d'abord la liste des contacts qui vont recevoir un email et demande :
> "Je vais envoyer X emails. Confirme avec YES pour lancer."

Si l'utilisateur répond YES → lance le script.

## Exécution

```powershell
python "C:\Users\33661\TreesOutreach\outreach.py" --mode send
```

## Affichage du résultat

Pendant l'envoi, affiche chaque email au fur et à mesure :
- Contact + email + séquence
- Progression (ex: 2/5)
- Temps d'attente entre chaque

À la fin : résumé total envoyé + confirmation que le tracker est mis à jour.
