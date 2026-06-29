---
name: outreach-drafts
description: Crée les brouillons Gmail de relance Trees Engineering pour aujourd'hui. Utilise ce skill dès que l'utilisateur mentionne "drafts", "brouillons", "relances", "outreach", "follow-up", "contacts à relancer", ou demande de préparer les emails du jour. Lance le script automatiquement sans poser de questions.
---

# Outreach Drafts — Trees Engineering

Crée les brouillons Gmail pour tous les contacts en fenêtre de relance aujourd'hui (J+5→7, J+12→14, J+20→25).

## Ce que fait ce skill

1. Lit le tracker Excel (`Trees_Engineering_Expanded_Prospects`)
2. Calcule les contacts à relancer selon la date du jour
3. Génère les emails HTML brandés (Version A/B/C selon le profil)
4. Crée les brouillons dans Gmail via IMAP
5. Met à jour le tracker avec la date de création

## Exécution

Lance directement :

```powershell
python "C:\Users\33661\create_drafts.py"
```

## Affichage du résultat

Après exécution, présente un tableau récapitulatif :

| Contact | Entreprise | Séquence | Email |
|---|---|---|---|
| ... | ... | ... | ... |

Indique le nombre total de brouillons créés et rappelle à l'utilisateur d'aller dans Gmail → Drafts pour envoyer.

Si 0 contacts → affiche "Aucune relance à envoyer aujourd'hui." et liste les prochaines échéances.
