"""
Installe les skills Claude Code pour l'outreach Trees Engineering.
Lance apres setup.py : python install_skills.py
"""
import os, shutil, sys

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
SKILLS_SRC = os.path.join(BASE_DIR, "skills")

# Destination : dossier .claude de l'utilisateur courant
home       = os.path.expanduser("~")
SKILLS_DST = os.path.join(home, ".claude", "skills")

SKILLS = ["outreach-drafts", "outreach-send", "outreach-status"]

print("=== Installation des skills Claude Code ===\n")

for skill in SKILLS:
    src = os.path.join(SKILLS_SRC, skill)
    dst = os.path.join(SKILLS_DST, skill)
    os.makedirs(dst, exist_ok=True)
    shutil.copy2(os.path.join(src, "SKILL.md"), os.path.join(dst, "SKILL.md"))
    print(f"  OK  /{skill}")

print(f"\nSkills installes dans : {SKILLS_DST}")
print("\nDisponibles dans Claude Code :")
print("  /outreach-status  → voir l'etat de la campagne")
print("  /outreach-drafts  → creer les brouillons Gmail du jour")
print("  /outreach-send    → envoyer les emails directement")
