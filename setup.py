"""
Trees Engineering — Outreach Automation
Setup: installe les dependances, genere les logos, configure le .env
"""
import subprocess, sys, os, shutil

def run(cmd):
    subprocess.check_call([sys.executable, "-m"] + cmd)

print("=== Trees Engineering Outreach Setup ===\n")

# 1. Install dependencies
print("1. Installation des dependances...")
run(["pip", "install", "pandas", "openpyxl", "python-dotenv", "Pillow"])
print("   OK\n")

# 2. Configure .env
env_path = os.path.join(os.path.dirname(__file__), ".env")
if not os.path.exists(env_path):
    print("2. Configuration Gmail...")
    gmail = input("   Votre adresse Gmail : ").strip()
    pwd   = input("   Votre App Password (16 chars) : ").strip()
    excel = input("   Chemin vers le fichier Excel tracker : ").strip()
    with open(env_path, "w") as f:
        f.write(f"GMAIL_ADDRESS={gmail}\n")
        f.write(f"GMAIL_APP_PASSWORD={pwd}\n")
        f.write(f"EXCEL_PATH={excel}\n")
    print("   .env cree\n")
else:
    print("2. .env deja configure\n")

# 3. Generate logos
print("3. Generation des logos PNG...")
run(["pip", "install", "Pillow"])
exec(open(os.path.join(os.path.dirname(__file__), "generate_logos.py")).read())
print("   Logos OK\n")

print("=== Setup termine ! ===")
print("Pour lancer manuellement : python outreach.py --mode drafts")
print("Pour envoyer directement : python outreach.py --mode send")
print("La tache planifiee Windows tourne automatiquement chaque matin a 8h.")
