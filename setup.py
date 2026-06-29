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

    # Auto-detect Excel file in common locations
    import glob
    home = os.path.expanduser("~")
    search_dirs = [
        os.path.join(home, "Downloads"),
        os.path.join(home, "Desktop"),
        os.path.join(home, "Documents"),
        home,
    ]
    found = []
    for d in search_dirs:
        found += glob.glob(os.path.join(d, "*.xlsx"))

    excel = ""
    trees_files = [f for f in found if "trees" in f.lower() or "prospect" in f.lower() or "contact" in f.lower()]

    if trees_files:
        print(f"\n   Fichier Excel detecte automatiquement :")
        for i, f in enumerate(trees_files):
            print(f"   [{i+1}] {f}")
        if len(trees_files) == 1:
            confirm = input(f"\n   Utiliser ce fichier ? (O/n) : ").strip().lower()
            excel = trees_files[0] if confirm != "n" else ""
        else:
            choice = input(f"\n   Choisir le numero (1-{len(trees_files)}) : ").strip()
            try: excel = trees_files[int(choice)-1]
            except: excel = ""

    if not excel:
        print("\n   Fichier Excel non detecte automatiquement.")
        print("   Ouvre ton fichier Excel, regarde en haut de la fenetre :")
        print("   ex: C:\\Users\\Prenom\\Downloads\\contacts.xlsx")
        print("   Copie ce chemin complet et colle-le ici.")
        excel = input("\n   Chemin du fichier Excel : ").strip().strip('"')

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
