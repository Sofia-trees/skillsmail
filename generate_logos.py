"""Genere les 9 logos PNG brandes pour les templates email."""
from PIL import Image, ImageDraw, ImageFont
import base64, os, re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE_DIR, "logos")
os.makedirs(OUT, exist_ok=True)

W, H = 240, 60

def find_font(size, bold=True):
    names = ["arialbd.ttf", "arial.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans.ttf"]
    dirs  = [r"C:\Windows\Fonts", "/usr/share/fonts/truetype/dejavu", "/usr/share/fonts"]
    for name in names:
        for d in dirs:
            fp = os.path.join(d, name)
            if os.path.exists(fp):
                try: return ImageFont.truetype(fp, size)
                except: pass
    return ImageFont.load_default()

def hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0,2,4))

def centered_text(draw, text, color, size, iw, ih):
    font = find_font(size)
    bb = draw.textbbox((0,0), text, font=font)
    x = (iw - (bb[2]-bb[0])) // 2
    y = (ih - (bb[3]-bb[1])) // 2 - bb[1]
    draw.text((x,y), text, fill=hex_rgb(color), font=font)

def save_png(img, name):
    bg = Image.new("RGB", img.size, (255,255,255))
    if img.mode == "RGBA":
        bg.paste(img, mask=img.split()[3])
    else:
        bg.paste(img)
    path = os.path.join(OUT, f"{name}.png")
    bg.save(path, "PNG")
    return path

def png_to_data_uri(path):
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")

# Simple text logos
SIMPLE = [
    ("totalenergies",   "TotalEnergies",   "#D4002A", 22),
    ("technipenergies", "Technip Energies", "#003B71", 18),
    ("nov",             "NOV",              "#002D62", 32),
    ("worley",          "Worley",           "#E8410A", 28),
    ("petrofac",        "Petrofac",         "#004B87", 22),
    ("pttep",           "PTTEP",            "#007DC5", 26),
    ("woodplc",         "wood.",            "#1F3D4E", 30),
]

paths = {}

for name, text, color, size in SIMPLE:
    img = Image.new("RGBA", (W,H), (0,0,0,0))
    centered_text(ImageDraw.Draw(img), text, color, size, W, H)
    paths[name] = save_png(img, name)

# Subsea7 — two colors
img = Image.new("RGBA", (W,H), (0,0,0,0))
draw = ImageDraw.Draw(img)
f1, f2 = find_font(26), find_font(32)
b1 = draw.textbbox((0,0), "subsea", font=f1)
b2 = draw.textbbox((0,0), "7", font=f2)
tw = (b1[2]-b1[0]) + 4 + (b2[2]-b2[0])
x0 = (W - tw) // 2
draw.text((x0, (H-(b1[3]-b1[1]))//2 - b1[1]), "subsea", fill=hex_rgb("#1A2B6B"), font=f1)
draw.text((x0+(b1[2]-b1[0])+4, (H-(b2[3]-b2[1]))//2 - b2[1]), "7", fill=hex_rgb("#CC0000"), font=f2)
paths["subsea7"] = save_png(img, "subsea7")

# Bureau Veritas — BV badge
img = Image.new("RGBA", (W,H), (0,0,0,0))
draw = ImageDraw.Draw(img)
rw, rh = 80, 40
rx, ry = (W-rw)//2, (H-rh)//2
draw.rectangle([rx, ry, rx+rw, ry+rh], fill=hex_rgb("#003087"))
f = find_font(20)
b = draw.textbbox((0,0), "BV", font=f)
draw.text((rx+(rw-(b[2]-b[0]))//2, ry+(rh-(b[3]-b[1]))//2 - b[1]), "BV", fill=(255,255,255), font=f)
paths["bureauveritas"] = save_png(img, "bureauveritas")

DOMAIN_MAP = {
    "totalenergies.com":   paths["totalenergies"],
    "technipenergies.com": paths["technipenergies"],
    "nov.com":             paths["nov"],
    "worley.com":          paths["worley"],
    "petrofac.com":        paths["petrofac"],
    "pttep.com":           paths["pttep"],
    "bureauveritas.com":   paths["bureauveritas"],
    "woodplc.com":         paths["woodplc"],
    "subsea7.com":         paths["subsea7"],
}

COMPANY_NAMES = {
    "totalenergies.com":   "TotalEnergies",
    "technipenergies.com": "Technip Energies",
    "nov.com":             "NOV",
    "worley.com":          "Worley",
    "petrofac.com":        "Petrofac",
    "pttep.com":           "PTTEP",
    "bureauveritas.com":   "Bureau Veritas",
    "woodplc.com":         "Wood",
    "subsea7.com":         "Subsea7",
}

TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
for tpl in ["template_A.html", "template_B.html", "template_C.html"]:
    path = os.path.join(TEMPLATES_DIR, tpl)
    if not os.path.exists(path):
        continue
    with open(path, "r", encoding="utf-8") as f:
        html = f.read()
    for domain, png_path in DOMAIN_MAP.items():
        uri = png_to_data_uri(png_path)
        cname = COMPANY_NAMES[domain]
        pattern = r'(<img\s[^>]*?)src="[^"]*"([^>]*?onerror="[^"]*' + re.escape(cname) + r'[^"]*"[^>]*?>)'
        def make_rep(u):
            def rep(m): return m.group(1) + f'src="{u}"' + m.group(2)
            return rep
        new_html, n = re.subn(pattern, make_rep(uri), html, flags=re.DOTALL)
        if n: html = new_html
        else:
            html, _ = re.subn(r'src="data:[^"]*"([^>]*?' + re.escape(cname) + r')', f'src="{uri}"\\1', html, flags=re.DOTALL)
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    print(f"Logos generes dans {OUT}/")
    for k,v in paths.items(): print(f"  {k}.png")
