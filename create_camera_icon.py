"""Erstellt ein CCTV-Kamera-Piktogramm im DSGVO-Stil (blau/weiß)."""
import sys
sys.path.insert(0, r'C:\Users\localUser\AppData\Roaming\Python\Python314\site-packages')

from PIL import Image, ImageDraw

SIZE   = 400
BG     = (155, 212, 219)    # Eiszeit-Blau (#9BD4DB)
WHITE  = (255, 255, 255)
DARK   = (26,  95, 122)     # Dunkelblau

img  = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# ── Hintergrundkreis ──────────────────────────────────────────────────────────
draw.ellipse([10, 10, SIZE-10, SIZE-10], fill=BG)

# ── Kameragehäuse (Rechteck, leicht abgerundet) ───────────────────────────────
draw.rounded_rectangle([80, 140, 290, 230], radius=18, fill=WHITE)

# ── Kameraobjektiv ────────────────────────────────────────────────────────────
# Äußerer Ring
draw.ellipse([200, 130, 310, 240], fill=DARK)
# Mittlerer Ring (weiß)
draw.ellipse([215, 145, 295, 225], fill=WHITE)
# Innerer Kreis (dunkel = Linse)
draw.ellipse([230, 160, 280, 210], fill=DARK)
# Glanzpunkt
draw.ellipse([238, 168, 255, 183], fill=(200, 235, 240))

# ── Halterung / Arm ───────────────────────────────────────────────────────────
# Vertikaler Stab oben
draw.rounded_rectangle([165, 85, 195, 150], radius=8, fill=WHITE)
# Horizontaler Sockel
draw.rounded_rectangle([130, 75, 235, 100], radius=8, fill=WHITE)

# ── Signalwellen (3 Bögen rechts von der Linse) ───────────────────────────────
for i, r in enumerate([28, 42, 56]):
    cx, cy = 254, 184          # Mittelpunkt Linse
    bbox = [cx+r-4, cy-r, cx+r*2, cy+r]
    draw.arc(bbox, start=-60, end=60, fill=WHITE, width=6)

img.save("kamera_piktogramm.png", "PNG")
print("Gespeichert: kamera_piktogramm.png")
