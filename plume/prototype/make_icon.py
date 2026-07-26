#!/usr/bin/env python3
"""Icône Plume pour l'écran d'accueil iOS : la lune sur l'horizon, ciel
de l'Heure Bleue (docs/02-DESIGN.md). Coins carrés — iOS arrondit lui-même.
"""
import base64
import math
from io import BytesIO

from PIL import Image, ImageDraw, ImageFilter

SIZE = 1024


def make_icon() -> Image.Image:
    img = Image.new("RGB", (SIZE, SIZE), "#0B0E1A")
    draw = ImageDraw.Draw(img)

    # Ciel : dégradé nuit (haut plus profond, bas plus doux).
    top = (8, 11, 26)
    bottom = (19, 23, 51)
    for y in range(SIZE):
        t = y / SIZE
        r = int(top[0] + (bottom[0] - top[0]) * t)
        g = int(top[1] + (bottom[1] - top[1]) * t)
        b = int(top[2] + (bottom[2] - top[2]) * t)
        draw.line([(0, y), (SIZE, y)], fill=(r, g, b))

    # Étoiles discrètes.
    import random
    rng = random.Random(7)
    star_layer = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    sd = ImageDraw.Draw(star_layer)
    for _ in range(70):
        x, y = rng.uniform(0, SIZE), rng.uniform(0, SIZE * 0.62)
        r = rng.uniform(1.5, 4)
        a = rng.randint(60, 160)
        sd.ellipse([x - r, y - r, x + r, y + r], fill=(244, 241, 232, a))
    img.paste(Image.alpha_composite(img.convert("RGBA"), star_layer).convert("RGB"), (0, 0))

    # Halo autour de la lune (lueur diffuse, glow gold).
    glow = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    cx, cy, moon_r = SIZE * 0.5, SIZE * 0.42, SIZE * 0.17
    for i in range(6, 0, -1):
        rr = moon_r * (1 + i * 0.18)
        alpha = int(30 / i)
        gd.ellipse([cx - rr, cy - rr, cx + rr, cy + rr],
                   fill=(232, 194, 135, alpha))
    glow = glow.filter(ImageFilter.GaussianBlur(SIZE * 0.03))
    img = Image.alpha_composite(img.convert("RGBA"), glow).convert("RGB")
    draw = ImageDraw.Draw(img)

    # La lune : dégradé champagne, légèrement mouchetée.
    moon_layer = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    md = ImageDraw.Draw(moon_layer)
    moon_cy = cy - moon_r * 0.35
    steps = 60
    for i in range(steps):
        t = i / steps
        rr = moon_r * (1 - t * 0.02)
        col = (
            int(245 - (245 - 201) * t),
            int(235 - (235 - 160) * t),
            int(216 - (216 - 94) * t),
        )
        md.ellipse([cx - rr, moon_cy - rr, cx + rr, moon_cy + rr],
                   fill=col + (255,))
    # Petits cratères doux.
    for dx, dy, rr, shade in (
        (-0.30, -0.15, 0.09, -18), (0.22, 0.05, 0.13, -22),
        (-0.05, 0.25, 0.07, -14), (0.30, -0.28, 0.05, -12),
    ):
        px, py = cx + dx * moon_r, moon_cy + dy * moon_r
        prr = rr * moon_r
        base = (201, 160, 94)
        col = tuple(max(0, c + shade) for c in base) + (140,)
        md.ellipse([px - prr, py - prr, px + prr, py + prr], fill=col)
    img = Image.alpha_composite(img.convert("RGBA"), moon_layer).convert("RGB")
    draw = ImageDraw.Draw(img)

    # Horizon : collines en silhouette (signature visuelle du design).
    hill_layer = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    hd = ImageDraw.Draw(hill_layer)

    def hill(base_y, amp, phase, color, alpha):
        pts = [(0, SIZE)]
        for x in range(0, SIZE + 1, 8):
            y = base_y + amp * math.sin(x / SIZE * math.pi * 1.4 + phase)
            pts.append((x, y))
        pts.append((SIZE, SIZE))
        hd.polygon(pts, fill=color + (alpha,))

    hill(SIZE * 0.78, SIZE * 0.05, 0.6, (16, 20, 42), 255)
    hill(SIZE * 0.86, SIZE * 0.06, 2.1, (12, 15, 32), 255)
    img = Image.alpha_composite(img.convert("RGBA"), hill_layer).convert("RGB")

    return img


def to_data_uri(img: Image.Image) -> str:
    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


if __name__ == "__main__":
    icon = make_icon()
    out_dir = "/tmp/claude-0/-home-user-ineart-devis/5d9bdc64-404a-559a-8693-2d05ab9c8456/scratchpad"
    icon.save(f"{out_dir}/icon-1024.png")
    for size in (180, 192, 512):
        icon.resize((size, size), Image.LANCZOS).save(f"{out_dir}/icon-{size}.png")
    print("icônes générées")
