from fontTools import subset
from fontTools.varLib.instancer import instantiateVariableFont
from fontTools.ttLib import TTFont
import base64, io

# Jeu de glyphes : latin de base + accents français + ponctuation typographique.
CHARS = (
    "".join(chr(c) for c in range(0x20, 0x7F))
    + "àâäçéèêëîïôöùûüÿœæÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸŒÆ"
    + "«»“”‘’‚„–—…€°·•→←↑↓×≥≤±§№™©®"
    + "0123456789"
    + "   "  # espaces insécables (typographie française)
)

def build(src, out, axes, name):
    font = TTFont(src)
    if axes:
        font = instantiateVariableFont(font, axes, updateFontNames=False, overlap=True)
    opts = subset.Options()
    opts.layout_features = ["kern", "liga", "calt", "ccmp", "locl", "mark", "mkmk", "tnum", "onum"]
    opts.desubroutinize = False
    opts.name_IDs = ["*"]
    opts.name_legacy = True
    opts.notdef_outline = True
    opts.recalc_bounds = True
    opts.drop_tables += ["DSIG"]
    s = subset.Subsetter(options=opts)
    s.populate(text=CHARS)
    s.subset(font)
    font.flavor = "woff2"
    buf = io.BytesIO()
    font.save(buf)
    data = buf.getvalue()
    open(out, "wb").write(data)
    b64 = base64.b64encode(data).decode()
    open(out + ".b64", "w").write(b64)
    print(f"{name}: {len(data)//1024} Ko woff2 -> {len(b64)//1024} Ko base64")

# Fraunces : on fige l'optique et le caractère (SOFT/WONK), on garde l'axe de graisse.
build("fraunces.ttf", "fraunces.woff2",
      {"opsz": 144, "SOFT": 40, "WONK": 1, "wght": (500, 700, 800)}, "Fraunces")

# Karla : axe de graisse conservé de 400 à 700.
build("karla.ttf", "karla.woff2", {"wght": (400, 450, 700)}, "Karla")
