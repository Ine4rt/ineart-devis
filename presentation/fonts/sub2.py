from fontTools import subset
from fontTools.varLib.instancer import instantiateVariableFont
from fontTools.ttLib import TTFont
import base64, io

CHARS = (
    "".join(chr(c) for c in range(0x20, 0x7F))
    + "àâäçéèêëîïôöùûüÿœæÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸŒÆ"
    + "«»“”‘’–—…€°·•→↗×"
    + "   "
)

def build(out, axes, name):
    font = TTFont("archivo.ttf")
    font = instantiateVariableFont(font, axes, updateFontNames=False, overlap=True)
    o = subset.Options()
    o.layout_features = ["kern", "liga", "calt", "ccmp", "locl", "mark", "mkmk", "tnum"]
    o.name_IDs = ["*"]; o.name_legacy = True; o.notdef_outline = True
    o.drop_tables += ["DSIG"]
    s = subset.Subsetter(options=o); s.populate(text=CHARS); s.subset(font)
    font.flavor = "woff2"
    buf = io.BytesIO(); font.save(buf); data = buf.getvalue()
    open(out, "wb").write(data)
    b64 = base64.b64encode(data).decode()
    open(out + ".b64", "w").write(b64)
    print(f"{name}: {len(data)//1024} Ko -> {len(b64)//1024} Ko base64")

# Plaques de signalétique : large et lourd.
build("archivo-exp.woff2", {"wdth": 118, "wght": (600, 700, 800)}, "Archivo Expanded")
# Texte courant : largeur normale.
build("archivo.woff2", {"wdth": 100, "wght": (400, 500, 700)}, "Archivo")
