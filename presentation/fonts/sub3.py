from fontTools import subset
from fontTools.varLib.instancer import instantiateVariableFont
from fontTools.ttLib import TTFont
import base64, io

CHARS = ("".join(chr(c) for c in range(0x20, 0x7F))
    + "àâäçéèêëîïôöùûüÿœæÀÂÄÇÉÈÊËÎÏÔÖÙÛÜŸŒÆ"
    + "«»“”‘’–—…€°·•→↗×" + "   ")

font = TTFont("archivo.ttf")
# Une seule graisse variable, de la légèreté à la demi-grasse. Le 800 étendu
# de la version précédente est exactement ce qui rendait la page lourde.
font = instantiateVariableFont(font, {"wdth": 100, "wght": (300, 400, 620)},
                               updateFontNames=False, overlap=True)
o = subset.Options()
o.layout_features = ["kern","liga","calt","ccmp","locl","mark","mkmk","tnum"]
o.name_IDs = ["*"]; o.name_legacy = True; o.notdef_outline = True
o.drop_tables += ["DSIG"]
s = subset.Subsetter(options=o); s.populate(text=CHARS); s.subset(font)
font.flavor = "woff2"
buf = io.BytesIO(); font.save(buf); data = buf.getvalue()
open("archivo.woff2","wb").write(data)
b64 = base64.b64encode(data).decode()
open("archivo.woff2.b64","w").write(b64)
print(f"Archivo 300–620 : {len(data)//1024} Ko -> {len(b64)//1024} Ko base64")
