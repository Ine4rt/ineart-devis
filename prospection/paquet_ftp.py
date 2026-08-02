"""
IneArt - Prospection Huy : preparation du paquet a deposer sur le FTP.

Rassemble tout ce qui doit etre televerse dans dev.ineart.be, ajoute une note
explicative et un calendrier de suppression, puis produit une archive zip.

Sortie :
    livraison/dev.ineart.be/<slug>/index.html
    livraison/dev.ineart.be/A_LIRE.txt
    livraison/dev-ineart-be.zip

Usage :
    python paquet_ftp.py
"""

import datetime as dt
import json
import os
import shutil
import sys
import zipfile

import config

DOSSIER_LIVRAISON = os.path.join(config.RACINE, "livraison")
NOM_RACINE = "dev.ineart.be"


def note_explicative(manifest):
    lignes = [
        "DEPOT SUR LE FTP - dev.ineart.be",
        "=" * 40,
        "",
        "Contenu : %d maquettes, une par dossier." % len(manifest),
        "",
        "COMMENT DEPOSER",
        "  1. Connectez-vous au FTP (FileZilla ou autre).",
        "  2. Ouvrez le dossier dev.ineart.be.",
        "  3. Glissez-y tous les dossiers de societes de cette archive.",
        "",
        "  Chaque site est alors accessible a l'adresse :",
        "      https://dev.ineart.be/<nom-du-dossier>/",
        "",
        "CE QUE VOIT LE PROSPECT",
        "  Un bandeau en haut affiche le temps restant avant expiration.",
        "  Passe la date, la page se remplace par un ecran \"apercu expire\"",
        "  qui invite a vous contacter.",
        "",
        "CALENDRIER DE SUPPRESSION",
        "  Supprimez les dossiers du FTP aux dates ci-dessous. Le compte a",
        "  rebours bloque deja l'affichage, mais la suppression reelle des",
        "  fichiers reste a faire a la main (ou automatiquement si vous me",
        "  donnez les acces FTP en secrets GitHub).",
        "",
    ]

    for entree in sorted(manifest, key=lambda x: x["expire_le"]):
        lignes.append("  %s  %-34s %s" % (
            entree["expire_le_fr"], entree["slug"], entree["nom"]))

    lignes += [
        "",
        "RETRAIT IMMEDIAT",
        "  Si une societe demande le retrait, supprimez simplement son",
        "  dossier du FTP et repondez-lui que c'est fait.",
        "",
        "Genere le %s" % dt.datetime.now().strftime("%d/%m/%Y a %H:%M"),
    ]
    return "\n".join(lignes) + "\n"


def main():
    if not os.path.exists(config.FICHIER_MANIFEST):
        raise SystemExit("manifest.json introuvable : lancez generateur.py d'abord.")

    with open(config.FICHIER_MANIFEST, encoding="utf-8") as fichier:
        manifest = json.load(fichier)

    if os.path.isdir(DOSSIER_LIVRAISON):
        shutil.rmtree(DOSSIER_LIVRAISON)
    racine = os.path.join(DOSSIER_LIVRAISON, NOM_RACINE)
    os.makedirs(racine)

    copies = 0
    for entree in manifest:
        source = os.path.join(config.RACINE, entree["chemin_local"])
        if not os.path.isdir(source):
            print("  ! dossier absent : %s" % source)
            continue
        shutil.copytree(source, os.path.join(racine, entree["slug"]))
        copies += 1

    with open(os.path.join(racine, "A_LIRE.txt"), "w", encoding="utf-8") as fichier:
        fichier.write(note_explicative(manifest))

    archive = os.path.join(DOSSIER_LIVRAISON, "dev-ineart-be.zip")
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zip_sortie:
        for dossier, _, fichiers in os.walk(racine):
            for nom in fichiers:
                chemin = os.path.join(dossier, nom)
                zip_sortie.write(chemin, os.path.relpath(chemin, DOSSIER_LIVRAISON))

    poids = os.path.getsize(archive) / 1024
    print("%d maquettes copiees dans %s" % (copies, racine))
    print("Archive : %s (%.0f ko)" % (archive, poids))
    return 0


if __name__ == "__main__":
    sys.exit(main())
