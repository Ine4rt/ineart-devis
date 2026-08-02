"""
IneArt - Prospection Huy : suppression reelle des maquettes expirees.

C'est la couche qui rend la limite de 8 jours effective : le compte a rebours
et l'ecran de fin sont cote navigateur, ici on supprime les fichiers du FTP.

A lancer une fois par jour (workflow nettoyage_demos.yml).

Usage :
    python nettoyage.py             # supprime tout ce qui a depasse la date
    python nettoyage.py --dry-run   # liste sans rien supprimer
    python nettoyage.py --slug xxx  # retrait immediat d'une societe
"""

import argparse
import datetime as dt
import ftplib
import json
import os
import sys

import config
from ftp_deploy import connexion


def supprimer_dossier(session, chemin):
    """Supprime recursivement un dossier distant."""
    try:
        entrees = session.nlst(chemin)
    except ftplib.error_perm as erreur:
        if str(erreur).startswith("550"):
            return False
        raise

    for entree in entrees:
        nom = entree.split("/")[-1]
        if nom in (".", ".."):
            continue
        complet = entree if entree.startswith("/") else "%s/%s" % (chemin, nom)
        try:
            session.delete(complet)
        except ftplib.error_perm:
            supprimer_dossier(session, complet)
    try:
        session.rmd(chemin)
    except ftplib.error_perm as erreur:
        print("    ! rmd %s : %s" % (chemin, erreur))
        return False
    return True


def main():
    analyseur = argparse.ArgumentParser(description="Suppression des maquettes expirees")
    analyseur.add_argument("--dry-run", action="store_true")
    analyseur.add_argument("--slug", help="retirer une societe immediatement")
    arguments = analyseur.parse_args()

    if not os.path.exists(config.FICHIER_MANIFEST):
        print("Aucun manifest : rien a nettoyer.")
        return 0

    with open(config.FICHIER_MANIFEST, encoding="utf-8") as fichier:
        manifest = json.load(fichier)

    maintenant = dt.datetime.now(dt.timezone.utc)
    expirees = []
    for entree in manifest:
        if entree.get("supprime"):
            continue
        if arguments.slug:
            if entree["slug"] == arguments.slug:
                expirees.append(entree)
            continue
        if not entree.get("deploye"):
            continue
        date_fin = dt.datetime.fromisoformat(entree["expire_le"].replace("Z", "+00:00"))
        if maintenant >= date_fin:
            expirees.append(entree)

    if not expirees:
        print("Aucune maquette expiree (%d suivies)." % len(manifest))
        return 0

    print("%d maquette(s) a retirer :" % len(expirees))
    for entree in expirees:
        print("  - %s (expiree le %s)" % (entree["nom"], entree.get("expire_le_fr", "?")))

    if arguments.dry_run:
        print("Simulation : rien n'a ete supprime.")
        return 0

    racine = config.FTP["dossier"].rstrip("/")
    session = connexion()
    supprimees = 0
    try:
        for entree in expirees:
            chemin = "%s/%s" % (racine, entree["slug"])
            print("  suppression de %s" % chemin)
            if supprimer_dossier(session, chemin):
                supprimees += 1
            entree["supprime"] = True
            entree["supprime_le"] = maintenant.isoformat().replace("+00:00", "Z")
    finally:
        try:
            session.quit()
        except Exception:  # noqa: BLE001
            session.close()

    with open(config.FICHIER_MANIFEST, "w", encoding="utf-8") as fichier:
        json.dump(manifest, fichier, ensure_ascii=False, indent=2)

    print("\n%d maquette(s) retiree(s) du FTP." % supprimees)
    return 0


if __name__ == "__main__":
    sys.exit(main())
