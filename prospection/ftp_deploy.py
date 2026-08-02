"""
IneArt - Prospection Huy : mise en ligne des maquettes sur le FTP.

Envoie sites/<slug>/ vers <FTP_DIR>/<slug>/ et met a jour data/manifest.json
(champ deploye + url reelle).

Variables d'environnement attendues (memes secrets que le projet ineweb) :
    FTP_HOST, FTP_USER, FTP_PASS
    FTP_DIR         dossier distant racine des demos   (defaut /demos)
    DEMO_BASE_URL   URL publique correspondante        (defaut https://www.ineart.be/demos)
    FTP_TLS         1 pour FTPS explicite (defaut), 0 pour FTP simple

Usage :
    python ftp_deploy.py                 # envoie tout ce qui n'est pas deploye
    python ftp_deploy.py --tout          # renvoie tout, meme deja deploye
    python ftp_deploy.py --slug xxx      # une seule societe
    python ftp_deploy.py --dry-run       # simulation, aucune connexion
"""

import argparse
import ftplib
import json
import os
import ssl
import sys

import config


def connexion():
    """Ouvre une session FTP(S). Retombe sur du FTP simple si TLS refuse."""
    if not config.FTP["host"]:
        raise SystemExit("FTP_HOST absent : configurez les secrets FTP.")

    if config.FTP["tls"]:
        try:
            contexte = ssl.create_default_context()
            session = ftplib.FTP_TLS(context=contexte)
            session.connect(config.FTP["host"], 21, timeout=45)
            session.login(config.FTP["user"], config.FTP["password"])
            session.prot_p()
            session.set_pasv(True)
            print("Connexion FTPS etablie sur %s" % config.FTP["host"])
            return session
        except Exception as erreur:  # noqa: BLE001
            print("! FTPS indisponible (%s), bascule en FTP simple." % erreur)

    session = ftplib.FTP()
    session.connect(config.FTP["host"], 21, timeout=45)
    session.login(config.FTP["user"], config.FTP["password"])
    session.set_pasv(True)
    print("Connexion FTP etablie sur %s" % config.FTP["host"])
    return session


def creer_dossier(session, chemin):
    """Cree recursivement un dossier distant, sans echouer s'il existe."""
    courant = ""
    for partie in chemin.strip("/").split("/"):
        courant += "/" + partie
        try:
            session.mkd(courant)
        except ftplib.error_perm as erreur:
            if not str(erreur).startswith("550"):
                raise


def envoyer_dossier(session, local, distant):
    """Copie recursive d'un dossier local vers le FTP."""
    creer_dossier(session, distant)
    for entree in sorted(os.listdir(local)):
        chemin_local = os.path.join(local, entree)
        chemin_distant = "%s/%s" % (distant.rstrip("/"), entree)
        if os.path.isdir(chemin_local):
            envoyer_dossier(session, chemin_local, chemin_distant)
        else:
            with open(chemin_local, "rb") as fichier:
                session.storbinary("STOR %s" % chemin_distant, fichier)
            print("    + %s" % chemin_distant)


def main():
    analyseur = argparse.ArgumentParser(description="Deploiement FTP des maquettes")
    analyseur.add_argument("--tout", action="store_true",
                           help="renvoyer meme les demos deja deployees")
    analyseur.add_argument("--slug", help="ne deployer qu'une societe")
    analyseur.add_argument("--dry-run", action="store_true",
                           help="simulation, aucune connexion FTP")
    arguments = analyseur.parse_args()

    if not os.path.exists(config.FICHIER_MANIFEST):
        raise SystemExit("manifest.json introuvable : lancez generateur.py d'abord.")

    with open(config.FICHIER_MANIFEST, encoding="utf-8") as fichier:
        manifest = json.load(fichier)

    cibles = [
        entree for entree in manifest
        if (arguments.tout or not entree.get("deploye"))
        and (not arguments.slug or entree["slug"] == arguments.slug)
    ]
    if not cibles:
        print("Rien a deployer.")
        return 0

    racine_distante = config.FTP["dossier"].rstrip("/")
    base_url = config.FTP["url_publique"].rstrip("/")

    if arguments.dry_run:
        print("Simulation : %d maquette(s) seraient envoyees vers %s"
              % (len(cibles), racine_distante))
        for entree in cibles:
            print("  %s -> %s/%s/" % (entree["nom"], racine_distante, entree["slug"]))
        return 0

    session = connexion()
    envoyees = 0
    try:
        for index, entree in enumerate(cibles, 1):
            dossier_local = os.path.join(config.RACINE, entree["chemin_local"])
            if not os.path.isdir(dossier_local):
                print("  ! dossier absent : %s" % dossier_local)
                continue
            print("  [%d/%d] %s" % (index, len(cibles), entree["nom"]))
            envoyer_dossier(session, dossier_local,
                            "%s/%s" % (racine_distante, entree["slug"]))
            entree["deploye"] = True
            entree["url"] = "%s/%s/" % (base_url, entree["slug"])
            envoyees += 1
    finally:
        try:
            session.quit()
        except Exception:  # noqa: BLE001
            session.close()

    with open(config.FICHIER_MANIFEST, "w", encoding="utf-8") as fichier:
        json.dump(manifest, fichier, ensure_ascii=False, indent=2)

    print("\n%d maquette(s) en ligne sous %s/" % (envoyees, base_url))
    return 0


if __name__ == "__main__":
    sys.exit(main())
