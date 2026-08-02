"""
IneArt - Prospection Huy : preparation des envois.

Produit, a partir de data/manifest.json :
    mails/<slug>.eml        message pret a envoyer (double-clic -> client mail)
    mails/<slug>-appel.txt  script d'appel quand aucune adresse mail n'existe
    mails/envois.csv        suivi complet : societe, mail, lien, expiration...

Aucun envoi automatique : les .eml sont ouverts et envoyes a la main, par
paquets de 20 a 30 par jour, pour proteger la reputation du domaine.

Usage :
    python mailer.py
    python mailer.py --relance      # version courte de relance a J+4
"""

import argparse
import csv
import datetime as dt
import json
import os
import sys
from email.message import EmailMessage
from email.utils import formatdate, make_msgid

import config

OBJET = "Votre site internet est déjà en ligne, {nom}"
OBJET_RELANCE = "{nom} — votre aperçu disparaît le {date}"


def texte_mail(entree, relance=False):
    tarif_site = config.TARIFS["site"]
    tarif_domaine = config.TARIFS["domaine"]
    agence = config.AGENCE

    if relance:
        return (
            "Bonjour,\n\n"
            "Petit rappel : l'aperçu du site que j'ai réalisé pour {nom} est "
            "encore visible ici jusqu'au {date} :\n\n"
            "    {url}\n\n"
            "Passé cette date, la page est supprimée automatiquement de mon "
            "serveur. Si vous souhaitez la garder, un simple mot suffit.\n\n"
            "Pour rappel : {prix_site} € une seule fois pour le site, "
            "{prix_domaine} € par an pour le nom de domaine et l'hébergement.\n\n"
            "Bien à vous,\n{signature}\n{tel} — {email}\n{site}\n"
        ).format(
            nom=entree["nom"], date=entree["expire_le_fr"], url=entree["url"],
            prix_site=tarif_site["prix"], prix_domaine=tarif_domaine["prix"],
            signature=agence["signature"], tel=agence["telephone"],
            email=agence["email"], site=agence["site"],
        )

    return (
        "Bonjour,\n\n"
        "Je m'appelle {prenom_signature}. Je crée des sites internet pour les "
        "commerces et les indépendants de la région de Huy.\n\n"
        "En cherchant {nom} sur internet, je me suis rendu compte que vous "
        "n'aviez pas encore de site. J'en ai donc réalisé un, gratuitement et "
        "sans engagement, à partir de vos informations publiques (adresse, "
        "horaires, téléphone). Vous pouvez le voir ici :\n\n"
        "    {url}\n\n"
        "Il est en ligne jusqu'au {date}, après quoi il est supprimé "
        "automatiquement. Rien à faire, rien à payer pour le regarder.\n\n"
        "S'il vous convient, voici mes tarifs :\n"
        "  - {lib_site} : {prix_site} € ({unite_site})\n"
        "  - {lib_domaine} : {prix_domaine} € ({unite_domaine})\n\n"
        "Tout est compris : le design, la version mobile, la mise en ligne et "
        "le formulaire de contact qui arrive directement dans votre boîte mail.\n\n"
        "Si le site ne vous intéresse pas, ignorez simplement ce message : la "
        "page sera retirée d'elle-même. Et si vous préférez qu'elle disparaisse "
        "tout de suite, répondez-moi, je la supprime dans la journée.\n\n"
        "Bien à vous,\n{signature}\n{tel} — {email}\n{site}\n"
    ).format(
        prenom_signature=config.AGENCE["signature"].split(" ")[0],
        nom=entree["nom"], url=entree["url"], date=entree["expire_le_fr"],
        lib_site=tarif_site["libelle"], prix_site=tarif_site["prix"],
        unite_site=tarif_site["unite"], lib_domaine=tarif_domaine["libelle"],
        prix_domaine=tarif_domaine["prix"], unite_domaine=tarif_domaine["unite"],
        signature=config.AGENCE["signature"], tel=config.AGENCE["telephone"],
        email=config.AGENCE["email"], site=config.AGENCE["site"],
    )


def html_mail(entree, corps_texte):
    lignes = "".join(
        "<p>%s</p>" % ligne.replace("\n", "<br>")
        for ligne in corps_texte.split("\n\n") if ligne.strip()
    )
    return (
        '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;'
        'font-size:15px;line-height:1.65;color:#1a1a1a;max-width:600px">%s'
        '<p style="margin-top:26px">'
        '<a href="%s" style="background:#1a1a1a;color:#fff;text-decoration:none;'
        'padding:13px 24px;border-radius:99px;display:inline-block;font-weight:600">'
        'Voir le site</a></p></div>'
        % (lignes, entree["url"])
    )


def script_appel(entree):
    tarif_site = config.TARIFS["site"]
    tarif_domaine = config.TARIFS["domaine"]
    return (
        "FICHE D'APPEL — {nom}\n"
        "{ligne}\n"
        "Téléphone   : {tel}\n"
        "Adresse     : {adresse}\n"
        "Maquette    : {url}\n"
        "Expire le   : {date}\n\n"
        "ACCROCHE\n"
        "  \"Bonjour, {prenom} d'{agence}. Je fais des sites internet pour les\n"
        "   commerces de Huy. J'ai vu que vous n'aviez pas encore de site, donc\n"
        "   j'en ai fait un pour vous, gratuitement. Il est déjà en ligne, je\n"
        "   peux vous envoyer le lien par SMS ?\"\n\n"
        "SI INTÉRESSÉ\n"
        "  - Envoyer le lien par SMS pendant l'appel\n"
        "  - Tarifs : {prix_site} € une fois + {prix_domaine} €/an (domaine + hébergement)\n"
        "  - Proposer un passage en boutique pour ajouter les vraies photos\n\n"
        "SI \"PAS INTÉRESSÉ\"\n"
        "  - \"Aucun souci, la page se supprime toute seule le {date}.\"\n"
        "  - Noter dans envois.csv : statut = refus\n\n"
        "OBJECTIONS COURANTES\n"
        "  \"J'ai déjà Facebook\"  -> Facebook ne sort pas dans Google. Le site,\n"
        "                            si. Les deux se complètent, le site reste.\n"
        "  \"C'est trop cher\"      -> {prix_site} € une seule fois, c'est le prix\n"
        "                            d'une demi-journée de travail.\n"
        "  \"Je n'ai pas le temps\" -> Le site est déjà fait. Il ne reste qu'à\n"
        "                            valider les textes.\n"
    ).format(
        nom=entree["nom"], ligne="=" * (16 + len(entree["nom"])),
        tel=entree["telephone"] or "à trouver",
        adresse=entree["adresse"], url=entree["url"], date=entree["expire_le_fr"],
        prenom=config.AGENCE["signature"].split(" ")[0], agence=config.AGENCE["nom"],
        prix_site=tarif_site["prix"], prix_domaine=tarif_domaine["prix"],
    )


def ecrire_eml(entree, relance):
    corps = texte_mail(entree, relance)
    message = EmailMessage()
    modele_objet = OBJET_RELANCE if relance else OBJET
    message["Subject"] = modele_objet.format(nom=entree["nom"],
                                             date=entree["expire_le_fr"])
    message["From"] = "%s <%s>" % (config.AGENCE["nom"], config.AGENCE["email"])
    message["To"] = entree["email"]
    message["Date"] = formatdate(localtime=True)
    message["Message-ID"] = make_msgid(domain=config.AGENCE["email"].split("@")[-1])
    message.set_content(corps)
    message.add_alternative(html_mail(entree, corps), subtype="html")

    suffixe = "-relance" if relance else ""
    chemin = os.path.join(config.DOSSIER_MAILS, "%s%s.eml" % (entree["slug"], suffixe))
    with open(chemin, "wb") as fichier:
        fichier.write(bytes(message))
    return chemin


def nettoyer_anciens(manifest):
    """Supprime les fichiers des campagnes precedentes.

    Sans cela, un prospect ecarte entre deux collectes laisse son mail dans le
    dossier et finit par etre contacte pour rien.
    """
    attendus = set()
    for entree in manifest:
        attendus.update((
            "%s.eml" % entree["slug"],
            "%s-relance.eml" % entree["slug"],
            "%s-appel.txt" % entree["slug"],
        ))
    attendus.add(os.path.basename(config.FICHIER_ENVOIS))

    retires = 0
    for nom in os.listdir(config.DOSSIER_MAILS):
        if nom not in attendus and nom.endswith((".eml", ".txt")):
            os.remove(os.path.join(config.DOSSIER_MAILS, nom))
            retires += 1
    if retires:
        print("%d fichier(s) d'une campagne precedente supprime(s)" % retires)


def main():
    analyseur = argparse.ArgumentParser(description="Preparation des mails de prospection")
    analyseur.add_argument("--relance", action="store_true",
                           help="generer la version courte de relance")
    arguments = analyseur.parse_args()

    if not os.path.exists(config.FICHIER_MANIFEST):
        raise SystemExit("manifest.json introuvable : lancez generateur.py d'abord.")

    with open(config.FICHIER_MANIFEST, encoding="utf-8") as fichier:
        manifest = json.load(fichier)

    os.makedirs(config.DOSSIER_MAILS, exist_ok=True)
    nettoyer_anciens(manifest)
    lignes_csv, avec_mail, sans_mail = [], 0, 0

    for entree in manifest:
        if entree.get("supprime"):
            continue
        canal, fichier_genere = "telephone", ""

        if entree.get("email"):
            fichier_genere = os.path.basename(ecrire_eml(entree, arguments.relance))
            canal = "mail"
            avec_mail += 1
        else:
            chemin = os.path.join(config.DOSSIER_MAILS, "%s-appel.txt" % entree["slug"])
            with open(chemin, "w", encoding="utf-8") as fichier:
                fichier.write(script_appel(entree))
            fichier_genere = os.path.basename(chemin)
            sans_mail += 1

        lignes_csv.append({
            "societe": entree["nom"],
            "segment": entree["segment"],
            "secteur": entree["secteur"],
            "email": entree.get("email", ""),
            "telephone": entree.get("telephone", ""),
            "adresse": entree.get("adresse", ""),
            "facebook": entree.get("facebook", ""),
            "lien_demo": entree["url"],
            "expire_le": entree["expire_le_fr"],
            "canal_conseille": canal,
            "fichier": fichier_genere,
            "objet": (OBJET_RELANCE if arguments.relance else OBJET).format(
                nom=entree["nom"], date=entree["expire_le_fr"]),
            "statut": "a_envoyer",
            "date_envoi": "",
            "reponse": "",
        })

    with open(config.FICHIER_ENVOIS, "w", encoding="utf-8", newline="") as fichier:
        writer = csv.DictWriter(fichier, fieldnames=list(lignes_csv[0].keys()),
                                delimiter=";")
        writer.writeheader()
        writer.writerows(lignes_csv)

    print("%d mails prets (.eml)" % avec_mail)
    print("%d fiches d'appel (pas d'adresse mail connue)" % sans_mail)
    print("Suivi : %s" % config.FICHIER_ENVOIS)
    print("\nRappel : envoyez par paquets de 20 a 30 par jour, jamais tout d'un coup.")
    print("Genere le %s" % dt.datetime.now().strftime("%d/%m/%Y %H:%M"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
