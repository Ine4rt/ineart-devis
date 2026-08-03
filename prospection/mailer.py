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


def accroche(entree):
    """Ouverture adaptee a ce que l'on sait reellement du prospect."""
    nom = entree["nom"]
    if entree.get("segment") == "B" and entree.get("facebook"):
        return ("En cherchant %s sur internet, j'ai trouvé votre page Facebook, "
                "mais pas de site. Or une page Facebook ne sort pas dans Google, "
                "et tout le monde n'y est pas." % nom)
    if entree.get("segment") == "C":
        return ("En cherchant %s sur internet, je suis tombé sur une adresse de "
                "site qui ne répond plus. Vos clients tombent dessus aussi." % nom)
    return ("En cherchant %s sur internet, je n'ai trouvé ni site, ni page : "
            "quelqu'un qui ne vous connaît pas encore n'a aucun moyen de vous "
            "trouver." % nom)


def preuve_sociale(entree):
    """Reprend la note Google du prospect quand elle existe."""
    note, avis = entree.get("note"), entree.get("avis")
    if not note:
        return ""
    note_fr = str(note).replace(".", ",")
    if avis:
        return ("Vos %s avis Google à %s/5 parlent déjà pour vous. Il ne vous "
                "manque qu'une adresse à donner.\n\n" % (avis, note_fr))
    return ("Votre note de %s/5 sur Google parle déjà pour vous. Il ne vous "
            "manque qu'une adresse à donner.\n\n" % note_fr)


def bloc_tarifs():
    site = config.TARIFS["site"]
    domaine = config.TARIFS["domaine"]
    return ("  - %s, exactement comme vous le voyez : %d € %s\n"
            "  - %s (ex. votre-nom.be) : %d € %s\n"
            % (site["libelle"], site["prix"], site["unite"],
               domaine["libelle"], domaine["prix"], domaine["unite"]))


def signature():
    """Signature de fin de mail. Le telephone n'apparait que s'il est defini."""
    agence = config.AGENCE
    coordonnees = [c for c in (agence.get("telephone"), agence["email"],
                               agence["site"]) if c]
    return "Bien à vous,\n%s\n%s\n" % (agence["signature"],
                                          "\n".join(coordonnees))


def texte_mail(entree, relance=False):
    if relance:
        return (
            "Bonjour,\n\n"
            "Petit rappel : l'aperçu du site que j'ai réalisé pour %s est "
            "encore visible ici jusqu'au %s :\n\n"
            "    %s\n\n"
            "Passé cette date, la page est supprimée automatiquement de mon "
            "serveur. Si vous souhaitez la garder, un simple mot suffit.\n\n"
            "Pour rappel, tout compris :\n%s\n"
            "%s"
            % (entree["nom"], entree["expire_le_fr"], entree["url"],
               bloc_tarifs(), signature())
        )

    return (
        "Bonjour,\n\n"
        "Je m'appelle %s, je crée des sites internet pour les commerces et les "
        "indépendants de Huy.\n\n"
        "%s\n\n"
        "%s"
        "Alors j'ai pris les devants : j'ai réalisé votre site, gratuitement et "
        "sans engagement, avec vos photos, vos horaires et vos coordonnées. "
        "Le voici :\n\n"
        "    %s\n\n"
        "Il reste en ligne jusqu'au %s, puis il est supprimé automatiquement. "
        "Rien à faire, rien à payer pour le regarder.\n\n"
        "Si vous voulez le garder, c'est simple et c'est tout compris :\n\n"
        "%s\n"
        "Pas d'abonnement caché, pas de frais de mise en ligne. Le formulaire "
        "de contact arrive directement dans votre boîte mail.\n\n"
        "Si cela ne vous intéresse pas, ignorez ce message : la page "
        "disparaîtra d'elle-même. Et si vous préférez qu'elle disparaisse tout "
        "de suite, un mot suffit, je la retire dans la journée.\n\n"
        "%s"
        % (config.AGENCE["signature"].split(" ")[0],
           accroche(entree), preuve_sociale(entree), entree["url"],
           entree["expire_le_fr"], bloc_tarifs(), signature())
    )


def html_mail(entree, corps_texte):
    """Version HTML : meme texte, avec le lien et les prix mis en evidence."""
    site = config.TARIFS["site"]
    domaine = config.TARIFS["domaine"]

    paragraphes = []
    for bloc in corps_texte.split("\n\n"):
        bloc = bloc.strip()
        if not bloc:
            continue
        # Le lien et le bloc tarifs sont rendus a part, en encadre.
        if bloc.startswith(entree["url"]) or bloc.startswith("- %s" % site["libelle"]):
            continue
        paragraphes.append("<p>%s</p>" % bloc.replace("\n", "<br>"))

    encadre_prix = (
        '<table role="presentation" style="width:100%%;border-collapse:collapse;'
        'margin:22px 0;border:1px solid #e3e6ea;border-radius:10px">'
        '<tr><td style="padding:18px 20px">'
        '<div style="font-size:13px;text-transform:uppercase;letter-spacing:.12em;'
        'color:#6b7280;margin-bottom:12px">Tout compris</div>'
        '<div style="font-size:16px;margin-bottom:8px">'
        '<b>%s €</b> une seule fois — %s, exactement comme vous le voyez</div>'
        '<div style="font-size:16px">'
        '<b>%s €</b> par an — %s (ex. votre-nom.be)</div>'
        '</td></tr></table>'
        % (site["prix"], site["libelle"], domaine["prix"], domaine["libelle"])
    )

    return (
        '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;'
        'font-size:15px;line-height:1.65;color:#1a1a1a;max-width:600px">'
        '%s'
        '<p style="margin:26px 0">'
        '<a href="%s" style="background:#1a1a1a;color:#fff;text-decoration:none;'
        'padding:13px 26px;border-radius:99px;display:inline-block;font-weight:600">'
        'Voir votre site</a><br>'
        '<a href="%s" style="font-size:13px;color:#6b7280">%s</a></p>'
        '%s</div>'
        % ("".join(paragraphes), entree["url"], entree["url"], entree["url"],
           encadre_prix)
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
    message["To"] = entree.get("email", "")
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
    attendus.add("TOUS_LES_MAILS.txt")

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
    lignes_csv, corps_complet, avec_mail, sans_mail = [], [], 0, 0

    for entree in manifest:
        if entree.get("supprime"):
            continue
        # Un mail est prepare pour chaque societe. Sans adresse connue, le
        # champ destinataire reste vide : il suffit de le completer une fois
        # l'adresse trouvee (page Facebook, vitrine, appel).
        fichier_genere = os.path.basename(ecrire_eml(entree, arguments.relance))
        if entree.get("email"):
            canal = "mail"
            avec_mail += 1
        else:
            canal = "telephone"
            sans_mail += 1
            chemin = os.path.join(config.DOSSIER_MAILS, "%s-appel.txt" % entree["slug"])
            with open(chemin, "w", encoding="utf-8") as fichier:
                fichier.write(script_appel(entree))

        corps_complet.append(
            "%s\n%s\nÀ       : %s\nObjet   : %s\nLien    : %s\n%s\n\n%s"
            % ("=" * 74, entree["nom"],
               entree.get("email") or "(adresse à trouver — %s)"
               % (entree.get("telephone") or "pas de téléphone connu"),
               (OBJET_RELANCE if arguments.relance else OBJET).format(
                   nom=entree["nom"], date=entree["expire_le_fr"]),
               entree["url"], "=" * 74,
               texte_mail(entree, arguments.relance))
        )

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

    chemin_recap = os.path.join(config.DOSSIER_MAILS, "TOUS_LES_MAILS.txt")
    with open(chemin_recap, "w", encoding="utf-8") as fichier:
        fichier.write("MAILS DE PROSPECTION — %d societes\n"
                      "Genere le %s\n\n%s\n"
                      % (len(corps_complet),
                         dt.datetime.now().strftime("%d/%m/%Y a %H:%M"),
                         "\n\n".join(corps_complet)))

    with open(config.FICHIER_ENVOIS, "w", encoding="utf-8", newline="") as fichier:
        writer = csv.DictWriter(fichier, fieldnames=list(lignes_csv[0].keys()),
                                delimiter=";")
        writer.writeheader()
        writer.writerows(lignes_csv)

    print("%d mails prets a envoyer (adresse connue)" % avec_mail)
    print("%d mails prets, destinataire a completer (+ fiche d'appel)" % sans_mail)
    print("Recapitulatif lisible : %s" % chemin_recap)
    print("Suivi : %s" % config.FICHIER_ENVOIS)
    print("\nRappel : envoyez par paquets de 20 a 30 par jour, jamais tout d'un coup.")
    print("Genere le %s" % dt.datetime.now().strftime("%d/%m/%Y %H:%M"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
