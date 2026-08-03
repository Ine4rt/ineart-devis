"""
IneArt - Prospection Huy : generation d'un site vitrine par societe.

Lit data/prospects.json et produit :
    sites/<slug>/index.html   le site complet (autonome, aucune dependance)
    sites/<slug>/img/*.jpg    les photos recuperees, si disponibles
    data/manifest.json        le suivi (url, dates, expiration, contact)
    tableau_de_bord.html      recapitulatif local pour IneArt

Chaque site expire au bout de DUREE_DEMO_JOURS : compte a rebours visible,
ecran de fin cote navigateur, et suppression reelle sur le FTP via nettoyage.py.

Usage :
    python generateur.py                     # utilise data/prospects.json
    python generateur.py --exemple           # utilise le jeu de demonstration
    python generateur.py --duree 8
"""

import argparse
import datetime as dt
import html
import json
import os
import re
import shutil
import unicodedata
from urllib.parse import quote

import config
import secteurs

TEMPLATE = os.path.join(config.DOSSIER_TEMPLATES, "site.html")


# --- utilitaires -------------------------------------------------------------

def e(valeur):
    """Echappe une valeur pour insertion dans du HTML."""
    return html.escape(str(valeur or ""), quote=True)


def initiales(nom):
    mots = [m for m in re.split(r"[^A-Za-zÀ-ÿ0-9]+", nom) if m]
    ignores = {"le", "la", "les", "de", "du", "des", "chez", "au", "aux", "l", "d", "et"}
    utiles = [m for m in mots if m.lower() not in ignores] or mots
    if len(utiles) >= 2:
        return (utiles[0][0] + utiles[1][0]).upper()
    return (utiles[0][:2] if utiles else "IN").upper()


def logo_svg(nom, theme):
    """Monogramme SVG : toujours disponible, meme sans logo officiel."""
    return (
        '<svg width="46" height="46" viewBox="0 0 46 46" aria-hidden="true">'
        '<rect width="46" height="46" rx="13" fill="%s"/>'
        '<text x="23" y="30" text-anchor="middle" font-family="%s" font-size="19" '
        'font-weight="700" fill="#fff">%s</text></svg>'
        % (theme["primaire"], "Georgia, serif", e(initiales(nom)))
    )


def favicon(nom, theme):
    svg = (
        "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 46 46'>"
        "<rect width='46' height='46' rx='10' fill='%s'/>"
        "<text x='23' y='31' text-anchor='middle' font-family='Georgia,serif' "
        "font-size='22' font-weight='700' fill='white'>%s</text></svg>"
        % (theme["primaire"], initiales(nom))
    )
    return quote(svg)


def tel_brut(numero):
    nettoye = re.sub(r"[^\d+]", "", numero or "")
    return nettoye or "+32"


def tel_affiche(numero):
    return numero.strip() if numero else "Numéro à confirmer"


def resume_horaires(horaires):
    if not horaires:
        return "Nous consulter"
    jours, heures = horaires[0]
    return ("%s : %s" % (jours, heures)).strip(" :") or "Nous consulter"


# --- fragments HTML ----------------------------------------------------------

def html_services(theme, fiche, secteur):
    blocs = []
    liste = secteurs.services(secteur, fiche.get("categorie_osm"))
    for index, (titre, texte) in enumerate(liste, 1):
        blocs.append(
            '<article class="carte"><div class="puce">%02d</div>'
            '<h3>%s</h3><p>%s</p></article>' % (index, e(titre), e(texte))
        )
    if fiche.get("facebook"):
        blocs.append(
            '<article class="carte"><div class="puce">%02d</div>'
            '<h3>Suivez-nous</h3><p>Nos nouveautés et nos actualités sont '
            'publiées sur notre page Facebook.</p></article>' % (len(blocs) + 1)
        )
    return "\n".join(blocs)


def html_horaires(horaires):
    if not horaires:
        return ('<li><span>Horaires</span><b>Nous consulter</b></li>'
                '<li><span>Sur la version définitive</span><b>Affichés en clair</b></li>')
    return "\n".join(
        '<li><span>%s</span><b>%s</b></li>' % (e(jours or "Ouverture"), e(heures))
        for jours, heures in horaires
    )


def html_galerie(photos_locales):
    if not photos_locales:
        return "", ""
    figures = "\n".join(
        '<figure><img src="%s" alt="" loading="lazy"></figure>' % e(chemin)
        for chemin in photos_locales
    )
    section = (
        '<section id="galerie"><div class="wrap">'
        '<div class="titre-bloc"><div class="surtitre">En images</div>'
        '<h2>Un aperçu de la maison</h2></div>'
        '<div class="galerie">%s</div></div></section>' % figures
    )
    return section, '<a href="#galerie">Photos</a>'


def html_note(fiche):
    note, avis = fiche.get("note"), fiche.get("avis")
    if not note:
        return "", ""
    pleines = int(round(note))
    etoiles = "★" * pleines + "☆" * (5 - pleines)
    bloc = (
        '<div class="note-bloc" style="margin-top:32px">'
        '<div class="note-chiffre">%s</div>'
        '<div><div class="etoiles">%s</div>'
        '<div style="color:var(--gris);font-size:14px">%s avis Google</div></div></div>'
        % (str(note).replace(".", ","), etoiles, avis or "—")
    )
    case = ('<div class="case"><div class="lab">Avis Google</div>'
            '<div class="val">%s ★ · %s avis</div></div>'
            % (str(note).replace(".", ","), avis or "—"))
    return bloc, case


def html_prix():
    blocs = []
    for cle in ("site", "domaine"):
        offre = config.TARIFS[cle]
        points = "\n".join("<li>%s</li>" % e(p) for p in offre["detail"])
        blocs.append(
            '<div class="prix"><div class="lib">%s</div>'
            '<div class="montant">%d €</div><div class="unite">%s</div>'
            '<ul>%s</ul></div>'
            % (e(offre["libelle"]), offre["prix"], e(offre["unite"]), points)
        )
    return "\n".join(blocs)


# --- photos ------------------------------------------------------------------

def telecharger_photos(fiche, dossier):
    """Recupere les photos. Silencieux en cas d'echec reseau."""
    urls = fiche.get("photos") or []
    if not urls:
        return []
    try:
        import requests
    except ImportError:
        return []

    dossier_img = os.path.join(dossier, "img")
    os.makedirs(dossier_img, exist_ok=True)
    locales = []
    for index, url in enumerate(urls[:6], 1):
        chemin = os.path.join(dossier_img, "photo%d.jpg" % index)
        # Photo deja recuperee lors d'un run precedent : on la garde, les URL
        # Google expirent et ne sont pas rejouables.
        if os.path.exists(chemin) and os.path.getsize(chemin) > 1024:
            locales.append("img/photo%d.jpg" % index)
            continue
        try:
            if url.startswith(("http://", "https://")):
                reponse = requests.get(url, timeout=20)
                reponse.raise_for_status()
                with open(chemin, "wb") as fichier:
                    fichier.write(reponse.content)
            elif os.path.exists(url):
                shutil.copy(url, chemin)
            else:
                continue
            locales.append("img/photo%d.jpg" % index)
        except Exception as erreur:  # noqa: BLE001
            print("     photo ignorée (%s)" % type(erreur).__name__)
    return locales


# --- generation --------------------------------------------------------------

def generer_site(fiche, duree_jours, modele, variante="v1"):
    secteur = secteurs.detecter_secteur(fiche)
    theme = secteurs.theme(secteur)
    slug = fiche.get("slug") or "societe"
    dossier = os.path.join(config.DOSSIER_SITES, slug)
    os.makedirs(dossier, exist_ok=True)

    maintenant = dt.datetime.now(dt.timezone.utc).replace(microsecond=0)
    expiration = maintenant + dt.timedelta(days=duree_jours)

    photos = telecharger_photos(fiche, dossier)
    galerie_section, lien_galerie = html_galerie(photos)
    bloc_note, case_avis = html_note(fiche)

    if photos:
        hero_fond = "background:url('%s') center/cover no-repeat;" % photos[0]
        hero_media = ('<div class="hero-media"><img src="%s" alt=""></div>'
                      % e(photos[0]))
    else:
        hero_fond = ("background:linear-gradient(135deg,%s,%s);"
                     % (theme["sombre"], theme["primaire"]))
        hero_media = ""

    # Le telephone de l'agence n'apparait que s'il est renseigne.
    if config.AGENCE.get("telephone"):
        bouton_agence = ('<a class="btn btn-vide" href="tel:%s">Appeler %s</a>'
                         % (e(tel_brut(config.AGENCE["telephone"])),
                            e(config.AGENCE["nom"])))
    else:
        bouton_agence = ('<a class="btn btn-vide" href="mailto:%s">Écrire à %s</a>'
                         % (e(config.AGENCE["email"]), e(config.AGENCE["nom"])))

    adresse = fiche.get("adresse") or {}
    ville = adresse.get("ville") or config.AGENCE["ville"]
    nom = fiche.get("nom", "Votre société")

    # Sans numero connu, mieux vaut orienter vers le formulaire que d'afficher
    # un bouton d'appel qui ne compose rien.
    telephone = (fiche.get("telephone") or "").strip()
    if telephone:
        lien_tel = "tel:%s" % e(tel_brut(telephone))
        bouton_entete = ('<a class="tel-btn" href="%s">%s</a>'
                         % (lien_tel, e(telephone)))
        bouton_hero = ('<a class="btn btn-plein" href="%s">Appeler maintenant</a>'
                       % lien_tel)
        label_contact, valeur_contact = "Téléphone", ('<a href="%s">%s</a>'
                                                      % (lien_tel, e(telephone)))
        ligne_telephone = '<a href="%s">%s</a>' % (lien_tel, e(telephone))
    else:
        bouton_entete = '<a class="tel-btn" href="#contact">Nous écrire</a>'
        bouton_hero = ('<a class="btn btn-plein" href="#contact">'
                       'Nous écrire</a>')
        label_contact, valeur_contact = "Contact", '<a href="#contact">Formulaire</a>'
        ligne_telephone = ('<a href="#contact">Nous écrire</a>')

    valeurs = {
        "NOM": e(nom),
        "VILLE": e(ville),
        "ACCROCHE": e(fiche.get("description") or theme["accroche"]),
        "CATEGORIE_LIBELLE": e(fiche.get("categorie_google") or theme["libelle"]),
        "ADRESSE_LIGNE1": e(adresse.get("ligne1") or "Adresse à confirmer"),
        "ADRESSE_LIGNE2": e(adresse.get("ligne2") or ville),
        "ADRESSE_COMPLETE": e(adresse.get("complete") or ville),
        "BOUTON_ENTETE": bouton_entete,
        "BOUTON_HERO": bouton_hero,
        "LABEL_CONTACT": label_contact,
        "VALEUR_CONTACT": valeur_contact,
        "LIGNE_TELEPHONE": ligne_telephone,
        "MAPS_URL": e(fiche.get("maps_url") or
                      "https://www.google.com/maps/search/" + quote("%s %s" % (nom, ville))),
        "HORAIRE_RESUME": e(resume_horaires(fiche.get("horaires"))),
        "CASE_AVIS": case_avis,
        "BLOC_NOTE": bloc_note,
        "SERVICES_HTML": html_services(theme, fiche, secteur),
        "HORAIRES_HTML": html_horaires(fiche.get("horaires")),
        "GALERIE_SECTION": galerie_section,
        "LIEN_GALERIE": lien_galerie,
        "PRIX_HTML": html_prix(),
        "TITRE_SERVICES": e("Ce que vous trouverez chez %s" % nom),
        "TEXTE_SERVICES": e("Un accueil direct, un travail soigné et des habitudes "
                            "qui durent. Voici l'essentiel de ce que nous proposons."),
        "TITRE_APROPOS": e("Une adresse de %s, tout simplement" % ville),
        "TEXTE_APROPOS": e("Installés à %s, nous mettons un point d'honneur à faire "
                           "les choses correctement et à rester disponibles pour "
                           "chacun de nos clients. Poussez la porte, on prend le "
                           "temps de vous répondre." % ville),
        "LIGNE_FACEBOOK": ('<br><a href="%s" target="_blank" rel="noopener">Notre page Facebook</a>'
                           % e(fiche["facebook"])) if fiche.get("facebook") else "",
        "VARIANTE": variante,
        "HERO_MEDIA": hero_media,
        "BOUTON_AGENCE": bouton_agence,
        "LOGO_SVG": logo_svg(nom, theme),
        "FAVICON": favicon(nom, theme),
        "C_PRIMAIRE": theme["primaire"],
        "C_SOMBRE": theme["sombre"],
        "C_ACCENT": theme["accent"],
        "C_CLAIR": theme["clair"],
        "POLICE_TITRE": theme["police_titre"],
        "HERO_FOND": hero_fond,
        "DUREE_JOURS": str(duree_jours),
        "CREATION_ISO": maintenant.isoformat().replace("+00:00", "Z"),
        "EXPIRATION_ISO": expiration.isoformat().replace("+00:00", "Z"),
        "DATE_EXPIRATION": expiration.strftime("%d/%m/%Y"),
        "AGENCE_NOM": e(config.AGENCE["nom"]),
        "AGENCE_EMAIL": e(config.AGENCE["email"]),
        "AGENCE_SITE": e(config.AGENCE["site"]),
        "AGENCE_TEL_BRUT": e(tel_brut(config.AGENCE["telephone"])),
        "SUJET_MAILTO": quote("Mon site internet — %s" % nom),
    }

    page = modele
    for cle, valeur in valeurs.items():
        page = page.replace("{{%s}}" % cle, str(valeur))

    restants = re.findall(r"\{\{([A-Z_]+)\}\}", page)
    if restants:
        raise RuntimeError("Placeholders non remplaces : %s" % sorted(set(restants)))

    with open(os.path.join(dossier, "index.html"), "w", encoding="utf-8") as fichier:
        fichier.write(page)

    return {
        "slug": slug,
        "nom": nom,
        "secteur": secteur,
        "variante": variante,
        "segment": fiche.get("segment", "A"),
        "categorie": fiche.get("categorie_google") or theme["libelle"],
        "note": fiche.get("note"),
        "avis": fiche.get("avis"),
        "email": fiche.get("email", ""),
        "telephone": fiche.get("telephone", ""),
        "adresse": adresse.get("complete", ""),
        "facebook": fiche.get("facebook", ""),
        "url": "%s/%s/" % (config.FTP["url_publique"].rstrip("/"), slug),
        "chemin_local": os.path.relpath(dossier, config.RACINE),
        "photos": len(photos),
        "cree_le": maintenant.isoformat().replace("+00:00", "Z"),
        "expire_le": expiration.isoformat().replace("+00:00", "Z"),
        "expire_le_fr": expiration.strftime("%d/%m/%Y"),
        "deploye": False,
    }


def tableau_de_bord(entrees):
    lignes = []
    for entree in sorted(entrees, key=lambda x: x["nom"].lower()):
        contact = entree["email"] or entree["telephone"] or "—"
        lignes.append(
            "<tr><td><b>%s</b><br><span class=s>%s</span></td><td>%s</td>"
            "<td>%s</td><td>%s</td><td><a href='%s' target=_blank>ouvrir</a></td>"
            "<td><a href='%s'>local</a></td></tr>"
            % (e(entree["nom"]), e(entree["adresse"]), e(entree["segment"]),
               e(contact), e(entree["expire_le_fr"]), e(entree["url"]),
               e(entree["chemin_local"] + "/index.html"))
        )
    return (
        "<!DOCTYPE html><html lang=fr><head><meta charset=utf-8>"
        "<meta name=viewport content='width=device-width,initial-scale=1'>"
        "<title>IneArt — Tableau de bord prospection</title><style>"
        "body{font:15px/1.6 -apple-system,Segoe UI,Roboto,sans-serif;margin:0;"
        "background:#0f1115;color:#e6e9ef;padding:28px}"
        "h1{font-size:22px;margin:0 0 4px}p.sub{color:#8b93a0;margin:0 0 24px}"
        "table{border-collapse:collapse;width:100%%;background:#171a21;border-radius:12px;"
        "overflow:hidden}th,td{padding:12px 14px;text-align:left;font-size:14px;"
        "border-bottom:1px solid #232833}th{background:#1e222b;color:#9aa3b2;"
        "text-transform:uppercase;font-size:11px;letter-spacing:.1em}"
        ".s{color:#8b93a0;font-size:12px}a{color:#6ea8fe}</style></head><body>"
        "<h1>Prospection Huy — %d maquettes</h1>"
        "<p class=sub>Chaque site expire %d jours apres sa generation.</p>"
        "<table><tr><th>Societe</th><th>Seg.</th><th>Contact</th>"
        "<th>Expire le</th><th>En ligne</th><th>Local</th></tr>%s</table>"
        "</body></html>" % (len(entrees), config.DUREE_DEMO_JOURS, "\n".join(lignes))
    )


def main():
    analyseur = argparse.ArgumentParser(description="Generation des sites demo")
    analyseur.add_argument("--exemple", action="store_true",
                           help="utiliser data/exemple_prospects.json")
    analyseur.add_argument("--duree", type=int, default=config.DUREE_DEMO_JOURS)
    analyseur.add_argument("--limite", type=int, default=0)
    arguments = analyseur.parse_args()

    source = config.FICHIER_EXEMPLE if arguments.exemple else config.FICHIER_PROSPECTS
    if not os.path.exists(source):
        raise SystemExit("Fichier introuvable : %s (lancez collecte.py d'abord)" % source)

    with open(source, encoding="utf-8") as fichier:
        prospects = json.load(fichier)
    if arguments.limite:
        prospects = prospects[:arguments.limite]

    with open(TEMPLATE, encoding="utf-8") as fichier:
        modele = fichier.read()

    os.makedirs(config.DOSSIER_SITES, exist_ok=True)
    entrees = []
    variantes = ["v1", "v2", "v3"]
    for index, fiche in enumerate(prospects, 1):
        variante = variantes[(index - 1) % len(variantes)]
        print("  [%d/%d] %s (%s)" % (index, len(prospects), fiche.get("nom"),
                                     variante))
        entrees.append(generer_site(fiche, arguments.duree, modele, variante))

    # Les societes ecartees depuis la collecte precedente ne doivent pas
    # rester en ligne ni repartir dans le paquet FTP.
    slugs = {entree["slug"] for entree in entrees}
    for nom in sorted(os.listdir(config.DOSSIER_SITES)):
        chemin = os.path.join(config.DOSSIER_SITES, nom)
        if os.path.isdir(chemin) and nom not in slugs:
            shutil.rmtree(chemin)
            print("  - maquette obsolete supprimee : %s" % nom)

    os.makedirs(config.DOSSIER_DATA, exist_ok=True)
    with open(config.FICHIER_MANIFEST, "w", encoding="utf-8") as fichier:
        json.dump(entrees, fichier, ensure_ascii=False, indent=2)

    chemin_bord = os.path.join(config.RACINE, "tableau_de_bord.html")
    with open(chemin_bord, "w", encoding="utf-8") as fichier:
        fichier.write(tableau_de_bord(entrees))

    print("\n%d sites generes dans %s" % (len(entrees), config.DOSSIER_SITES))
    print("Manifest       : %s" % config.FICHIER_MANIFEST)
    print("Tableau de bord: %s" % chemin_bord)


if __name__ == "__main__":
    main()
