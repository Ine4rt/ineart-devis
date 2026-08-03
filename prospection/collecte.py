"""
IneArt - Prospection Huy : collecte des entreprises sans site internet.

Etape 1 : OpenStreetMap (Overpass API)
    Liste tous les commerces, artisans, professions liberales et services de la
    commune, avec nom, adresse, telephone, horaires et tags de contact.
    Gratuit, sans compte, sans limite d'usage raisonnable.

Etape 2 (optionnelle, --enrich) : Google Maps via Playwright
    Complete la fiche : note, nombre d'avis, categorie reelle, presence d'un
    site, photo de devanture, capture d'ecran de la fiche.

Segmentation produite :
    A = aucune presence web detectee          -> prospect vierge
    B = uniquement Facebook / Instagram       -> meilleur segment
    C = site declare mais mort ou inaccessible -> a verifier
    D = site fonctionnel                       -> exclu du fichier final

Usage :
    python collecte.py                    # OSM seul
    python collecte.py --enrich           # + Google Maps (Playwright requis)
    python collecte.py --limite 20        # borne le nombre de prospects
"""

import argparse
import json
import os
import re
import sys
import time
import unicodedata

import requests

import config

OVERPASS_MIRRORS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.osm.ch/api/interpreter",
]

NOMINATIM = "https://nominatim.openstreetmap.org/search"

# Filet de securite : la Belgique entiere. Sans cela, une commune homonyme
# ailleurs en Europe (il existe un Huy en Saxe-Anhalt) remonte dans les
# resultats.
BBOX_BELGIQUE = (49.40, 2.45, 51.55, 6.45)  # sud, ouest, nord, est

USER_AGENT = "IneArt-Prospection/1.0 (contact: info@ineart.be)"

# Categories OSM retenues : tout ce qui est une activite commerciale locale.
REQUETE_OVERPASS = """
[out:json][timeout:180];
{zones}
(
  nwr["shop"](area.z);
  nwr["craft"](area.z);
  nwr["office"](area.z);
  nwr["healthcare"](area.z);
  nwr["amenity"~"^(restaurant|cafe|bar|pub|fast_food|ice_cream|pharmacy|veterinary|dentist|doctors|driving_school|childcare|kindergarten|car_wash|car_rental|bank|nightclub|internet_cafe)$"](area.z);
  nwr["leisure"~"^(fitness_centre|sports_centre|dance)$"](area.z);
  nwr["tourism"~"^(hotel|guest_house|bed_and_breakfast|apartment)$"](area.z);
);
out center tags;
"""

CLES_SITE = ("website", "contact:website", "url", "website:menu", "brand:website")

# Enseignes de chaine et services publics : ce ne sont pas des prospects.
# Une franchise depend du site de son groupe, une administration a le sien.
CLES_ENSEIGNE = ("brand", "brand:wikidata", "operator:wikidata")

OPERATEURS_PUBLICS = re.compile(
    r"\b(ville|commune|cpas|province|r[ée]gion|f[ée]d[ée]ration|"
    r"communaut[ée]|asbl communale|infrabel|forem|onem)\b", re.I)

NOMS_EXCLUS = re.compile(
    r"\b(delhaize|carrefour|aldi|lidl|colruyt|okay|spar|intermarch[ée]|match|"
    r"louis delhaize|cora|makro|action|casa|kruidvat|di\b|hubo|brico|gamma|"
    r"basic.?fit|mcdonald|quick|burger\s?king|panos|hema|jbc|zeb|c&a|bel&bo|"
    r"kr[ée]fel|mediamarkt|fnac|standaard boekhandel|club\b|torfs|"
    r"multipharma|medi.?market|pharmacie du peuple|dp pharma|"
    r"espace public num[ée]rique|biblioth[èe]que|maison communale|h[ôo]tel de ville|"
    r"agence locale pour l'emploi|\(ale\)|\bmcae\b|accueil de l'enfance|"
    r"mutualit[ée]|maison de l'emploi|centre pms|"
    r"office du tourisme|piscine communale|centre culturel)\b", re.I)
CLES_SOCIAL = ("contact:facebook", "facebook", "contact:instagram", "instagram")
CLES_TEL = ("phone", "contact:phone", "contact:mobile", "mobile")
CLES_MAIL = ("email", "contact:email")

JOURS_FR = {
    "Mo": "Lundi", "Tu": "Mardi", "We": "Mercredi", "Th": "Jeudi",
    "Fr": "Vendredi", "Sa": "Samedi", "Su": "Dimanche",
}


def slugifier(texte):
    texte = unicodedata.normalize("NFKD", texte)
    texte = texte.encode("ascii", "ignore").decode("ascii")
    texte = re.sub(r"[^a-zA-Z0-9]+", "-", texte).strip("-").lower()
    return re.sub(r"-+", "-", texte) or "societe"


def resoudre_communes(communes):
    """Traduit des noms de communes en identifiants de zones OSM belges.

    Passer par Nominatim avec countrycodes=be evite de ramasser les communes
    homonymes d'autres pays.
    """
    identifiants = []
    for commune in communes:
        nom = commune.strip()
        try:
            reponse = requests.get(
                NOMINATIM,
                params={"q": nom, "countrycodes": "be", "format": "json",
                        "limit": 3, "addressdetails": 1},
                headers={"User-Agent": USER_AGENT},
                timeout=30,
            )
            reponse.raise_for_status()
            for resultat in reponse.json():
                if resultat.get("osm_type") == "relation":
                    identifiants.append(3600000000 + int(resultat["osm_id"]))
                    print("  %s -> zone OSM %d" % (nom, identifiants[-1]))
                    break
            else:
                print("  ! %s : aucune relation trouvee en Belgique" % nom)
        except Exception as erreur:  # noqa: BLE001
            print("  ! Nominatim indisponible pour %s (%s)" % (nom, erreur))
    return identifiants


def interroger_overpass(communes):
    """Interroge Overpass avec bascule automatique entre les miroirs."""
    identifiants = resoudre_communes(communes)
    if identifiants:
        zones = "(%s)->.z;" % "".join("area(%d);" % i for i in identifiants)
    else:
        # Repli : recherche par nom, bornee ensuite a la Belgique.
        noms = "|".join(re.escape(c.strip()) for c in communes)
        zones = ('area["name"~"^(%s)$"]["boundary"="administrative"]'
                 '["admin_level"="8"]->.z;' % noms)
        print("  repli sur la recherche par nom (filtrage Belgique applique)")
    requete = REQUETE_OVERPASS.replace("{zones}", zones)

    derniere_erreur = None
    for miroir in OVERPASS_MIRRORS:
        for tentative in range(3):
            try:
                print("  -> %s (tentative %d)" % (miroir, tentative + 1))
                reponse = requests.post(
                    miroir,
                    data={"data": requete},
                    headers={"User-Agent": USER_AGENT},
                    timeout=200,
                )
                reponse.raise_for_status()
                return reponse.json().get("elements", [])
            except Exception as erreur:  # noqa: BLE001
                derniere_erreur = erreur
                print("     echec : %s" % erreur)
                time.sleep(2 ** tentative)
    raise RuntimeError("Overpass injoignable sur tous les miroirs : %s" % derniere_erreur)


def categorie_osm(tags):
    for cle in ("shop", "craft", "amenity", "office", "healthcare", "leisure", "tourism"):
        if tags.get(cle) and tags[cle] != "yes":
            return tags[cle]
    for cle in ("shop", "craft", "office", "healthcare"):
        if tags.get(cle):
            return cle
    return ""


def lire_horaires(brut):
    """Convertit un tag opening_hours OSM en liste lisible. Best effort."""
    if not brut:
        return []
    lignes = []
    for bloc in brut.split(";"):
        bloc = bloc.strip()
        if not bloc:
            continue
        correspondance = re.match(r"^([A-Za-z,\-]+)\s+(.*)$", bloc)
        if not correspondance:
            lignes.append(("", bloc))
            continue
        jours, heures = correspondance.groups()
        for code, nom in JOURS_FR.items():
            jours = jours.replace(code, nom)
        lignes.append((jours.replace(",", ", ").replace("-", " à "), heures))
    return lignes


def construire_adresse(tags):
    rue = tags.get("addr:street", "")
    numero = tags.get("addr:housenumber", "")
    cp = tags.get("addr:postcode", "")
    ville = tags.get("addr:city", "")
    ligne1 = (" ".join(x for x in (rue, numero) if x)).strip()
    ligne2 = (" ".join(x for x in (cp, ville) if x)).strip()
    return {
        "ligne1": ligne1,
        "ligne2": ligne2,
        "complete": ", ".join(x for x in (ligne1, ligne2) if x),
        "code_postal": cp,
        "ville": ville,
    }


def segmenter(tags):
    """A = rien, B = reseaux sociaux uniquement, D = site declare."""
    site = next((tags[c] for c in CLES_SITE if tags.get(c)), "")
    social = next((tags[c] for c in CLES_SOCIAL if tags.get(c)), "")
    if site:
        if re.search(r"(facebook|instagram|linktr\.ee)", site, re.I):
            return "B", "", site
        return "D", site, social
    if social:
        return "B", "", social
    return "A", "", ""


def est_exclu(tags, nom):
    """Chaines, franchises et services publics : hors cible."""
    if any(tags.get(cle) for cle in CLES_ENSEIGNE):
        return "enseigne de chaine"
    if OPERATEURS_PUBLICS.search(tags.get("operator", "")):
        return "service public"
    if tags.get("office") == "government" or tags.get("government"):
        return "administration"
    if NOMS_EXCLUS.search(nom):
        return "enseigne connue"
    return ""


def normaliser(element):
    tags = element.get("tags", {})
    nom = tags.get("name", "").strip()
    if not nom:
        return None
    motif = est_exclu(tags, nom)
    if motif:
        return {"_exclu": motif, "nom": nom}

    segment, site, social = segmenter(tags)
    adresse = construire_adresse(tags)
    centre = element.get("center") or element
    lat, lon = centre.get("lat"), centre.get("lon")
    recherche_maps = requests.utils.quote(
        "%s %s" % (nom, adresse["ligne2"] or "Huy Belgique"))

    return {
        "id": "%s%s" % (element.get("type", "n")[0], element.get("id")),
        "nom": nom,
        "slug": slugifier(nom),
        "segment": segment,
        "categorie_osm": categorie_osm(tags),
        "categorie_google": "",
        "adresse": adresse,
        "lat": lat,
        "lon": lon,
        "telephone": next((tags[c] for c in CLES_TEL if tags.get(c)), ""),
        "email": next((tags[c] for c in CLES_MAIL if tags.get(c)), ""),
        "facebook": social,
        "site_declare": site,
        "horaires": lire_horaires(tags.get("opening_hours", "")),
        "description": tags.get("description", ""),
        "note": None,
        "avis": None,
        "photos": [],
        "capture_google": "",
        "capture_facebook": "",
        "source": "openstreetmap",
        "osm_url": "https://www.openstreetmap.org/%s/%s" % (element.get("type"), element.get("id")),
        "maps_url": "https://www.google.com/maps/search/%s" % recherche_maps,
    }


def deja_traites():
    """Slugs des societes livrees lors des campagnes precedentes."""
    chemin = os.path.join(config.DOSSIER_DATA, "deja_traites.json")
    if not os.path.exists(chemin):
        return set()
    with open(chemin, encoding="utf-8") as fichier:
        return set(json.load(fichier).get("slugs", []))


def enregistrer_traites(prospects):
    """Ajoute les nouveaux slugs pour qu'ils ne ressortent pas au prochain run."""
    chemin = os.path.join(config.DOSSIER_DATA, "deja_traites.json")
    slugs = deja_traites() | {p["slug"] for p in prospects}
    with open(chemin, "w", encoding="utf-8") as fichier:
        json.dump({"commentaire": "Societes deja traitees : exclues des "
                                  "collectes suivantes.",
                   "slugs": sorted(slugs)},
                  fichier, ensure_ascii=False, indent=2)


def collecter_osm(communes):
    print("Collecte OpenStreetMap : %s" % ", ".join(communes))
    elements = interroger_overpass(communes)
    print("  %d objets recus" % len(elements))

    sud, ouest, nord, est = BBOX_BELGIQUE
    traites = deja_traites()
    prospects, vus, hors_zone, exclus, revus = [], set(), 0, [], 0
    for element in elements:
        fiche = normaliser(element)
        if not fiche:
            continue
        if fiche.get("_exclu"):
            exclus.append("%s (%s)" % (fiche["nom"], fiche["_exclu"]))
            continue
        lat, lon = fiche.get("lat"), fiche.get("lon")
        if lat is None or not (sud <= lat <= nord and ouest <= lon <= est):
            hors_zone += 1
            continue
        if fiche["slug"] in traites:
            revus += 1
            continue
        cle = (fiche["nom"].lower(), fiche["adresse"]["ligne1"].lower())
        if cle in vus:
            continue
        vus.add(cle)
        prospects.append(fiche)

    if hors_zone:
        print("  %d objets ecartes : hors Belgique (commune homonyme)" % hors_zone)
    if revus:
        print("  %d objets ecartes : deja traites lors d'une campagne precedente"
              % revus)
    if exclus:
        print("  %d objets ecartes : chaines et services publics" % len(exclus))
        for libelle in exclus[:12]:
            print("     - %s" % libelle)

    print("  %d etablissements nommes et dedoublonnes" % len(prospects))
    return prospects


# --- Enrichissement Google Maps ---------------------------------------------

def enrichir_google(prospects, dossier_captures, limite=None):
    """Complete les fiches via Google Maps. Necessite playwright."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("! playwright absent : enrichissement Google ignore.")
        print("  pip install playwright && playwright install chromium")
        return prospects

    os.makedirs(dossier_captures, exist_ok=True)
    cibles = prospects[:limite] if limite else prospects

    with sync_playwright() as p:
        navigateur = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        contexte = navigateur.new_context(
            locale="fr-BE",
            viewport={"width": 1366, "height": 900},
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
        )
        page = contexte.new_page()

        for index, fiche in enumerate(cibles, 1):
            requete = "%s %s" % (fiche["nom"], fiche["adresse"]["ligne2"] or "Huy Belgique")
            url = "https://www.google.com/maps/search/%s?hl=fr&gl=be" % requests.utils.quote(requete)
            print("  [%d/%d] %s" % (index, len(cibles), fiche["nom"]))
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=45000)
                _accepter_cookies(page)
                page.wait_for_timeout(3500)

                fiche["note"], fiche["avis"] = _note_et_avis(page)
                fiche["categorie_google"] = _extraire_texte(page, "button[jsaction*='category']")
                site_google = _extraire_site(page)
                if site_google:
                    if re.search(r"(facebook|instagram)", site_google, re.I):
                        fiche["facebook"] = fiche["facebook"] or site_google
                        fiche["segment"] = "B"
                    else:
                        fiche["site_declare"] = site_google
                        fiche["segment"] = "D"
                fiche["photos"] = _extraire_photos(page)

                chemin = os.path.join(dossier_captures, "%s-google.png" % fiche["slug"])
                page.screenshot(path=chemin, full_page=False)
                fiche["capture_google"] = os.path.relpath(chemin, config.RACINE)
                fiche["source"] = "openstreetmap+google"
            except Exception as erreur:  # noqa: BLE001
                print("     ! %s" % erreur)
            time.sleep(1.5)

        navigateur.close()
    return prospects


def _accepter_cookies(page):
    for selecteur in ("button[aria-label*='Tout accepter']",
                      "button:has-text('Tout accepter')",
                      "form[action*='consent'] button"):
        try:
            bouton = page.locator(selecteur).first
            if bouton.is_visible(timeout=1500):
                bouton.click()
                page.wait_for_timeout(1500)
                return
        except Exception:  # noqa: BLE001
            continue


def _extraire_texte(page, selecteur):
    try:
        element = page.locator(selecteur).first
        if element.is_visible(timeout=1200):
            return (element.inner_text() or "").strip()
    except Exception:  # noqa: BLE001
        pass
    return ""


def _note_et_avis(page):
    """Lit la note et le nombre d'avis dans le texte visible de la fiche.

    Les classes CSS de Google changent sans preavis : on cherche donc le
    motif '4,5 (128 avis)' dans le texte plutot qu'un selecteur precis.
    """
    try:
        texte = page.locator("body").inner_text(timeout=4000)
    except Exception:  # noqa: BLE001
        return None, None

    motif = re.search(r"([1-5][,.]\d)[^\d]{0,12}([\d\s.]{1,9})\s*avis", texte)
    if motif:
        chiffres = re.sub(r"\D", "", motif.group(2))
        return (float(motif.group(1).replace(",", ".")),
                int(chiffres) if chiffres else None)

    seule = re.search(r"\b([1-5][,.]\d)\s*(?:etoile|\u00e9toile|star)", texte, re.I)
    if seule:
        return float(seule.group(1).replace(",", ".")), None
    return None, None


def _extraire_site(page):
    try:
        lien = page.locator("a[data-item-id='authority']").first
        if lien.is_visible(timeout=1500):
            return lien.get_attribute("href") or ""
    except Exception:  # noqa: BLE001
        pass
    return ""


def _extraire_photos(page, maximum=4):
    urls = []
    try:
        for image in page.locator("img[src*='googleusercontent']").all()[:12]:
            src = image.get_attribute("src") or ""
            if src.startswith("http") and "=w" in src and src not in urls:
                urls.append(re.sub(r"=w\d+-h\d+", "=w1200-h800", src))
            if len(urls) >= maximum:
                break
    except Exception:  # noqa: BLE001
        pass
    return urls


# --- Verification des sites declares ----------------------------------------

def verifier_sites(prospects):
    """Passe en segment C les sites declares qui ne repondent pas."""
    for fiche in prospects:
        if fiche["segment"] != "D" or not fiche["site_declare"]:
            continue
        url = fiche["site_declare"]
        if not url.startswith("http"):
            url = "https://" + url
        try:
            reponse = requests.get(url, timeout=12, headers={"User-Agent": USER_AGENT},
                                   allow_redirects=True)
            if reponse.status_code >= 400 or len(reponse.text) < 500:
                fiche["segment"] = "C"
                fiche["motif_segment"] = "site en erreur (%s)" % reponse.status_code
        except Exception as erreur:  # noqa: BLE001
            fiche["segment"] = "C"
            fiche["motif_segment"] = "site injoignable (%s)" % type(erreur).__name__
    return prospects


def main():
    analyseur = argparse.ArgumentParser(description="Collecte des prospects de Huy")
    analyseur.add_argument("--enrich", action="store_true",
                           help="enrichir via Google Maps (Playwright)")
    analyseur.add_argument("--limite", type=int, default=config.LIMITE_PROSPECTS,
                           help="nombre max de prospects conserves (0 = tous)")
    analyseur.add_argument("--communes", default=",".join(config.COMMUNES))
    arguments = analyseur.parse_args()

    communes = [c.strip() for c in arguments.communes.split(",") if c.strip()]
    prospects = collecter_osm(communes)

    print("Verification des sites declares...")
    prospects = verifier_sites(prospects)

    # On ne garde que les prospects sans vrai site.
    retenus = [p for p in prospects if p["segment"] in ("A", "B", "C")]
    ordre = {"B": 0, "C": 1, "A": 2}
    # Priorite absolue aux societes joignables par mail : ce sont les seules
    # que l'on peut contacter directement.
    retenus.sort(key=lambda p: (not p["email"], ordre[p["segment"]],
                                not p["telephone"], p["nom"]))

    if arguments.enrich:
        # Google revele des sites qu'OSM ignore : on enrichit large, puis on
        # elimine ceux qui se revelent avoir un vrai site avant de couper.
        marge = int(arguments.limite * 2) if arguments.limite else 0
        candidats = retenus[:marge] if marge else retenus
        print("Enrichissement Google Maps sur %d candidats..." % len(candidats))
        try:
            candidats = enrichir_google(candidats,
                                        os.path.join(config.DOSSIER_DATA, "captures"))
        except Exception as erreur:  # noqa: BLE001
            # Google peut bloquer les IP de datacenter : on garde les donnees OSM.
            print("! enrichissement abandonne (%s) : donnees OSM conservees."
                  % type(erreur).__name__)

        ecartes = [p for p in candidats if p["segment"] == "D"]
        if ecartes:
            print("%d prospect(s) ecarte(s) : un site a ete trouve sur Google"
                  % len(ecartes))
            for fiche in ecartes:
                print("   - %s (%s)" % (fiche["nom"], fiche["site_declare"]))
        retenus = [p for p in candidats if p["segment"] in ("A", "B", "C")]

    if arguments.limite:
        retenus = retenus[:arguments.limite]

    os.makedirs(config.DOSSIER_DATA, exist_ok=True)
    with open(config.FICHIER_PROSPECTS, "w", encoding="utf-8") as fichier:
        json.dump(retenus, fichier, ensure_ascii=False, indent=2)
    enregistrer_traites(retenus)

    compteur = {}
    for fiche in retenus:
        compteur[fiche["segment"]] = compteur.get(fiche["segment"], 0) + 1
    print("\n%d prospects retenus -> %s" % (len(retenus), config.FICHIER_PROSPECTS))
    for segment in ("A", "B", "C"):
        if segment in compteur:
            print("   segment %s : %d" % (segment, compteur[segment]))
    avec_mail = sum(1 for f in retenus if f["email"])
    print("   avec adresse mail : %d (%d%%)"
          % (avec_mail, round(100 * avec_mail / max(len(retenus), 1))))


if __name__ == "__main__":
    sys.exit(main())
