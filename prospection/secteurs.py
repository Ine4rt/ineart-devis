"""
Classification metier -> secteur -> theme graphique.

Le secteur pilote a la fois la palette du site demo, le vocabulaire des textes
et la liste de services proposee par defaut.
"""

import re
import unicodedata

# Mapping des valeurs OSM (shop / amenity / craft / office / healthcare)
# vers un secteur interne.
MAPPING = {
    "horeca": [
        "restaurant", "cafe", "bar", "pub", "fast_food", "food_court",
        "ice_cream", "bakery", "pastry", "butcher", "deli", "confectionery",
        "greengrocer", "cheese", "chocolate", "brewery", "caterer", "wine",
        "beverages", "alcohol", "coffee", "seafood", "tea",
        "hotel", "guest_house", "bed_and_breakfast", "apartment",
    ],
    "beaute": [
        "hairdresser", "beauty", "massage", "cosmetics", "nail_salon",
        "tattoo", "perfumery", "hairdresser_supply", "spa", "sauna",
        "fitness_centre", "sports_centre", "dance",
    ],
    "auto": [
        "car_repair", "car", "car_parts", "tyres", "motorcycle",
        "motorcycle_repair", "bicycle", "driving_school", "car_wash",
        "fuel", "caravan", "boat",
    ],
    "batiment": [
        "electrician", "plumber", "carpenter", "roofer", "painter",
        "builder", "hvac", "tiler", "plasterer", "glaziery", "locksmith",
        "metal_construction", "stonemason", "window_construction", "gardener",
        "doors", "flooring", "insulation", "scaffolder", "sawmill",
        "hardware", "doityourself", "paint", "trade", "garden_centre",
    ],
    "sante": [
        "pharmacy", "dentist", "doctors", "veterinary", "optician",
        "hearing_aids", "medical_supply", "physiotherapist", "psychotherapist",
        "podiatrist", "nutrition_counselling", "midwife", "chiropractor",
        "clinic", "herbalist",
    ],
    "commerce": [
        "clothes", "shoes", "jewelry", "gift", "florist", "furniture",
        "books", "toys", "sports", "electronics", "mobile_phone", "pet",
        "second_hand", "antiques", "art", "music", "musical_instrument",
        "photo", "houseware", "interior_decoration", "kitchen", "bed",
        "curtain", "fabric", "wool", "watches", "bag", "boutique",
        "department_store", "variety_store", "supermarket", "convenience",
        "newsagent", "stationery", "tobacco", "e-cigarette", "outdoor",
        "fishing", "hunting", "video_games", "computer", "copyshop",
    ],
    "services": [
        "lawyer", "accountant", "insurance", "estate_agent", "travel_agency",
        "notary", "financial", "employment_agency", "it", "consulting",
        "advertising_agency", "architect", "surveyor", "tax_advisor",
        "laundry", "dry_cleaning", "funeral_directors", "photographer",
        "shoe_repair", "tailor", "watchmaker", "key_cutting", "cleaning",
        "moving_company", "storage_rental", "veterinary_office", "pet_grooming",
        "childcare", "kindergarten", "bank", "car_rental", "internet_cafe",
    ],
}

# Palette + identite visuelle par secteur.
THEMES = {
    "horeca": {
        "primaire": "#B2402F",
        "sombre": "#2A1512",
        "accent": "#E0A83C",
        "clair": "#FBF5EE",
        "police_titre": "'Georgia', 'Times New Roman', serif",
        "libelle": "Restaurant / Commerce de bouche",
        "accroche": "Une cuisine qui se raconte, une table dont on se souvient",
        "services": [
            ("Notre carte", "Des produits frais, travaillés chaque jour sur place."),
            ("Réservation", "Réservez votre table en un appel, on garde la meilleure."),
            ("À emporter", "Vos plats préparés, prêts à l'heure que vous choisissez."),
        ],
    },
    "beaute": {
        "primaire": "#8E5BA6",
        "sombre": "#241A2B",
        "accent": "#D9A7B0",
        "clair": "#FAF6F8",
        "police_titre": "'Georgia', 'Times New Roman', serif",
        "libelle": "Beauté / Bien-être",
        "accroche": "Prenez rendez-vous avec vous-même",
        "services": [
            ("Prestations", "Coupe, couleur, soin : un résultat pensé pour vous."),
            ("Sur rendez-vous", "Un créneau réservé rien que pour vous, sans attente."),
            ("Conseils", "Les bons gestes et les bons produits, pour que ça tienne."),
        ],
    },
    "auto": {
        "primaire": "#1F6FB2",
        "sombre": "#111A22",
        "accent": "#F0A202",
        "clair": "#F2F5F8",
        "police_titre": "'Helvetica Neue', Arial, sans-serif",
        "libelle": "Automobile / Mobilité",
        "accroche": "Votre véhicule entre de bonnes mains",
        "services": [
            ("Entretien", "Vidange, freins, distribution : tout est contrôlé."),
            ("Réparation", "Diagnostic clair, devis avant intervention, sans surprise."),
            ("Dépannage", "Un souci sur la route ? On décroche et on vous dépanne."),
        ],
    },
    "batiment": {
        "primaire": "#D97706",
        "sombre": "#1C1A17",
        "accent": "#2F6F4F",
        "clair": "#F7F4F0",
        "police_titre": "'Helvetica Neue', Arial, sans-serif",
        "libelle": "Bâtiment / Artisanat",
        "accroche": "Du travail bien fait, du devis à la finition",
        "services": [
            ("Devis gratuit", "On passe, on mesure, vous recevez un prix clair."),
            ("Réalisation", "Un chantier propre, tenu dans les délais annoncés."),
            ("Garantie", "Un travail assuré et un suivi après intervention."),
        ],
    },
    "sante": {
        "primaire": "#0E8C7A",
        "sombre": "#12211F",
        "accent": "#5EBFA8",
        "clair": "#F3F9F7",
        "police_titre": "'Helvetica Neue', Arial, sans-serif",
        "libelle": "Santé / Paramédical",
        "accroche": "À votre écoute, près de chez vous",
        "services": [
            ("Consultations", "Un accueil attentif et le temps qu'il faut."),
            ("Rendez-vous", "Des créneaux souples, y compris en dehors des heures de bureau."),
            ("Conseils", "Des réponses claires à vos questions, sans jargon."),
        ],
    },
    "commerce": {
        "primaire": "#2E5A88",
        "sombre": "#141B23",
        "accent": "#C4703B",
        "clair": "#F5F7FA",
        "police_titre": "'Helvetica Neue', Arial, sans-serif",
        "libelle": "Commerce de détail",
        "accroche": "Une sélection choisie, et le conseil qui va avec",
        "services": [
            ("En boutique", "Une sélection renouvelée, visible et disponible tout de suite."),
            ("Conseil", "On prend le temps de comprendre ce que vous cherchez."),
            ("Commande", "Pas en stock ? On le commande pour vous."),
        ],
    },
    "services": {
        "primaire": "#3F4A5A",
        "sombre": "#15181D",
        "accent": "#C89B3C",
        "clair": "#F6F7F9",
        "police_titre": "'Georgia', 'Times New Roman', serif",
        "libelle": "Services aux particuliers et entreprises",
        "accroche": "Un interlocuteur, des réponses claires",
        "services": [
            ("Notre métier", "Un accompagnement précis, du premier contact au résultat."),
            ("Sur mesure", "Chaque dossier est traité selon votre situation."),
            ("Réactivité", "Une réponse rapide, des délais annoncés et tenus."),
        ],
    },
}

SECTEUR_DEFAUT = "services"

# Mots-cles francais, pour les categories Google et les noms d'enseigne.
MOTS_FR = {
    "horeca": ["restaurant", "brasserie", "taverne", "friterie", "friture",
               "snack", "pizzeria", "boulangerie", "patisserie", "boucherie",
               "traiteur", "chocolaterie", "glacier", "cafe", "bistrot",
               "creperie", "sandwicherie", "caviste", "vins", "hotel",
               "chambre", "epicerie", "fromagerie", "poissonnerie"],
    "beaute": ["coiffure", "coiffeur", "barbier", "beaute", "esthetique",
               "institut", "onglerie", "massage", "spa", "parfumerie",
               "tatouage", "bien-etre", "fitness", "sport", "salle"],
    "auto": ["garage", "carrosserie", "pneus", "auto", "automobile", "moto",
             "velo", "cycles", "depannage", "car-wash", "lavage",
             "auto-ecole", "concession"],
    "batiment": ["toiture", "couvreur", "toitures", "electricite", "electricien",
                 "plomberie", "plombier", "chauffage", "menuiserie",
                 "menuisier", "peinture", "peintre", "macon", "maconnerie",
                 "carrelage", "renovation", "construction", "jardin",
                 "paysagiste", "terrassement", "sanitaire", "chassis",
                 "isolation", "serrurerie"],
    "sante": ["pharmacie", "dentiste", "dentaire", "medecin", "docteur",
              "kine", "kinesitherapie", "veterinaire", "opticien", "optique",
              "podologue", "psychologue", "logopede", "infirmier",
              "audition", "orthopedie"],
    "commerce": ["boutique", "magasin", "librairie", "fleuriste", "bijouterie",
                 "chaussures", "vetements", "mode", "meubles", "decoration",
                 "quincaillerie", "jouets", "informatique", "telephonie",
                 "animalerie", "papeterie", "tabac", "friperie", "brocante"],
    "services": ["comptable", "fiduciaire", "avocat", "notaire", "assurance",
                 "immobilier", "agence", "architecte", "photographe",
                 "nettoyage", "pressing", "blanchisserie", "pompes",
                 "funeraire", "toilettage", "demenagement", "voyage",
                 "consultant", "formation"],
}

# Index inverse, construit une fois.
_INDEX = {}
for _secteur, _valeurs in MAPPING.items():
    for _v in _valeurs:
        _INDEX[_v] = _secteur

_INDEX_FR = {}
for _secteur, _valeurs in MOTS_FR.items():
    for _v in _valeurs:
        _INDEX_FR[_v] = _secteur


def _jetons(valeur):
    """Decoupe une valeur en mots comparables, accents retires."""
    sans_accent = unicodedata.normalize("NFKD", str(valeur))
    sans_accent = sans_accent.encode("ascii", "ignore").decode("ascii").lower()
    return [j for j in re.split(r"[^a-z0-9]+", sans_accent) if j], sans_accent


def detecter_secteur(prospect):
    """Deduit le secteur depuis les tags OSM puis la categorie Google.

    La comparaison se fait par mot entier : un simple 'in' ferait par exemple
    correspondre 'it' (informatique) a l'interieur de 'fitness_centre'.
    """
    # 1. Valeur OSM exacte, la plus fiable.
    cle_osm = str(prospect.get("categorie_osm", "")).strip().lower()
    if cle_osm in _INDEX:
        return _INDEX[cle_osm]

    # 2. Mots-cles francais dans la categorie Google puis le nom.
    for champ in ("categorie_google", "nom"):
        jetons, complet = _jetons(prospect.get(champ, ""))
        for jeton in jetons:
            if jeton in _INDEX_FR:
                return _INDEX_FR[jeton]
        for mot, secteur in _INDEX_FR.items():
            if "-" in mot and mot in complet:
                return secteur

    # 3. Dernier recours : mot entier d'une valeur OSM composee.
    jetons, _ = _jetons(cle_osm)
    for jeton in jetons:
        if jeton in _INDEX:
            return _INDEX[jeton]

    return SECTEUR_DEFAUT


def theme(secteur):
    return THEMES.get(secteur, THEMES[SECTEUR_DEFAUT])
