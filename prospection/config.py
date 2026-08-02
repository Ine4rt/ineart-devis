"""
IneArt - Prospection Huy : configuration centrale.

Tout ce qui se regle a la main est ici : tarifs, duree de validite des demos,
coordonnees IneArt, perimetre geographique.
"""

import os

# --- Agence -----------------------------------------------------------------
AGENCE = {
    "nom": "IneArt",
    "email": os.environ.get("INEART_EMAIL", "info@ineart.be"),
    "telephone": os.environ.get("INEART_TEL", "+32 471 00 00 00"),
    "site": os.environ.get("INEART_SITE", "https://www.ineart.be"),
    "signature": "Dimitri - IneArt",
    "ville": "Huy",
}

# --- Tarifs affiches sur les demos et dans les mails -------------------------
TARIFS = {
    "site": {
        "libelle": "Site internet complet",
        "prix": 250,
        "unite": "une seule fois",
        "detail": [
            "Design sur mesure, adapté à votre métier",
            "Version mobile, tablette et ordinateur",
            "Vos photos, vos horaires, vos coordonnées",
            "Formulaire de contact directement dans votre boîte mail",
            "Mise en ligne comprise",
        ],
    },
    "domaine": {
        "libelle": "Nom de domaine + hébergement",
        "prix": 50,
        "unite": "par an",
        "detail": [
            "Votre adresse personnalisée (ex. votre-nom.be)",
            "Hébergement sécurisé en HTTPS",
            "Sauvegardes et maintenance technique",
        ],
    },
}

# --- Duree de vie des demos --------------------------------------------------
DUREE_DEMO_JOURS = int(os.environ.get("DUREE_DEMO_JOURS", "8"))

# --- Perimetre de collecte ---------------------------------------------------
# Communes OSM (admin_level 8 en Belgique).
COMMUNES = os.environ.get("COMMUNES", "Huy").split(",")

# Nombre max de prospects traites par run (0 = illimite).
LIMITE_PROSPECTS = int(os.environ.get("LIMITE_PROSPECTS", "20"))

# --- FTP ---------------------------------------------------------------------
# Memes noms de secrets que le projet ineweb.
FTP = {
    "host": os.environ.get("FTP_HOST", ""),
    "user": os.environ.get("FTP_USER", ""),
    "password": os.environ.get("FTP_PASS", ""),
    # Dossier distant racine des demos.
    "dossier": os.environ.get("FTP_DIR", "/demos"),
    # URL publique correspondant a FTP_DIR (sert a construire les liens).
    "url_publique": os.environ.get("DEMO_BASE_URL", "https://www.ineart.be/demos"),
    "tls": os.environ.get("FTP_TLS", "1") == "1",
}

# --- Chemins -----------------------------------------------------------------
RACINE = os.path.dirname(os.path.abspath(__file__))
DOSSIER_DATA = os.path.join(RACINE, "data")
DOSSIER_SITES = os.path.join(RACINE, "sites")
DOSSIER_MAILS = os.path.join(RACINE, "mails")
DOSSIER_TEMPLATES = os.path.join(RACINE, "templates")

FICHIER_PROSPECTS = os.path.join(DOSSIER_DATA, "prospects.json")
FICHIER_EXEMPLE = os.path.join(DOSSIER_DATA, "exemple_prospects.json")
FICHIER_MANIFEST = os.path.join(DOSSIER_DATA, "manifest.json")
FICHIER_ENVOIS = os.path.join(DOSSIER_MAILS, "envois.csv")
