"""
IneArt - Standard telephonique IA.

Recoit un vrai appel via Twilio, mene le dialogue en francais, structure la
demande avec Claude et envoie la fiche de devis par mail.

    Appelant  ->  Twilio  ->  ce serveur  ->  Claude  ->  mail SMTP

La reconnaissance vocale et la voix de synthese sont fournies par Twilio :
aucune API vocale supplementaire a payer.

Variables d'environnement :
    ANTHROPIC_API_KEY   obligatoire
    MAIL_USER           compte SMTP d'envoi   (identique au projet devis)
    MAIL_PASS           mot de passe SMTP
    MAIL_DEST           destinataire des fiches (defaut : MAIL_USER)
    SMTP_HOST           defaut send.one.com
    SMTP_PORT           defaut 465 (SSL)
    METIER              "artisan", "toiture", "plomberie"... sert au discours
    SOCIETE             nom annonce au decroche (defaut : IneArt)

Lancement local :
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=... MAIL_USER=... MAIL_PASS=...
    python standard.py
    cloudflared tunnel --url http://localhost:5000     # URL publique gratuite

Puis dans Twilio : Phone Numbers > votre numero > A call comes in >
Webhook POST  https://<votre-tunnel>/appel
"""

import datetime as dt
import json
import os
import smtplib
import ssl
import threading
from email.message import EmailMessage

import anthropic
from flask import Flask, Response, request

MODELE = "claude-haiku-4-5-20251001"
SOCIETE = os.environ.get("SOCIETE", "IneArt")
METIER = os.environ.get("METIER", "artisan")
MAX_TOURS = 12

SMTP_HOST = os.environ.get("SMTP_HOST", "send.one.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))
MAIL_USER = os.environ.get("MAIL_USER", "")
MAIL_PASS = os.environ.get("MAIL_PASS", "")
MAIL_DEST = os.environ.get("MAIL_DEST", MAIL_USER)

# Voix Twilio francaise. Alternatives : Polly.Lea-Neural, Polly.Remi-Neural.
VOIX = os.environ.get("VOIX_TWILIO", "Polly.Lea-Neural")

CONSIGNE = """Tu es la standardiste telephonique de {societe}, {metier} en region de Huy.
Un client appelle pour demander un devis. Tu menes la conversation a l'oral.

REGLES DE PAROLE
- Une seule question a la fois, jamais deux.
- Phrases courtes, ton chaleureux et naturel. Ta reponse est lue a voix haute :
  aucune liste, aucune puce, aucun emoji, aucune mise en forme.
- Ne demande jamais le numero de telephone : il est deja connu.
- Si le client est vague, relance une fois puis passe a la suite.
- Si le client s'impatiente ou demande un humain, conclus poliment en promettant
  un rappel rapide.
- Reformule brievement a la fin pour confirmer.

INFORMATIONS A RECUEILLIR, dans cet ordre
1. Le nom du client
2. La nature des travaux, avec un minimum de detail
3. L'adresse ou au moins la commune
4. L'urgence : depannage immediat, cette semaine, ou pas presse
5. Les disponibilites pour une visite

REPONSE ATTENDUE
Reponds toujours en JSON strict, sans aucun texte autour :
{{"dire": "ce que tu dis a voix haute", "termine": false, "fiche": null}}

Quand tout est recueilli et confirme :
{{"dire": "conclusion et au revoir", "termine": true,
  "fiche": {{"nom": "", "travaux": "", "resume": "", "adresse": "",
            "urgence": "faible|moyenne|elevee", "disponibilites": "",
            "rappel_humain": false}}}}

Au premier tour, l'utilisateur envoie "[DECROCHE]" : presente-toi au nom de
{societe} et demande en quoi tu peux aider."""

URGENCES = {
    "faible": "Pas presse",
    "moyenne": "Cette semaine",
    "elevee": "URGENT - depannage",
}

application = Flask(__name__)
client_ia = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

# Conversations en cours, indexees par identifiant d'appel Twilio.
# Suffisant tant qu'un seul processus tourne ; passer a Redis au-dela.
appels = {}
verrou = threading.Lock()


# --- outils -----------------------------------------------------------------

def echapper(texte):
    """Echappe le texte insere dans du XML TwiML."""
    return (str(texte or "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def twiml(contenu):
    return Response('<?xml version="1.0" encoding="UTF-8"?><Response>'
                    + contenu + "</Response>",
                    mimetype="text/xml")


def dire_et_ecouter(phrase, action="/tour"):
    """Fait parler l'assistante puis ecoute la reponse du client."""
    return twiml(
        '<Gather input="speech" language="fr-FR" speechTimeout="auto" '
        'action="%s" method="POST" actionOnEmptyResult="true">'
        '<Say voice="%s" language="fr-FR">%s</Say>'
        '</Gather>' % (action, VOIX, echapper(phrase))
    )


def dire_et_raccrocher(phrase):
    return twiml('<Say voice="%s" language="fr-FR">%s</Say><Hangup/>'
                 % (VOIX, echapper(phrase)))


def reflechir(historique):
    """Interroge Claude et renvoie le dictionnaire de reponse."""
    reponse = client_ia.messages.create(
        model=MODELE,
        max_tokens=400,
        system=CONSIGNE.format(societe=SOCIETE, metier=METIER),
        messages=historique,
    )
    brut = "".join(bloc.text for bloc in reponse.content if bloc.type == "text").strip()
    debut, fin = brut.find("{"), brut.rfind("}")
    try:
        return json.loads(brut[debut:fin + 1])
    except (ValueError, json.JSONDecodeError):
        return {"dire": brut, "termine": False, "fiche": None}


# --- envoi de la fiche ------------------------------------------------------

def transcription(historique):
    lignes = []
    for tour in historique:
        if tour["content"] == "[DECROCHE]":
            continue
        if tour["role"] == "user":
            lignes.append("Client      : " + tour["content"])
        else:
            try:
                lignes.append("Assistante  : " + json.loads(tour["content"])["dire"])
            except (ValueError, KeyError):
                lignes.append("Assistante  : " + tour["content"])
    return "\n".join(lignes)


def corps_du_mail(fiche, numero, historique, debut):
    duree = int((dt.datetime.now() - debut).total_seconds())
    lignes = [
        ("Client", fiche.get("nom")),
        ("Telephone", numero),
        ("Travaux", fiche.get("resume") or fiche.get("travaux")),
        ("Adresse", fiche.get("adresse")),
        ("Urgence", URGENCES.get(fiche.get("urgence"), fiche.get("urgence"))),
        ("Disponibilites", fiche.get("disponibilites")),
        ("Appel recu a", debut.strftime("%H:%M le %d/%m/%Y")),
        ("Duree", "%d min %02d s" % (duree // 60, duree % 60)),
    ]
    entete = "\n".join("%-16s : %s" % (cle, valeur)
                       for cle, valeur in lignes if valeur)
    if fiche.get("rappel_humain"):
        entete = "!! Le client a demande a parler a une personne !!\n\n" + entete
    return entete + "\n\n--- Transcription de l'appel ---\n" + transcription(historique)


def envoyer_fiche(fiche, numero, historique, debut):
    """Envoie la fiche par mail. Les erreurs ne doivent jamais couper l'appel."""
    if not (MAIL_USER and MAIL_PASS):
        print("! SMTP non configure, fiche non envoyee :\n%s"
              % corps_du_mail(fiche, numero, historique, debut))
        return

    urgence = fiche.get("urgence") == "elevee"
    message = EmailMessage()
    message["Subject"] = "%s%s - %s" % (
        "[URGENT] " if urgence else "",
        fiche.get("travaux") or "Demande de devis",
        fiche.get("nom") or numero)
    message["From"] = "%s <%s>" % (SOCIETE, MAIL_USER)
    message["To"] = MAIL_DEST
    message["Reply-To"] = MAIL_USER
    message.set_content(corps_du_mail(fiche, numero, historique, debut))

    try:
        contexte = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=contexte) as serveur:
            serveur.login(MAIL_USER, MAIL_PASS)
            serveur.send_message(message)
        print("Fiche envoyee a %s" % MAIL_DEST)
    except Exception as erreur:  # noqa: BLE001
        print("! envoi du mail impossible (%s) : %s" % (type(erreur).__name__, erreur))


# --- webhooks Twilio --------------------------------------------------------

@application.route("/appel", methods=["POST"])
def appel():
    """Premier contact : Twilio vient de recevoir un appel entrant."""
    identifiant = request.form.get("CallSid", "inconnu")
    numero = request.form.get("From", "inconnu")
    print("Appel entrant de %s (%s)" % (numero, identifiant))

    historique = [{"role": "user", "content": "[DECROCHE]"}]
    reponse = reflechir(historique)
    historique.append({"role": "assistant", "content": json.dumps(reponse)})

    with verrou:
        appels[identifiant] = {
            "numero": numero,
            "historique": historique,
            "debut": dt.datetime.now(),
            "tours": 0,
        }

    # Mention RGPD : l'appelant doit savoir qu'il parle a un assistant.
    annonce = ("Bonjour, vous etes en relation avec l'assistant vocal de %s. "
               "%s" % (SOCIETE, reponse["dire"]))
    return dire_et_ecouter(annonce)


@application.route("/tour", methods=["POST"])
def tour():
    """Le client vient de parler : on poursuit le dialogue."""
    identifiant = request.form.get("CallSid", "inconnu")
    dit = (request.form.get("SpeechResult") or "").strip()

    with verrou:
        etat = appels.get(identifiant)
    if not etat:
        return dire_et_raccrocher("Desolee, la communication a ete interrompue. "
                                  "Rappelez-nous, nous restons a votre disposition.")

    etat["tours"] += 1
    if etat["tours"] > MAX_TOURS:
        return conclure(identifiant, etat,
                        "Je vous remercie, je transmets votre demande. "
                        "Vous serez rappele tres vite. Bonne journee.")

    if not dit:
        return dire_et_ecouter("Je ne vous ai pas entendu. Pouvez-vous repeter ?")

    print("  client : %s" % dit)
    etat["historique"].append({"role": "user", "content": dit})
    reponse = reflechir(etat["historique"])
    etat["historique"].append({"role": "assistant", "content": json.dumps(reponse)})
    print("  assistante : %s" % reponse["dire"])

    if reponse.get("termine"):
        fiche = reponse.get("fiche") or {}
        envoyer_fiche(fiche, etat["numero"], etat["historique"], etat["debut"])
        return conclure(identifiant, etat, reponse["dire"])

    return dire_et_ecouter(reponse["dire"])


def conclure(identifiant, etat, phrase):
    with verrou:
        appels.pop(identifiant, None)
    return dire_et_raccrocher(phrase)


@application.route("/statut", methods=["POST"])
def statut():
    """Fin d'appel signalee par Twilio : on envoie ce qu'on a si le client
    a raccroche avant la fin du questionnaire."""
    identifiant = request.form.get("CallSid", "")
    with verrou:
        etat = appels.pop(identifiant, None)
    if etat and etat["tours"] > 1:
        print("Raccroche premature, envoi de la fiche partielle")
        envoyer_fiche({"nom": "", "travaux": "Appel interrompu",
                       "resume": "Le client a raccroche avant la fin",
                       "urgence": "moyenne"},
                      etat["numero"], etat["historique"], etat["debut"])
    return ("", 204)


@application.route("/", methods=["GET"])
def accueil():
    with verrou:
        actifs = len(appels)
    return ("Standard IA %s - operationnel. %d appel(s) en cours.\n"
            "Webhook a declarer dans Twilio : POST /appel\n" % (SOCIETE, actifs))


if __name__ == "__main__":
    application.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
