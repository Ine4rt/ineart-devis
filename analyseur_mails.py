import imaplib
import email
from email.header import decode_header
import json
import os
import re
import sys
from datetime import datetime, timedelta
import anthropic

IMAP_HOST = "imap.one.com"
IMAP_PORT = 993

MAIL_USER = os.environ["MAIL_USER"]
MAIL_PASS = os.environ["MAIL_PASS"]
CLAUDE_KEY = os.environ["ANTHROPIC_API_KEY"]
CLAUDE_MODEL = "claude-haiku-4-5-20251001"
JOURS_RETOUR = 1
OUTPUT_FILE = "mails.json"
MAX_MAILS = 50


def decoder_header(valeur):
    parties = decode_header(valeur or "")
    result = []
    for data, charset in parties:
        if isinstance(data, bytes):
            result.append(data.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(data)
    return " ".join(result)


def extraire_texte(msg):
    texte = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                charset = part.get_content_charset() or "utf-8"
                try:
                    texte += part.get_payload(decode=True).decode(charset, errors="replace")
                except Exception:
                    pass
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            texte = msg.get_payload(decode=True).decode(charset, errors="replace")
        except Exception:
            texte = ""
    return texte.strip()


def lire_mails_imap():
    print("[IMAP] Connexion a imap.one.com...")
    mails = []
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
        imap.login(MAIL_USER, MAIL_PASS)
        imap.select("INBOX")
        depuis = (datetime.now() - timedelta(days=JOURS_RETOUR)).strftime("%d-%b-%Y")
        _, ids = imap.search(None, "(UNSEEN SINCE \"" + depuis + "\")")
        ids_liste = ids[0].split()
        print("[IMAP] " + str(len(ids_liste)) + " mail(s) non lu(s).")
        for uid in ids_liste:
            _, data = imap.fetch(uid, "(BODY.PEEK[])")
            msg = email.message_from_bytes(data[0][1])
            mails.append({
                "uid": uid.decode(),
                "de": decoder_header(msg.get("From", "")),
                "sujet": decoder_header(msg.get("Subject", "")),
                "date": msg.get("Date", ""),
                "corps": extraire_texte(msg),
            })
    return mails


PROMPT_SYSTEME = """Tu es un assistant commercial pour IneArt, entreprise belge de broderie et personnalisation textile.
Analyse l'email et reponds UNIQUEMENT en JSON valide, sans markdown ni texte autour.
Types : DEMANDE_DEVIS, COMMANDE_VALIDEE, AUTRE.
Format :
{
  "type": "DEMANDE_DEVIS",
  "confiance": 0.9,
  "client": {"nom": "...", "email": "...", "telephone": null, "entreprise": null},
  "articles": [{"description": "...", "quantite": 1, "prix_unitaire": null, "notes": null}],
  "delai_souhaite": null,
  "notes_commerciales": "resume en 1-2 phrases"
}"""


def analyser_mail(mail):
    client = anthropic.Anthropic(api_key=CLAUDE_KEY)
    contenu = "De : " + mail["de"] + "\nSujet : " + mail["sujet"] + "\nDate : " + mail["date"] + "\n\n" + mail["corps"][:3000]
    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=1024,
        system=PROMPT_SYSTEME,
        messages=[{"role": "user", "content": contenu}],
    )
    texte = response.content[0].text.strip()
    texte = re.sub(r"^```json\s*", "", texte)
    texte = re.sub(r"\s*```$", "", texte)
    try:
        return json.loads(texte)
    except Exception:
        return {"type": "AUTRE", "confiance": 0, "articles": [], "client": {}}


def charger_historique():
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get("mails", [])
        except Exception:
            pass
    return []


def sauvegarder_json(nouveaux):
    historique = charger_historique()
    uids = {m.get("uid") for m in historique}
    a_ajouter = [m for m in nouveaux if m.get("uid") not in uids]
    tous = (a_ajouter + historique)[:MAX_MAILS]
    data = {
        "mis_a_jour": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total": len(tous),
        "mails": tous
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("[JSON] mails.json : " + str(len(tous)) + " mails (" + str(len(a_ajouter)) + " nouveaux)")


def main(dry_run=False):
    print("\n" + "="*50)
    print("IneArt - " + datetime.now().strftime("%d/%m/%Y %H:%M"))
    print("Mode : TEST" if dry_run else "Mode : NORMAL")
    print("="*50 + "\n")

    mails_bruts = lire_mails_imap()

    if not mails_bruts:
        print("Aucun nouveau mail.")
        if not dry_run:
            sauvegarder_json([])
        return

    mails_structures = []
    nb_devis = 0

    for mail in mails_bruts:
        print("\n📧 " + mail["de"] + "\n   " + mail["sujet"])
        analyse = analyser_mail(mail)
        type_m = analyse.get("type", "AUTRE")
        conf = analyse.get("confiance", 0)
        print("   -> " + type_m + " (" + str(round(conf * 100)) + "%)")

        if type_m == "DEMANDE_DEVIS":
            type_court = "D"
        elif type_m == "COMMANDE_VALIDEE":
            type_court = "C"
        else:
            type_court = "A"

        cl = analyse.get("client", {})

        mails_structures.append({
            "uid": mail["uid"],
            "de": mail["de"],
            "sujet": mail["sujet"],
            "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "type": type_court,
            "cf": round(conf * 100),
            "cl": {
                "nom": cl.get("nom", ""),
                "email": cl.get("email", ""),
                "ent": cl.get("entreprise", ""),
                "tel": cl.get("telephone", ""),
            },
            "arts": [{"d": a.get("description",""), "q": a.get("quantite",1), "n": a.get("notes","")} for a in analyse.get("articles", [])],
            "delai": analyse.get("delai_souhaite", ""),
            "note": analyse.get("notes_commerciales", ""),
        })

        if type_m in ("DEMANDE_DEVIS", "COMMANDE_VALIDEE") and conf >= 0.6:
            nb_devis += 1

    if not dry_run:
        sauvegarder_json(mails_structures)

    print("\n" + "="*50)
    print("Resultat : " + str(nb_devis) + " devis sur " + str(len(mails_bruts)) + " mail(s)")
    print("="*50 + "\n")


if __name__ == "__main__":
    main(dry_run="--dry-run" in sys.argv)
