"""
IneArt — Analyseur de mails & Récap par email
Version 2 : sauvegarde mails.json dans le dépôt GitHub pour le dashboard
"""

import imaplib
import email
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
import os
import re
import sys
import smtplib
from datetime import datetime, timedelta
import anthropic

IMAP_HOST  = "imap.one.com"
IMAP_PORT  = 993
SMTP_HOST  = "send.one.com"
SMTP_PORT  = 587

MAIL_USER  = os.environ["MAIL_USER"]
MAIL_PASS  = os.environ["MAIL_PASS"]
CLAUDE_KEY = os.environ["ANTHROPIC_API_KEY"]
CLAUDE_MODEL = "claude-haiku-4-5-20251001"
JOURS_RETOUR = 3

# Fichier de sortie pour le dashboard
OUTPUT_FILE = "mails.json"
MAX_MAILS_SAVED = 100  # garder les 50 derniers mails dans le JSON


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
    print(f"[IMAP] Connexion a {IMAP_HOST}...")
    mails = []
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
        imap.login(MAIL_USER, MAIL_PASS)
        imap.select("INBOX")
        depuis = (datetime.now() - timedelta(days=JOURS_RETOUR)).strftime("%d-%b-%Y")
        _, ids = imap.search(None, f'(SINCE "{depuis}")')
        ids_liste = ids[0].split()
        print(f"[IMAP] {len(ids_liste)} mail(s) trouvé(s) (lus + non lus).")
        for uid in ids_liste:
            # BODY.PEEK = lit sans marquer comme lu
            _, data = imap.fetch(uid, "(BODY.PEEK[])")
            msg = email.message_from_bytes(data[0][1])
            mails.append({
                "uid":   uid.decode(),
                "de":    decoder_header(msg.get("From", "")),
                "sujet": decoder_header(msg.get("Subject", "")),
                "date":  msg.get("Date", ""),
                "corps": extraire_texte(msg),
            })
    return mails


PROMPT_SYSTEME = """Tu es un assistant commercial pour IneArt, entreprise belge de broderie et personnalisation textile (serigraphie, broderie, DTF, flocage).
Analyse l'email et reponds UNIQUEMENT en JSON valide, sans markdown ni texte autour.

REGLES DE CLASSIFICATION :
- DEMANDE_DEVIS : le client demande un prix, un devis, une offre, parle de quantites, de broderie, impression, textiles, vetements personnalises. Inclus aussi les reponses a des devis (RE: Devis), relances, demandes de precisions sur un devis existant.
- COMMANDE_VALIDEE : le client confirme une commande, valide un devis, demande la production, envoie un bon de commande.
- AUTRE : newsletters, notifications automatiques, factures fournisseurs, emails de livraison (bpost, PostNL), reseaux sociaux, GitHub, PayPal, Anthropic, publicites, confirmations de retour colis. En cas de doute entre DEVIS et AUTRE, choisis DEVIS.

Format exact :
{
  "type": "DEMANDE_DEVIS",
  "confiance": 0.9,
  "client": {"nom": "...", "email": "...", "telephone": null, "entreprise": null},
  "articles": [{"description": "...", "quantite": 1, "prix_unitaire": null, "notes": null}],
  "delai_souhaite": null,
  "notes_commerciales": "resume en 1-2 phrases"
}"""


EXPEDITEURS_IGNORES = [
    "github.com", "anthropic.com", "paypal.be", "paypal.com",
    "bpost.be", "postnl.nl", "inpost-group.com", "zalando.be",
    "facebook.com", "business-updates.facebook.com", "relevanceai.com",
    "bigmat.be", "irobot.com", "link.com", "email.claude.com",
    "alun.dk", "q8.com", "odoo.com",
]

def est_expediteur_ignore(expediteur):
    exp = expediteur.lower()
    return any(domaine in exp for domaine in EXPEDITEURS_IGNORES)



def analyser_mail(mail):
    client = anthropic.Anthropic(api_key=CLAUDE_KEY)
    contenu = f"De : {mail['de']}\nSujet : {mail['sujet']}\nDate : {mail['date']}\n\n{mail['corps'][:3000]}"
    response = client.messages.create(
        model=CLAUDE_MODEL, max_tokens=1024,
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
    """Charge le fichier mails.json existant pour y ajouter les nouveaux."""
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("mails", [])
        except Exception:
            pass
    return []


def sauvegarder_json(nouveaux_mails: list):
    """Fusionne avec l'historique et sauvegarde mails.json."""
    historique = charger_historique()

    # Éviter les doublons par uid
    uids_existants = {m.get("uid") for m in historique}
    a_ajouter = [m for m in nouveaux_mails if m.get("uid") not in uids_existants]

    # Fusionner et garder les MAX_MAILS_SAVED derniers
    tous = a_ajouter + historique
    tous = tous[:MAX_MAILS_SAVED]

    data = {
        "mis_a_jour": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total": len(tous),
        "mails": tous
    }

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"[JSON] mails.json sauvegarde : {len(tous)} mails ({len(a_ajouter)} nouveaux)")
    return len(a_ajouter)


def construire_html(devis_list: list) -> str:
    date_str = datetime.now().strftime("%d/%m/%Y a %H:%M")
    blocs = ""
    for i, d in enumerate(devis_list, 1):
        a  = d["analyse"]
        m  = d["mail"]
        cl = a.get("client", {})
        arts = a.get("articles", [])
        type_m = a.get("type", "")
        couleur = "#2ecc71" if type_m == "DEMANDE_DEVIS" else "#f39c12"
        label   = "Devis" if type_m == "DEMANDE_DEVIS" else "Commande"
        rows = ""
        for art in arts:
            prix = f"{art['prix_unitaire']} EUR" if art.get("prix_unitaire") else "A definir"
            note = f" ({art['notes']})" if art.get("notes") else ""
            rows += f"<tr><td style='padding:6px 10px;border-bottom:1px solid #eee'>{art.get('description','?')}{note}</td><td style='padding:6px 10px;border-bottom:1px solid #eee;text-align:center'>{art.get('quantite','?')}</td><td style='padding:6px 10px;border-bottom:1px solid #eee;text-align:center'>{prix}</td></tr>"
        if not rows:
            rows = "<tr><td colspan='3' style='padding:8px;color:#888'>Voir mail original</td></tr>"
        notes_com = a.get("notes_commerciales", "")
        delai = a.get("delai_souhaite") or "Non precise"
        tel = cl.get("telephone", "") or ""
        blocs += f"""
<div style='background:#fff;border:1px solid #ddd;border-radius:8px;margin-bottom:20px;overflow:hidden'>
  <div style='background:#1a1a2e;padding:14px 20px'>
    <span style='color:#fff;font-weight:bold;font-size:15px'>Demande #{i} — {cl.get("nom","Client inconnu")}</span>
    <span style='background:{couleur};color:#fff;padding:3px 10px;border-radius:12px;font-size:12px;float:right'>{label}</span>
  </div>
  <div style='padding:12px 20px;background:#f8f9fa;border-bottom:1px solid #eee;font-size:13px'>
    <b>Client :</b> {cl.get("nom","—")}{" — " + cl.get("entreprise","") if cl.get("entreprise") else ""}<br>
    <b>Email :</b> {cl.get("email","—")}<br>
    {"<b>Tel :</b> " + tel + "<br>" if tel else ""}
    <b>Delai :</b> {delai}
  </div>
  {"<div style='padding:10px 20px;background:#fff8e1;border-bottom:1px solid #eee;font-size:13px'>💡 " + notes_com + "</div>" if notes_com else ""}
  <div style='padding:14px 20px'>
    <table style='width:100%;border-collapse:collapse;font-size:13px'>
      <thead><tr style='background:#f8f9fa'>
        <th style='padding:8px 10px;text-align:left'>Article</th>
        <th style='padding:8px 10px;text-align:center'>Qte</th>
        <th style='padding:8px 10px;text-align:center'>Prix</th>
      </tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
  <div style='padding:12px 20px;background:#f8f9fa;border-top:1px solid #eee;text-align:right'>
    <a href='https://devis-ineart.odoo.com/odoo/sales/new' style='background:#1a1a2e;color:#fff;padding:8px 18px;border-radius:6px;text-decoration:none;font-size:13px'>
      Creer le devis dans Odoo
    </a>
  </div>
</div>"""
    return f"""<!DOCTYPE html><html><head><meta charset='utf-8'></head>
<body style='font-family:Arial,sans-serif;background:#f0f2f5;padding:20px;margin:0'>
<div style='max-width:660px;margin:0 auto'>
  <div style='background:#1a1a2e;border-radius:8px 8px 0 0;padding:20px 24px;text-align:center'>
    <h1 style='color:#fff;margin:0;font-size:20px'>IneArt — Recap Devis</h1>
    <p style='color:#aaa;margin:6px 0 0;font-size:13px'>Genere le {date_str} · {len(devis_list)} demande(s)</p>
  </div>
  <div style='background:#f0f2f5;padding:20px 0'>{blocs}</div>
  <div style='background:#1a1a2e;border-radius:0 0 8px 8px;padding:14px 24px;text-align:center'>
    <p style='color:#aaa;margin:0;font-size:12px'>Dashboard : https://dashboard.ineart.be</p>
  </div>
</div></body></html>"""


def envoyer_recap(devis_list: list):
    nb = len(devis_list)
    sujet = f"[IneArt AI] {nb} nouvelle(s) demande(s) — {datetime.now().strftime('%d/%m/%Y')}"
    msg = MIMEMultipart("alternative")
    msg["Subject"] = sujet
    msg["From"]    = MAIL_USER
    msg["To"]      = MAIL_USER
    msg.attach(MIMEText(construire_html(devis_list), "html", "utf-8"))
    print(f"[SMTP] Envoi recap...")
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(MAIL_USER, MAIL_PASS)
        smtp.sendmail(MAIL_USER, MAIL_USER, msg.as_bytes())
    print("[SMTP] Email envoye !")


def main(dry_run=False):
    print(f"\n{'='*50}")
    print(f"IneArt — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"Mode : {'TEST' if dry_run else 'NORMAL'}")
    print(f"{'='*50}\n")

    mails_bruts = lire_mails_imap()
    if not mails_bruts:
        print("Aucun nouveau mail non lu.")
        # Sauvegarder quand même le JSON (avec l'historique)
        sauvegarder_json([])
        return

    devis_detectes = []
    mails_structures = []

    for mail in mails_bruts:
        print(f"\n📧 {mail['de']}\n   {mail['sujet']}")

        # Filtrer les expéditeurs parasites sans appeler Claude
        if est_expediteur_ignore(mail['de']):
            print(f"   → IGNORÉ (expéditeur système)")
            mail_dash = {
                "uid": mail["uid"], "de": mail["de"], "sujet": mail["sujet"],
                "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "type": "A", "cf": 99,
                "cl": {"nom": "", "email": "", "ent": "", "tel": ""},
                "arts": [], "delai": "", "note": "Expéditeur automatique filtré.",
            }
            mails_structures.append(mail_dash)
            continue

        analyse = analyser_mail(mail)
        type_m  = analyse.get("type", "AUTRE")
        conf    = analyse.get("confiance", 0)
        print(f"   → {type_m} ({conf:.0%})")

        # Structurer pour le dashboard
        mail_dash = {
            "uid":     mail["uid"],
            "de":      mail["de"],
            "sujet":   mail["sujet"],
            "date":    datetime.now().strftime("%d/%m/%Y %H:%M"),
            "type":    type_m[0] if type_m else "A",  # D, C, ou A
            "cf":      round(conf * 100),
            "cl": {
                "nom":  analyse.get("client", {}).get("nom", ""),
                "email":analyse.get("client", {}).get("email", ""),
                "ent":  analyse.get("client", {}).get("entreprise", ""),
                "tel":  analyse.get("client", {}).get("telephone", ""),
            },
            "arts":    [{"d": a.get("description",""), "q": a.get("quantite",1), "n": a.get("notes","")} for a in analyse.get("articles", [])],
            "delai":   analyse.get("delai_souhaite", ""),
            "note":    analyse.get("notes_commerciales", ""),
        }
        mails_structures.append(mail_dash)

        if type_m in ("DEMANDE_DEVIS", "COMMANDE_VALIDEE") and conf >= 0.6:
            devis_detectes.append({"mail": mail, "analyse": analyse})

    # Sauvegarder le JSON pour le dashboard
    if not dry_run:
        sauvegarder_json(mails_structures)

    print(f"\n{'='*50}")
    print(f"Resultat : {len(devis_detectes)} devis sur {len(mails_bruts)} mail(s)")

    if devis_detectes:
        if dry_run:
            print("\n[TEST] Donnees :")
            for d in devis_detectes:
                print(json.dumps(d["analyse"], indent=2, ensure_ascii=False))
        else:
            envoyer_recap(devis_detectes)

    print(f"{'='*50}\n")


if __name__ == "__main__":
    main(dry_run="--dry-run" in sys.argv)
