import imaplib
import email
from email.header import decode_header
import json
import os
import re
import sys
import xmlrpc.client
from datetime import datetime, timedelta
import anthropic

IMAP_HOST  = "imap.one.com"
IMAP_PORT  = 993
IMAP_USER  = os.environ["MAIL_USER"]
IMAP_PASS  = os.environ["MAIL_PASS"]
ODOO_URL   = os.environ["ODOO_URL"]
ODOO_DB    = os.environ["ODOO_DB"]
ODOO_USER  = os.environ["ODOO_USER"]
ODOO_PASS  = os.environ["ODOO_PASS"]
CLAUDE_KEY   = os.environ["ANTHROPIC_API_KEY"]
CLAUDE_MODEL = "claude-haiku-4-5-20251001"
JOURS_RETOUR = 1

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
    print(f"[IMAP] Connexion à {IMAP_HOST}...")
    mails = []
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
        imap.login(IMAP_USER, IMAP_PASS)
        imap.select("INBOX")
        depuis = (datetime.now() - timedelta(days=JOURS_RETOUR)).strftime("%d-%b-%Y")
        _, ids = imap.search(None, f'(UNSEEN SINCE "{depuis}")')
        ids_liste = ids[0].split()
        print(f"[IMAP] {len(ids_liste)} mail(s) non lu(s).")
        for uid in ids_liste:
            _, data = imap.fetch(uid, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            mails.append({
                "uid":   uid.decode(),
                "de":    decoder_header(msg.get("From", "")),
                "sujet": decoder_header(msg.get("Subject", "")),
                "date":  msg.get("Date", ""),
                "corps": extraire_texte(msg),
            })
    return mails

PROMPT_SYSTEME = """Tu es un assistant commercial pour IneArt, entreprise belge de broderie et personnalisation textile.
Analyse l'email et réponds UNIQUEMENT en JSON valide, sans markdown ni texte autour.
Type : DEMANDE_DEVIS, COMMANDE_VALIDEE, ou AUTRE.
Format :
{
  "type": "DEMANDE_DEVIS",
  "confiance": 0.9,
  "client": {"nom": "...", "email": "...", "telephone": null, "entreprise": null},
  "articles": [{"description": "...", "quantite": 1, "prix_unitaire": null, "notes": null}],
  "delai_souhaite": null,
  "notes_commerciales": "..."
}"""

def analyser_mail(mail):
    client = anthropic.Anthropic(api_key=CLAUDE_KEY)
    contenu = f"De : {mail['de']}\nSujet : {mail['sujet']}\nDate : {mail['date']}\n\n{mail['corps'][:3000]}"
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

def connexion_odoo():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_PASS, {})
    if not uid:
        raise ConnectionError("Authentification Odoo échouée.")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    return uid, models

def trouver_ou_creer_client(uid, models, analyse):
    info  = analyse.get("client", {})
    email = info.get("email", "")
    nom   = info.get("nom") or info.get("entreprise") or "Client inconnu"
    if email:
        ids = models.execute_kw(ODOO_DB, uid, ODOO_PASS, "res.partner", "search", [[["email", "=", email]]])
        if ids:
            return ids[0]
    return models.execute_kw(ODOO_DB, uid, ODOO_PASS, "res.partner", "create", [{
        "name": nom, "email": email or False,
        "phone": info.get("telephone") or False,
        "customer_rank": 1,
    }])

def creer_devis(analyse, mail):
    try:
        uid, models = connexion_odoo()
    except Exception as e:
        print(f"[ODOO] Erreur : {e}")
        return None
    partner_id = trouver_ou_creer_client(uid, models, analyse)
    validite = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")
    note = f"Mail de {mail['de']} — {mail['sujet']}\n{analyse.get('notes_commerciales','')}"
    ids_prod = models.execute_kw(ODOO_DB, uid, ODOO_PASS, "product.product", "search",
                                  [[["type","=","service"],["sale_ok","=",True]]], {"limit":1})
    produit_id = ids_prod[0] if ids_prod else None
    lignes = []
    for art in analyse.get("articles", []):
        ligne = {"name": art.get("description","Article à définir"),
                 "product_uom_qty": art.get("quantite") or 1,
                 "price_unit": art.get("prix_unitaire") or 0.0}
        if produit_id:
            ligne["product_id"] = produit_id
        lignes.append((0, 0, ligne))
    if not lignes:
        lignes = [(0, 0, {"name": "Demande à compléter", "product_uom_qty": 1, "price_unit": 0.0})]
    devis_id = models.execute_kw(ODOO_DB, uid, ODOO_PASS, "sale.order", "create", [{
        "partner_id": partner_id, "validity_date": validite,
        "note": note, "order_line": lignes, "state": "draft",
    }])
    return devis_id

def main(dry_run=False):
    print(f"\n{'='*50}")
    print(f"IneArt — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"Mode : {'TEST' if dry_run else 'NORMAL'}")
    print(f"{'='*50}\n")
    mails = lire_mails_imap()
    if not mails:
        print("Aucun nouveau mail.")
        return
    nb_devis = 0
    for mail in mails:
        print(f"\n📧 {mail['de']}\n   {mail['sujet']}")
        analyse = analyser_mail(mail)
        type_m = analyse.get("type", "AUTRE")
        conf = analyse.get("confiance", 0)
        print(f"   → {type_m} ({conf:.0%})")
        if type_m in ("DEMANDE_DEVIS", "COMMANDE_VALIDEE") and conf >= 0.6:
            if dry_run:
                print("   [TEST] " + json.dumps(analyse, ensure_ascii=False))
            else:
                devis_id = creer_devis(analyse, mail)
                if devis_id:
                    print(f"   ✅ Devis #{devis_id} créé dans Odoo")
                    nb_devis += 1
        else:
            print("   ⏭ Ignoré")
    print(f"\n✅ {len(mails)} mail(s) — {nb_devis} devis créé(s)\n")

if __name__ == "__main__":
    main(dry_run="--dry-run" in sys.argv)
