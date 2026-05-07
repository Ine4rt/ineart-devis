"""
IneArt — Agent Mail v3
- Lit tous les mails recents (lus + non lus)
- Detecte devis / commandes / reponses deja envoyees
- Filtre les spams et expediteurs systeme
- Exporte un CSV compatible Odoo
- Sauvegarde mails.json pour le dashboard
"""

import imaplib
import email
from email.header import decode_header
import json
import os
import re
import sys
import csv
from datetime import datetime, timedelta
import anthropic

IMAP_HOST    = "imap.one.com"
IMAP_PORT    = 993
MAIL_USER    = os.environ["MAIL_USER"]
MAIL_PASS    = os.environ["MAIL_PASS"]
CLAUDE_KEY   = os.environ["ANTHROPIC_API_KEY"]
CLAUDE_MODEL = "claude-haiku-4-5-20251001"
JOURS_RETOUR = 5
OUTPUT_FILE  = "mails.json"
CSV_FILE     = "devis_odoo.csv"
MAX_MAILS    = 150

EXPEDITEURS_IGNORES = [
    "github.com", "anthropic.com", "paypal.be", "paypal.com",
    "bpost.be", "postnl.nl", "inpost-group.com", "zalando.be",
    "facebook.com", "business-updates.facebook.com", "relevanceai.com",
    "bigmat.be", "irobot.com", "link.com", "email.claude.com",
    "alun.dk", "odoo.com", "noreply", "no-reply",
    "notifications@", "notify@", "service-mail", "edm.",
]

def est_ignore(expediteur):
    exp = expediteur.lower()
    return any(x in exp for x in EXPEDITEURS_IGNORES)

def decoder_header(valeur):
    parties = decode_header(valeur or "")
    result = []
    for data, charset in parties:
        if isinstance(data, bytes):
            result.append(data.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(str(data))
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
    inbox = []
    sent_ids = set()

    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
        imap.login(MAIL_USER, MAIL_PASS)

        # INBOX
        imap.select("INBOX")
        depuis = (datetime.now() - timedelta(days=JOURS_RETOUR)).strftime("%d-%b-%Y")
        _, ids = imap.search(None, f'(SINCE "{depuis}")')
        ids_liste = ids[0].split()
        print(f"[IMAP] {len(ids_liste)} mail(s) dans INBOX.")

        for uid in ids_liste:
            _, data = imap.fetch(uid, "(BODY.PEEK[])")
            msg = email.message_from_bytes(data[0][1])
            expediteur = decoder_header(msg.get("From", ""))
            if est_ignore(expediteur):
                continue
            inbox.append({
                "uid":         uid.decode(),
                "de":          expediteur,
                "sujet":       decoder_header(msg.get("Subject", "")),
                "date":        msg.get("Date", ""),
                "corps":       extraire_texte(msg),
                "msg_id":      msg.get("Message-ID", ""),
                "in_reply_to": msg.get("In-Reply-To", ""),
            })

        # Dossier ENVOYE pour detecter les reponses
        for folder in ["Sent", "Sent Messages", "INBOX.Sent", "Gesendete"]:
            try:
                res, _ = imap.select(f'"{folder}"', readonly=True)
                if res != "OK":
                    continue
                _, sids = imap.search(None, f'(SINCE "{depuis}")')
                for sid in sids[0].split():
                    _, sdata = imap.fetch(sid, "(BODY.PEEK[HEADER.FIELDS (IN-REPLY-TO MESSAGE-ID SUBJECT)])")
                    smsg = email.message_from_bytes(sdata[0][1])
                    ref = smsg.get("In-Reply-To", "")
                    if ref:
                        sent_ids.add(ref.strip())
                    subj = decoder_header(smsg.get("Subject", ""))
                    if subj:
                        sent_ids.add("subj:" + subj.lower().replace("re: ", "").strip())
                print(f"[IMAP] {len(sent_ids)} reponse(s) detectee(s) dans {folder}.")
                break
            except Exception:
                continue

    return inbox, sent_ids


PROMPT_SYSTEME = """Tu es un assistant commercial pour IneArt, entreprise belge de broderie et personnalisation textile (serigraphie, broderie, DTF, flocage).
Analyse l'email et reponds UNIQUEMENT en JSON valide, sans markdown ni texte autour.

REGLES :
- DEMANDE_DEVIS : client demande prix/devis/offre, parle de quantites, broderie, impression, textiles. Inclus RE: Devis, relances, demandes de precisions.
- COMMANDE_VALIDEE : client confirme une commande, valide un devis, demande la production.
- AUTRE : newsletters, notifications, factures fournisseurs, livraisons. En cas de doute, choisis DEMANDE_DEVIS.

Format exact :
{
  "type": "DEMANDE_DEVIS",
  "confiance": 0.9,
  "client": {"nom": "...", "email": "...", "telephone": null, "entreprise": null},
  "articles": [{"description": "...", "quantite": 1, "technique": "broderie|serigraphie|dtf|flocage|autre", "notes": null}],
  "delai_souhaite": null,
  "budget_mentionne": null,
  "notes_commerciales": "resume en 1-2 phrases",
  "statut_suggere": "nouveau|relance|en_attente|urgent"
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
    for m in historique:
        for n in nouveaux:
            if m["uid"] == n["uid"] and n.get("repondu") and not m.get("repondu"):
                m["repondu"] = True
    tous = (a_ajouter + historique)[:MAX_MAILS]
    data = {
        "mis_a_jour": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total": len(tous),
        "mails": tous
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[JSON] {OUTPUT_FILE} : {len(tous)} mails ({len(a_ajouter)} nouveaux)")


def exporter_csv_odoo(mails_structures):
    lignes = []
    for m in mails_structures:
        if m.get("type") not in ("D", "C"):
            continue
        cl = m.get("cl", {})
        arts = m.get("arts", [])
        if not arts:
            arts = [{"d": m.get("sujet", ""), "q": 1, "tech": "", "n": ""}]
        for art in arts:
            lignes.append({
                "Order Date":         m.get("date", ""),
                "Customer Name":      cl.get("nom", "") or cl.get("email", ""),
                "Customer Email":     cl.get("email", ""),
                "Customer Company":   cl.get("ent", ""),
                "Customer Phone":     cl.get("tel", ""),
                "Product Description": art.get("d", ""),
                "Quantity":           art.get("q", 1),
                "Technique":          art.get("tech", ""),
                "Unit Price":         "",
                "Notes":              art.get("n", ""),
                "Deadline":           m.get("delai", ""),
                "Budget":             m.get("budget", ""),
                "Status":             "Devis" if m.get("type") == "D" else "Commande",
                "Replied":            "Oui" if m.get("repondu") else "Non",
                "Mail Subject":       m.get("sujet", ""),
                "AI Summary":         m.get("note", ""),
            })
    if not lignes:
        print("[CSV] Aucun devis/commande a exporter.")
        return
    with open(CSV_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=lignes[0].keys())
        writer.writeheader()
        writer.writerows(lignes)
    print(f"[CSV] {CSV_FILE} : {len(lignes)} ligne(s) exportee(s).")


def main(dry_run=False):
    print("\n" + "="*50)
    print(f"IneArt Agent Mail v3 - {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("Mode : TEST" if dry_run else "Mode : NORMAL")
    print("="*50 + "\n")

    mails_bruts, sent_ids = lire_mails_imap()

    if not mails_bruts:
        print("Aucun mail a traiter.")
        if not dry_run:
            sauvegarder_json([])
        return

    mails_structures = []
    nb_devis = nb_commandes = nb_repondus = 0

    for mail in mails_bruts:
        print(f"\n📧 {mail['de']}\n   {mail['sujet']}")

        # Detecter si deja repondu (par Message-ID ou par sujet)
        msg_id = mail.get("msg_id", "").strip()
        sujet_key = "subj:" + mail.get("sujet", "").lower().replace("re: ", "").strip()
        repondu = msg_id in sent_ids or sujet_key in sent_ids
        if repondu:
            nb_repondus += 1
            print(f"   ↩ Deja repondu")

        analyse = analyser_mail(mail)
        type_m  = analyse.get("type", "AUTRE")
        conf    = analyse.get("confiance", 0)

        if type_m == "DEMANDE_DEVIS":
            type_court = "D"
            nb_devis += 1
        elif type_m == "COMMANDE_VALIDEE":
            type_court = "C"
            nb_commandes += 1
        else:
            type_court = "A"

        print(f"   -> {type_m} ({round(conf*100)}%){' - REPONDU' if repondu else ''}")

        cl = analyse.get("client", {})
        mails_structures.append({
            "uid":     mail["uid"],
            "de":      mail["de"],
            "sujet":   mail["sujet"],
            "date":    datetime.now().strftime("%d/%m/%Y %H:%M"),
            "type":    type_court,
            "cf":      round(conf * 100),
            "repondu": repondu,
            "statut":  analyse.get("statut_suggere", "nouveau"),
            "cl": {
                "nom":   cl.get("nom", ""),
                "email": cl.get("email", ""),
                "ent":   cl.get("entreprise", ""),
                "tel":   cl.get("telephone", ""),
            },
            "arts": [
                {
                    "d":    a.get("description", ""),
                    "q":    a.get("quantite", 1),
                    "tech": a.get("technique", ""),
                    "n":    a.get("notes", ""),
                }
                for a in analyse.get("articles", [])
            ],
            "delai":  analyse.get("delai_souhaite", ""),
            "budget": analyse.get("budget_mentionne", ""),
            "note":   analyse.get("notes_commerciales", ""),
        })

    if not dry_run:
        sauvegarder_json(mails_structures)
        exporter_csv_odoo(mails_structures)

    print("\n" + "="*50)
    print(f"Devis     : {nb_devis}")
    print(f"Commandes : {nb_commandes}")
    print(f"Repondus  : {nb_repondus}")
    print(f"Total     : {len(mails_bruts)} mail(s)")
    print("="*50 + "\n")


if __name__ == "__main__":
    main(dry_run="--dry-run" in sys.argv)
