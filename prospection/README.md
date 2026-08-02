# IneArt — Prospection Huy

Chaîne complète : trouver les entreprises de Huy sans site internet, leur
générer un site de démonstration, le mettre en ligne 8 jours sur le FTP, et
préparer le mail ou l'appel de prospection.

## Le pipeline

```
collecte.py    OSM (+ Google Maps)  ->  data/prospects.json
generateur.py  prospects.json       ->  sites/<slug>/index.html + data/manifest.json
ftp_deploy.py  sites/               ->  FTP /demos/<slug>/
mailer.py      manifest.json        ->  mails/<slug>.eml + mails/envois.csv
nettoyage.py   manifest.json        ->  supprime du FTP les maquettes expirées
```

## Démarrage rapide (en local)

```bash
pip install -r prospection/requirements.txt
cd prospection

python generateur.py --exemple   # 20 sites de démonstration, sans réseau
python mailer.py                 # mails + fiches d'appel + envois.csv
open tableau_de_bord.html        # récapitulatif de tout ce qui a été produit
```

Pour la collecte réelle :

```bash
python collecte.py --limite 20              # OSM seul, rapide
python collecte.py --limite 20 --enrich     # + Google Maps (Playwright)
python generateur.py
python ftp_deploy.py --dry-run              # vérifier avant d'envoyer
python ftp_deploy.py
python mailer.py
```

## Depuis GitHub Actions

- **`Prospection Huy`** (manuel) : collecte → génération → FTP → mails.
  Paramètres : communes, nombre max, enrichissement Google, déploiement.
  Les mails et le tableau de bord sont téléchargeables en artefact.
- **`Nettoyage des maquettes expirees`** (quotidien, 04h17 UTC) : supprime du
  FTP tout ce qui a dépassé sa date. C'est ce workflow qui rend la limite des
  8 jours réelle — le compte à rebours seul est cosmétique.

Retrait immédiat d'une société qui le demande : lancer le workflow de
nettoyage avec le champ `slug` (ou `python nettoyage.py --slug <slug>`).

## Secrets GitHub à configurer

| Secret | Rôle | Exemple |
|---|---|---|
| `FTP_HOST` | serveur FTP | `ftp.ineart.be` |
| `FTP_USER` | identifiant FTP | `ineart` |
| `FTP_PASS` | mot de passe FTP | — |
| `FTP_DIR` | dossier distant des démos | `/demos` |
| `DEMO_BASE_URL` | URL publique correspondante | `https://www.ineart.be/demos` |
| `INEART_EMAIL` | adresse de contact affichée | `info@ineart.be` |
| `INEART_TEL` | téléphone affiché | `+32 471 ...` |
| `INEART_SITE` | site de l'agence | `https://www.ineart.be` |

`FTP_DIR` et `DEMO_BASE_URL` doivent désigner le même dossier, l'un côté FTP,
l'autre côté web.

## Réglages

Tout se règle dans `config.py` : tarifs (250 € le site, 50 €/an le domaine),
durée de validité (8 jours), communes, coordonnées de l'agence. Les palettes
et les textes par métier sont dans `secteurs.py`.

## Segmentation des prospects

| Segment | Signification | Intérêt |
|---|---|---|
| A | aucune présence web | prospect vierge |
| B | uniquement Facebook / Instagram | **meilleur segment** |
| C | site déclaré mais mort ou injoignable | bon segment |
| D | site fonctionnel | exclu automatiquement |

## Points d'attention

- **Adresses mail** : aucune source publique ne les fournit de façon fiable.
  Comptez 25 à 40 % de couverture. Les autres prospects reçoivent une fiche
  d'appel (`mails/<slug>-appel.txt`) avec accroche, tarifs et objections.
- **Envoi** : jamais en masse. 20 à 30 mails par jour maximum, sinon le
  domaine part en spam. Aucun envoi automatique n'est implémenté, c'est
  volontaire.
- **Mention légale** : chaque maquette porte en pied de page la mention
  « site non officiel, non affilié », la date de suppression automatique et
  l'adresse de retrait sur demande.
- **Prospection e-mail en Belgique** : opt-out toléré vers les sociétés,
  opt-in requis vers les indépendants personnes physiques. La colonne
  `canal_conseille` d'`envois.csv` aide à trier.
