# Standard téléphonique IA

Un client appelle, une IA décroche, qualifie sa demande, et la fiche de devis
arrive par mail. Pensé pour les artisans qui ratent leurs appels sur chantier.

```
Le client appelle  →  Twilio  →  serveur  →  Claude  →  mail avec la fiche
```

Le numéro de l'appelant est capté automatiquement : l'IA ne le demande jamais.

## Deux niveaux, deux coûts

| | Coût | Ce que ça valide |
|---|---|---|
| **`demo/`** — démonstrateur navigateur | **0 €** | Le dialogue et la fiche, en face à face avec un prospect |
| **`serveur/`** — vrai standard | ~20 € une fois | Un vrai appel téléphonique |

---

## 1. Le démonstrateur — à utiliser en rendez-vous

Ouvrez `demo/index.html` dans **Chrome sur Android** ou **Safari sur iPhone**
(la reconnaissance vocale n'existe pas sur tous les navigateurs). Dépliez
**Réglages**, collez votre clé API Anthropic et votre adresse mail, puis
appuyez sur **Décrocher** et parlez.

La clé reste dans le stockage local du téléphone et n'est envoyée qu'à l'API
Anthropic. Aucun serveur, aucun compte, aucune installation.

**En rendez-vous** : tendez votre téléphone à l'artisan et laissez-le parler à
l'IA. Il comprend en trente secondes, sans rien avoir à imaginer.

Coût réel d'un dialogue complet en Haiku : environ **0,002 €**.

---

## 2. Le vrai standard

### Ce qu'il faut

- Un compte Twilio (l'essai offre 15,50 $ de crédit et 75 minutes de voix)
- Une clé API Anthropic
- Un compte mail SMTP — les mêmes identifiants que le projet devis

### Installation

```bash
cd standard-ia/serveur
pip install -r requirements.txt

export ANTHROPIC_API_KEY=sk-ant-...
export MAIL_USER=info@ineart.be
export MAIL_PASS=...
export SOCIETE="IneArt"
export METIER="artisan"

python standard.py
```

Exposez le serveur avec une URL publique gratuite :

```bash
cloudflared tunnel --url http://localhost:5000
```

Puis dans Twilio : **Phone Numbers → votre numéro → A call comes in →
Webhook POST** `https://<votre-tunnel>/appel`.
Réglez aussi **Call status changes** sur `https://<votre-tunnel>/statut` pour
recevoir une fiche partielle si le client raccroche en cours de route.

### Tester sans payer

Le compte d'essai Twilio permet un vrai appel de bout en bout :

1. Vérifiez votre mobile dans Twilio (SMS, trente secondes).
2. Achetez un numéro **américain ou britannique** avec le crédit d'essai —
   un numéro belge exige un justificatif d'adresse et plusieurs jours de
   validation réglementaire.
3. Appelez-le depuis votre mobile vérifié.

Limites de l'essai : seuls les numéros vérifiés peuvent appeler, un court
message d'essai précède l'IA, et le crédit expire après 30 jours.

### Passer en réel

Créditez le compte (~20 €), demandez un numéro belge avec justificatif
d'adresse, et hébergez le serveur — **Oracle Cloud Always Free** convient et
reste gratuit à vie.

Astuce : plutôt qu'un nouveau numéro, activez le **renvoi conditionnel** de
votre numéro actuel. Sur Proximus, Orange et Base, `**61*<numéro IA>#` renvoie
les appels sans réponse après cinq sonneries. Vous décrochez quand vous
pouvez, l'IA prend le relais sinon. C'est aussi l'argument de vente le plus
fort auprès des artisans : ils ne changent rien.

### Coût par appel

| Poste | Coût |
|---|---|
| Réception de l'appel (2 min) | ~0,03 € |
| Reconnaissance vocale et voix | incluses dans Twilio |
| Claude Haiku | ~0,002 € |
| Mail | 0 € |
| **Total** | **~0,05 à 0,15 €** |

Soit 2 à 12 € par mois pour un artisan recevant 40 à 80 appels — à comparer à
un abonnement facturé 49 €.

---

## Réglages

Tout se pilote par variables d'environnement : `SOCIETE`, `METIER`,
`VOIX_TWILIO` (`Polly.Lea-Neural` par défaut, essayez `Polly.Remi-Neural` pour
une voix masculine), `MAIL_DEST`, `SMTP_HOST`, `SMTP_PORT`.

Le déroulé des questions se modifie dans la constante `CONSIGNE` de
`standard.py` — et dans `CONSIGNE` de `demo/index.html` pour le démonstrateur.

## Points de vigilance

- **RGPD** : l'annonce d'accueil précise que l'appelant parle à un assistant
  vocal. Ne retirez pas cette phrase.
- **Conversations en mémoire** : l'état des appels vit dans le processus. Un
  seul serveur suffit pour des dizaines de clients ; au-delà, passez sur Redis.
- **Silence du client** : l'IA relance une fois, puis conclut proprement après
  douze échanges pour éviter les appels sans fin.
