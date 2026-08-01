# Page de présentation

Page commerciale destinée au démarchage : un lien à envoyer avant ou après un
premier contact avec une entreprise qui n'a pas encore de site.

Elle est **volontairement séparée de la console**. La console est privée ; cette
page est publique. Aucun code n'est partagé entre les deux, donc rien de ce que
vous modifiez ici ne peut casser votre outil de gestion.

## Utilisation

`index.html` est un document autonome de ~110 Ko : aucune requête réseau, aucune
dépendance, aucun script. Déposez-le sur n'importe quel hébergement statique
(votre hébergeur actuel, Netlify, Vercel, un simple dossier FTP) et c'est en
ligne.

## Modifier le contenu

Le fichier source est **`page.body.html`**. Après modification :

```bash
node presentation/build.mjs
```

La construction régénère `index.html` (document complet) et `page.embed.html`
(le même contenu sans enveloppe, pour un aperçu partagé). Les polices sont
réinjectées en base64 à chaque construction.

## À personnaliser avant diffusion

| Où | Quoi |
|---|---|
| Section `#contact` | Téléphone, e-mail, zone d'intervention |
| Section `#tarifs` | Les deux montants (1 250 € et 480 €/an sont des ordres de grandeur cohérents avec le marché, à confirmer) |
| Après vos premiers sites | Ajouter une section « Réalisations » avec de vrais projets |

**Aucun témoignage, chiffre ou logo client n'a été inventé.** Les arguments
reposent sur des situations concrètes et des engagements que vous prenez —
notamment celui, dans la FAQ, de transmettre le site et le domaine si le client
part un jour. Vérifiez que vous êtes d'accord avec cet engagement avant de
diffuser la page.

## Choix de conception

Le propos tient en une idée : vos prospects sont des artisans et des
commerçants, et **ils travaillent eux-mêmes sur mesure**. La page se construit
sur ce parallèle plutôt que sur un argumentaire technique — d'où le vert des
auvents de devanture, la plaque émaillée du bandeau d'accueil, et un vocabulaire
sans jargon (« nom de domaine » expliqué, jamais « DNS » ou « CMS »).

Les polices (Fraunces pour les titres, Karla pour le texte) sont intégrées en
base64 dans le fichier. C'est ce qui garantit que la page s'affiche à
l'identique partout, y compris derrière un réseau qui bloquerait un CDN.
`fonts/sub.py` régénère les sous-ensembles à partir des fichiers d'origine
(licence SIL Open Font, redistribution et intégration autorisées).
