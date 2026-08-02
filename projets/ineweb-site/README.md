# Site IneWeb — mise en ligne

Un seul fichier : `index.html`. Tout est dedans — la mise en forme, la police
Archivo, les maquettes de sites, les animations. Aucune image externe, aucune
feuille de style à côté, aucun script tiers. Poids : **63 Ko**.

## Par FTP

1. Connectez-vous à l'hébergement avec FileZilla, Cyberduck ou le gestionnaire
   de fichiers du panneau d'administration.
2. Ouvrez le dossier public — il s'appelle `www/`, `public_html/` ou
   `htdocs/` selon l'hébergeur.
3. Déposez `index.html` à la racine de ce dossier.

C'est tout. Le site est en ligne à l'adresse du domaine.

Aucune configuration serveur n'est nécessaire : ni PHP, ni base de données, ni
certificat particulier au-delà du HTTPS que l'hébergeur fournit déjà.

## Ne pas modifier ce fichier directement

`index.html` est **généré**. La source se trouve dans `presentation/` à la
racine du dépôt :

```
presentation/
├── page.body.html     ← la source, c'est ici qu'on écrit
├── build.mjs          ← la construction
└── fonts/             ← Archivo découpée, en base64
```

Pour modifier le site :

```bash
node presentation/build.mjs
```

La commande réécrit `projets/ineweb-site/index.html` et
`presentation/page.embed.html`. Une correction faite directement dans le
fichier généré sera écrasée à la prochaine construction.

## Ce que la page respecte

Ces contraintes ont été posées en cours de route ; elles sont dans le contenu,
pas dans le code, donc rien ne les protège automatiquement. À vérifier après
chaque modification :

- **Aucun prix affiché.** Le tarif se discute, il ne s'affiche pas.
- **Aucune mention de rendez-vous, d'appel ou de téléphone.** Tout passe par
  e-mail : `info@ineart.be`.
- **La voix est celle d'une société** — « nous », jamais « je ».
- **Les bandes alternent** blanc, gris clair et un unique bloc noir. Une page
  d'un seul ton fait un pavé.
- **L'orange sert à marquer l'action**, pas à remplir des surfaces.

## Sous-domaine ou domaine dédié ?

Le site se suffit d'un domaine à lui (`ineweb.be`). S'il doit cohabiter avec un
site existant, un sous-dossier fonctionne aussi : déposez `index.html` dans
`www/ineweb/` et l'adresse devient `ledomaine.be/ineweb/`. Le fichier ne
contient aucun chemin absolu, il fonctionne à n'importe quelle profondeur.

Note : l'adresse de contact est `info@ineart.be` alors que la marque s'appelle
IneWeb. C'est cohérent si Ineart reste la société et IneWeb l'activité web —
mais un prospect peut s'en étonner. Une redirection `info@ineweb.be` vers la
même boîte lèverait l'ambiguïté sans rien changer à votre organisation.
