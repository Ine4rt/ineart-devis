# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

Deux publics, strictement séparés.

**Public de la page de présentation** : dirigeants de très petites entreprises
belges qui n'ont pas encore de site web — artisans, commerçants, professions
libérales (boulangerie, garage, menuiserie, fleuriste, cabinet). Peu ou pas à
l'aise avec le numérique. Beaucoup n'ont qu'une page Facebook, parfois rien.
Ils évaluent la page depuis un téléphone, souvent en fin de journée, entre deux
tâches. Ils ne connaissent pas le vocabulaire technique.

**Public de la console** : l'équipe d'IneWeb — aujourd'hui une personne. Outil
interne, ouvert plusieurs heures par jour pour suivre clients, échéances et
encaissements.

## Product Purpose

Concevoir, héberger et entretenir des sites web sur mesure pour de très petites
entreprises, contre un paiement à la création puis un abonnement annuel qui
couvre nom de domaine, hébergement et maintenance. Le client ne gère rien.

La page de présentation a un seul objectif : déclencher un e-mail.

## Positioning

Développement entièrement sur mesure — aucun modèle, aucun constructeur de
site — assorti d'une prise en charge complète de l'infrastructure. Le client
n'a aucun compte à créer, aucun outil à apprendre, aucun mot de passe à retenir.
Un seul interlocuteur, qui connaît son dossier.

## Operating Context

**Le parcours commercial passe exclusivement par l'e-mail.** Le prospect écrit,
IneWeb propose, la validation se fait ensemble par échanges écrits. Il n'y
a pas de prise de rendez-vous, pas de formulaire d'appel, pas de créneau à
réserver. Cette contrainte est confirmée par le client et lie toute la page :
aucune formulation ne doit suggérer un rendez-vous, un appel ou une rencontre.

Côté production : IneWeb achète les noms de domaine, gère l'hébergement et
assure la maintenance. La console interne suit ces engagements.

## Capabilities and Constraints

- Sites sur mesure : vitrine, e-commerce, landing, application web, refonte.
- Prise en charge : domaine, hébergement, SSL, sauvegardes, mises à jour,
  modifications de contenu en cours d'année.
- **Les tarifs ne sont pas affichés publiquement.** Décision confirmée : le prix
  se propose au cas par cas par e-mail, puis se valide avec le client. Aucun
  montant ne doit figurer sur la page de présentation.
- Zone d'intervention non formalisée à ce stade ; le travail se fait à distance.

## Brand Commitments

- Nom : **IneWeb** (remplace « Ine4rt », y compris dans la console interne).
- **IneWeb se présente comme une société, pas comme un indépendant.** La voix
  publique est « nous » ; jamais « je », jamais « je crée des sites ». Cette
  contrainte est confirmée par le client et lie toute rédaction future.
- Palette imposée : **gris très foncé et orange**. L'orange s'emploie **par
  touches** — le client a explicitement rejeté les aplats pleine largeur, jugés
  trop lourds.
- Registre visuel demandé : **épuré, élégant, professionnel**. Graisses légères,
  larges respirations, filets fins. Un lettrage gras et étendu a été refusé.
- Mouvement attendu : **glissements**, dont une bannière défilante.
- E-mail de contact confirmé : **info@ineart.be** — le domaine reste `ineart.be`
  alors que la marque devient IneWeb ; à trancher.
- Aucun numéro de téléphone : le contact est uniquement par e-mail.

## Evidence on Hand

- Console interne fonctionnelle (44 routes) : `src/`.
- Jeu de données de démonstration : `prisma/seed.ts` — **fictif**, il ne
  constitue pas une preuve commerciale.
- **Aucun témoignage, logo client, chiffre d'affaires, statistique de marché ni
  réalisation publique n'existe à ce jour.** Rien de tel ne doit être inventé sur
  la page. Une section « réalisations » ne pourra exister qu'une fois de vrais
  sites en ligne.
- Les maquettes de la bannière (boulangerie, garage, cabinet) sont des **mises
  en page de démonstration**, annoncées comme telles et servies par des adresses
  en `-exemple.be`. Elles ne doivent jamais être présentées comme des clients.
- **Aucune photographie disponible.** Les emplacements d'image des maquettes
  portent la mention « photo » en clair, en attendant de vraies prises de vue.

## Product Principles

1. **L'e-mail est la seule porte.** Toute incitation à l'action mène à un
   message écrit, jamais à un appel ou à un rendez-vous.
2. **Aucun jargon.** « Nom de domaine » s'explique ; « DNS », « CMS » et
   « responsive » ne s'écrivent pas.
3. **Le sur-mesure est l'argument, et il parle à des gens qui en font.** Les
   prospects sont eux-mêmes des artisans : le parallèle porte mieux qu'un
   discours technique.
4. **Rien d'inventé.** Pas de faux avis, pas de statistiques non sourcées, pas
   de prix non validés.
5. **Le client ne gère rien.** C'est la promesse centrale, et elle doit rester
   vérifiable dans les faits.

## Accessibility & Inclusion

Public peu à l'aise avec le numérique, lecture fréquente sur téléphone, souvent
en conditions de faible attention. Contrastes élevés, corps de texte généreux,
zones tactiles larges, et respect de `prefers-reduced-motion` — le mouvement ne
doit jamais être une condition de compréhension.
