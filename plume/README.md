# 🪶 Plume — le monde qui grandit avec votre enfant

> « Chaque soir, le monde de votre enfant continue. »

Plume n'est pas un générateur d'histoires : c'est un **monde persistant et unique**
qui naît avec l'enfant et évolue pendant des années. Les personnages se souviennent
de tout, les choix ont des conséquences des mois plus tard, le compagnon est
génétiquement unique, et les émotions du jour sont tissées en filigrane dans
l'épisode du soir.

## Contenu du dossier

| Dossier | Contenu |
|---|---|
| [`docs/01-CONCEPT.md`](docs/01-CONCEPT.md) | Vision, **35 innovations** classées (effet waouh, difficulté, valeur, viral, commercial), roadmap en 4 actes, modèle économique, sécurité |
| [`docs/02-DESIGN.md`](docs/02-DESIGN.md) | Design system « L'Heure Bleue » : palette, typographies, mascottes, **tous les écrans**, règles d'animation, accessibilité, parcours d'une soirée |
| [`docs/03-ARCHITECTURE.md`](docs/03-ARCHITECTURE.md) | Architecture technique : Flutter + Supabase + Claude, moteur de canon narratif, offline-first, coûts IA |
| [`app/`](app/) | Application **Flutter** (Clean Architecture, Riverpod, go_router, dark mode natif, tests) |
| [`supabase/`](supabase/) | Schéma Postgres (RLS par famille) + Edge Functions : `generate-episode` (le Conteur), `morning-dream` (le rêve du matin), `record-choice` (les graines narratives) |

## Lancer l'application

Prérequis : Flutter ≥ 3.22.

```bash
cd app
flutter pub get
flutter run          # démarre en mode DÉMO (aucun backend requis)
flutter test         # tests du domaine (ADN compagnon, graines, pacte de sommeil)
```

Sans configuration, l'app tourne sur `DemoWorldRepository` : le monde de Léo,
son compagnon Pipo, et l'épisode 214 « Le retour de Cendreflamme » — parfait
pour la revue design et les démos.

### Brancher Supabase (production)

```bash
supabase db push                       # applique migrations/0001_init.sql
supabase functions deploy generate-episode morning-dream record-choice
supabase secrets set ANTHROPIC_API_KEY=sk-ant-…
```

Puis dans `app/lib/main.dart` : décommenter `Supabase.initialize(...)` et
surcharger `worldRepositoryProvider` avec `SupabaseWorldRepository`.

## Ce qui est implémenté dans ce dépôt

- **Domaine complet et testé** : ADN unique du compagnon (genèse déterministe),
  stades de croissance pluriannuels, graines narratives à longue portée
  (sélection « une promesse n'est jamais oubliée »), Pacte de sommeil
  (calibrage de la durée d'épisode).
- **8 écrans** : onboarding-cérémonie (avec éclosion de l'œuf), Carte du Monde,
  Foyer du Compagnon, Rituel de respiration, Théâtre de l'Histoire (choix
  illustrés, progression en fil d'or), Grande Bibliothèque, Météo émotionnelle,
  espace parent (verrou « deux lunes »).
- **Design system** : ciel étoilé + lucioles en un seul CustomPainter (60 fps),
  cartes-vitraux, bouton-lanterne qui respire, transition « vol de lanterne »,
  respect de reduce-motion, haptique.
- **Backend** : schéma canonique complet avec RLS, moteur de génération
  d'épisodes contraint par le canon (Claude), rêve du matin, plantation des choix.

## Prochaines étapes (voir roadmap docs/01-CONCEPT.md §3)

TTS multi-voix + endormissement adaptatif · pipeline d'illustrations avec fiches
de référence visuelles · notifications push · cache offline Drift · Livre de
l'Année (export PDF) · golden tests des écrans.
