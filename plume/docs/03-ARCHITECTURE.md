# PLUME — Architecture technique

Stack : **Flutter** (iOS/Android/iPad) · **Supabase** (Postgres, Auth, Storage, Realtime,
Edge Functions Deno) · **Claude API** (moteur narratif) · TTS expressif · Riverpod · Clean Architecture.

---

## 1. Vue d'ensemble

```
┌────────────────────────── Flutter app ──────────────────────────┐
│ presentation/  écrans + widgets + contrôleurs Riverpod           │
│ domain/        entités immuables + use-cases + interfaces repo   │
│ data/          repositories impl + DTO + sources (supabase/local)│
│ core/          thème, routing, DI, utils, offline queue          │
└───────────────┬─────────────────────────────────────────────────┘
                │ supabase-dart (auth, pg, storage, realtime)
┌───────────────▼───────────── Supabase ──────────────────────────┐
│ Postgres + RLS   world graph, épisodes, profils, check-ins       │
│ Storage          illustrations, audio, exports PDF               │
│ Edge Functions   generate-episode · evolve-world · companion-brain│
│                  morning-dream · yearbook-export                 │
│ pg_cron          évolution nocturne du monde, rêves du matin     │
└───────────────┬─────────────────────────────────────────────────┘
                │ (les clés IA ne quittent JAMAIS le serveur)
        ┌───────▼────────┐   ┌──────────────┐   ┌────────────────┐
        │ Claude API      │   │ TTS expressif │   │ Génération img │
        │ (canon + récit) │   │ multi-voix    │   │ (style monde)  │
        └────────────────┘   └──────────────┘   └────────────────┘
```

Principes :
- **Aucun appel IA depuis l'app** : tout passe par les Edge Functions (clés protégées, coûts contrôlés, garde-fous serveurs).
- **Le monde évolue côté serveur** (pg_cron à 3 h du matin) : l'app ne fait que *découvrir* les changements.
- **Offline-first pour le soir** : l'épisode du soir + audio + images sont pré-générés et téléchargés en avance (rien de pire qu'un enfant qui attend).

## 2. Le moteur de canon narratif (cœur du système)

Le canon est un **graphe versionné** dans Postgres :

- `world_entities` (personnage, lieu, objet, légende) avec `traits jsonb`, fiche visuelle, état courant.
- `world_events` : journal append-only de tout ce qui s'est passé (source de vérité).
- `narrative_seeds` : les « graines » — promesses narratives horodatées (choix de l'enfant,
  peurs à travailler, moments de vie à préparer) avec `germinate_after` et `priority`.
- `story_arcs` : le Fil d'Or — arcs planifiés (saison, acte, épisodes cibles).

Pipeline de génération d'un épisode (Edge Function `generate-episode`) :
1. **Contexte** : requête SQL qui assemble le sous-graphe pertinent (entités actives, 
   dernières `world_events`, graines mûres, check-in du jour, âge/passions/peurs du profil).
2. **Plan** : Claude produit un plan d'épisode JSON (scènes, choix, graines plantées/récoltées) 
   contraint par le canon fourni — *jamais* de génération libre.
3. **Récit** : Claude écrit le texte scène par scène, ton calibré à l'âge, avec le tissage 
   émotionnel du jour en filigrane.
4. **Vérification canon** : passe de validation (contradictions, sécurité, vocabulaire par âge).
5. **Commit** : nouvelles entités/événements/graines écrits en base dans une transaction ; 
   audio TTS et 3 illustrations générés ; le tout poussé en Storage ; notification « épisode prêt ».

## 3. Schéma de données (extrait)

Voir `supabase/migrations/0001_init.sql` — familles, enfants, compagnons (ADN jsonb),
mondes, entités, événements, graines, épisodes, scènes, check-ins, membres de famille castés,
capsules temporelles. **RLS sur toutes les tables** : une famille ne voit que son monde.

## 4. Application Flutter

- **State** : Riverpod (`Notifier`/`AsyncNotifier`, codegen évitable pour lisibilité).
- **Navigation** : go_router, transitions custom « vol de lanterne ».
- **Rendu magique** : shaders/CustomPainter pour ciel, lucioles, halos (60 fps, Impeller).
- **Offline** : Drift (SQLite) comme cache canonique local ; file d'attente d'écritures 
  (check-ins, choix) rejouée à la reconnexion ; épisodes précachés (texte+audio+images).
- **Audio** : just_audio + session audio configurée pour le mode « écran posé ».
- **Notifications** : FCM/APNs via Supabase — rêve du matin, épisode prêt, événements mondiaux.
- **Thème** : dark = principal (soir), light = espace parent ; tokens dans `core/theme`.
- **Tests** : domaine pur testé unitairement (ADN compagnon, moteur de graines, calcul de durée
  d'épisode), golden tests des écrans clés, tests de widgets des flux critiques.

Arborescence (implémentée dans `app/`) :

```
lib/
  main.dart · app.dart
  core/{theme,router,widgets,utils}
  domain/{entities,usecases,repositories}
  data/{models,repositories,sources}
  features/
    onboarding/  · worldmap/ · companion/ · ritual/ · story/
    album/       · checkin/  · parent/
```

## 5. Sécurité, vie privée, coûts

- RGPD-K/COPPA : minimisation, chiffrement au repos, export & effacement en un geste,
  aucune pub, aucun tracker tiers.
- Garde-fous IA côté serveur : liste de thèmes autorisés par âge, filtre de sortie,
  bornage strict du narrateur interactif au canon, journal d'audit.
- Coûts IA maîtrisés : 1 épisode/jour/enfant pré-généré la nuit (heures creuses),
  prompt caching sur la story bible, images en batch, TTS streamé et mis en cache.
- Observabilité : traces par épisode (tokens, latence, coût), alerte dérive de coût.
