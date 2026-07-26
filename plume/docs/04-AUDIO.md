# PLUME — Chaîne audio

## Voix de narration

| Contexte | Solution | Coût |
|---|---|---|
| Prototype (défaut) | **Voix OpenAI via Pollinations** (`text.pollinations.ai`, modèle `openai-audio`, voix `onyx` grave / `fable` conteur) — [doc API](https://github.com/pollinations/pollinations/blob/main/APIDOCS.md) | Gratuit (1 req/15 s) |
| Prototype (repli hors-ligne) | **Piper** voix `fr_FR-upmc` (Pierre) embarquée en mp3 | Gratuit |
| Production (défaut) | **Piper auto-hébergé** (`narrate-episode`, `TTS_SERVER_URL`) | ~5 €/mois fixes |
| Production (option premium) | **ElevenLabs** (`ELEVENLABS_API_KEY`) | payant |

Le rythme du conteur : le moteur narratif écrit chaque scène en fragments avec
durées de pause ; l'audio est assemblé phrase par phrase (validé en v9,
pauses ×0,6 après retour client).

## Musique & bruitages

**Prototype** : tout est synthétisé sur mesure (`prototype/build_soundscape.py`)
— berceuse de boîte à musique, craquement de branche à micro-fractures,
bruit « velours » pour le feuillage, merle, gorgées… Zéro dépendance,
zéro licence, fonctionne hors-ligne. Bruitages calés sur les scènes
(le découpage narratif garantit la synchro : « cric, crac » ouvre sa scène).

**Production — banques de sons libres à intégrer** (inaccessibles depuis
l'environnement de dev, à télécharger depuis le poste de build) :

| Banque | Licence | Usage |
|---|---|---|
| [Pixabay Sound Effects](https://pixabay.com/sound-effects/) | Pixabay License (libre, sans attribution) | bruitages nature (brindilles, feuilles, oiseaux, eau) |
| [Freesound](https://freesound.org/search/?q=&f=license:%22Creative+Commons+0%22) (filtre CC0) | CC0 | bruitages fins, boucles d'ambiance |
| [OpenGameArt — CC0 Sounds Library](https://opengameart.org/content/cc0-sounds-library) | CC0 | pas, portes, matières |
| [BBC Sound Effects](https://sound-effects.bbcrewind.co.uk/) | RemArc (usage personnel/édu — vérifier avant commercial) | références qualité |

Pipeline production : le moteur narratif tague chaque scène
(`feuilles`, `oiseau`, `pas`, `eau`, `carillon`…) → `narrate-episode`
mixe voix + tag SFX + musique (pistes séparées, ducking -12 dB sous la voix).

## Mixage (validé en test v13)

- musique : volume élément **0,22**, boucle sans couture, démarre dans le
  geste utilisateur (jamais en autoplay différé — leçon du bug v11) ;
- bruitages : **0,20**, un par scène, déclenchés au début de scène ;
- voix : plein niveau ; pause globale = voix + musique + bruitages,
  et la pause gagne toujours contre les minuteurs d'enchaînement.
