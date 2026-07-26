# PLUME — Chaîne audio

## Le conteur IA (prototype v17)

Le prototype ne récite plus un texte figé : à chaque soirée, **l'IA écrit le
chapitre**. La génération est lancée au début des trois respirations — l'attente
devient le rituel — et le repli sur le feuilleton écrit à la main garantit que
l'enfant n'attend jamais.

Ce que reçoit le conteur : prénom, âge, compagnon, rappel du chapitre
précédent, réussite et inquiétude du jour, durée demandée (1-10 min →
~115 mots/min), et la palette de bruitages. Ce qu'il rend un JSON
`{titre, scenes:[{texte, son}], rappel_suivant}`.

Consignes d'écriture (identiques prototype ↔ production) : ton album jeunesse
(phrases courtes, onomatopées, répétitions, images concrètes, tutoiement),
monde réel uniquement, journée de l'enfant tissée sans jamais la nommer, fin
qui conduit au sommeil.

| Contexte | Modèle | Coût |
|---|---|---|
| Prototype | Pollinations (`text.pollinations.ai`, OpenAI-compatible, sans clé) | Gratuit |
| Production | Claude via `generate-episode` | à l'usage |

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

**Tags de bruitage** (choisis par l'IA, un par scène) : `feuilles`,
`craquement`, `oiseau`, `pas`, `etoiles`, `maison`, `joie`, `reconfort`,
`eau`, `nuit`, `silence`. Validés côté serveur (`SFX_TAGS`) et mappés côté
client sur les clips du mixeur.

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

## Mixage (mesuré en test v16)

⚠️ **Piège majeur : Safari iOS ignore `HTMLMediaElement.volume`.** Baisser
`audio.volume` n'a aucun effet sur iPhone — le son sort à plein niveau. C'est
la cause des « c'est trop fort » répétés en v13-v15 malgré des chiffres
divisés par cinq. Deux conséquences pour l'app Flutter comme pour le web :

1. **Le mixage doit être gravé dans les fichiers audio** (seul niveau
   garanti partout) : musique à crête **0,05**, bruitages à **0,10**, pour
   une narration à crête ~0,9 — soit −25 dB et −19 dB sous la voix.
2. **Le réglage dynamique passe par un vrai mixeur** : nœuds de gain
   Web Audio (côté Flutter : volume de piste `just_audio`), ouverts dans un
   geste utilisateur, jamais par `element.volume`.

- **Curseur « ambiance sonore »** dans l'espace parent (0-100 %, défaut 50 %),
  persisté : le parent règle lui-même musique + bruitages. À 0 % → voix seule.
- **Ducking** : la musique descend à **45 %** de son niveau quand la voix
  parle (rampe 0,7 s) et remonte en 1,2 s dans les silences.
- Entrée en fondu de 1,8 s ; boucle sans couture ; démarre dans le geste
  utilisateur (jamais en autoplay différé — leçon du bug v11).
- Pause globale = voix + musique + bruitages, et la pause gagne toujours
  contre les minuteurs d'enchaînement.

**Vérification** : le test e2e branche un `AnalyserNode` sur le mixeur et
**mesure** la crête réelle (0,0038 pendant la voix, 0 curseur à zéro). Un test
qui se contenterait de lire `element.volume` ne prouverait rien.
