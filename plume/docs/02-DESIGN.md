# PLUME — Design System & Expérience

> Objectif : un Apple Design Award. Chaque écran doit être une scène, pas une page.

---

## 1. Direction artistique — « L'Heure Bleue »

L'app vit au moment où le jour bascule dans la nuit. Tout le design découle de cette heure-là.

### Palette

| Rôle | Nom | Hex | Usage |
|---|---|---|---|
| Fond nuit | **Encre d'Heure Bleue** | `#0E1230` | Fond principal (dark = mode natif) |
| Fond nuit profond | **Abysse** | `#080B20` | Scènes d'histoire, ciel étoilé |
| Lueur primaire | **Or de Lanterne** | `#FFC66E` | Actions, magie, focus |
| Secondaire | **Lavande de Crépuscule** | `#9D8CFF` | Compagnon, éléments interactifs doux |
| Accent | **Pêche d'Aube** | `#FFA98F` | Émotions, cœurs, check-in |
| Succès | **Vert Luciole** | `#7FE3A8` | Graines de lumière, croissance |
| Texte | **Blanc d'Étoile** | `#F4F1E8` | Jamais du blanc pur : blanc chaud |
| Texte doux | **Brume** | `#9BA0C0` | Sous-titres, méta |
| Mode jour | **Lin d'Aube** | `#FAF6EC` | Fond clair (app parent, matin) |

Règle : **le mode sombre est le mode principal** (c'est une app du soir). Le mode clair
existe pour l'espace parent et les usages diurnes (rêves du matin, album).

### Typographie

- **Titres** : `Fraunces` (serif chaleureux, optique « soft ») — voix du conte.
- **Texte & UI** : `Nunito Sans` — rond, lisible, amical.
- **Texte enfant / dyslexie** : option `OpenDyslexic`, interlettrage augmenté.
- Échelle : Display 40/48 · H1 32 · H2 24 · Corps 17 (jamais moins de 16 pour le parent, 20 pour l'enfant).

### Matière & lumière

- Aucun aplat mort : chaque fond respire (dégradé animé très lent + particules-lucioles au parallaxe gyroscopique).
- Les cartes sont des **vitraux** : verre dépoli sombre, bord lumineux 1 px Or de Lanterne à 20 %.
- Les ombres sont des **halos** (lumière émise), pas des ombres portées : dans la nuit, ce qui compte émet de la lumière.
- Grain de papier subtil (2 %) sur les pages de l'album : matière de livre.

### Iconographie

Icônes dessinées à la main, trait 2 px arrondi, remplies d'un dégradé de lueur quand actives.
Set custom : lanterne (accueil), plume (histoires), patte d'étoile (compagnon), cœur-météo
(check-in), livre-monde (album), lune-parent (espace parent).

---

## 2. Les mascottes

- **Plume** — la mascotte de l'app (pas le compagnon de l'enfant) : une petite
  renarde-chouette porteuse de lanterne, gardienne de l'Heure Bleue. Elle guide
  l'onboarding, annonce les événements mondiaux, console quand il n'y a pas d'histoire.
  Silhouette simple (déclinable en icône d'app) : oreilles rondes, queue en plume qui luit.
- **Le Compagnon** — unique par enfant, généré par « ADN » : espèce hybride (12 bases ×
  12 traits × motifs × palette), tempérament (curieux/timide/farceur/protecteur…),
  démarche, cri. Rendu par fiche de référence visuelle pour cohérence des illustrations.
- **Les Archivistes** — hiboux copistes qui « écrivent » l'album : justification
  diégétique de toute l'UI de l'album et des lettres.

---

## 3. Navigation & architecture d'information

Deux mondes, deux grammaires :

- **Côté enfant (plein écran, paysage du monde)** : pas de tab bar. On navigue en
  *se déplaçant dans le monde* (la carte est le menu). Gestes larges, cibles ≥ 64 px.
- **Côté parent (verrouillé par geste adulte)** : navigation classique et dense,
  onglets : Aujourd'hui · Enfants · Album · Réglages.

```
Lancement
 └─ Porte de l'Heure Bleue (splash vivant)
     ├─ Onboarding (7 scènes) — 1re fois
     └─ La Carte du Monde (hub enfant)
         ├─ Le Foyer du Compagnon
         ├─ La Lanterne (rituel du soir → histoire)
         ├─ La Grande Bibliothèque (album, coloriages, capsules)
         └─ La Lune (espace parent, verrou)
              ├─ Aujourd'hui (check-in, pacte de sommeil)
              ├─ Enfants (profils, famille, moments de vie)
              ├─ Album & Livre de l'Année
              └─ Réglages (voix, abonnement, confidentialité)
```

---

## 4. Les écrans, un par un

### 4.1 La Porte de l'Heure Bleue (splash)
Un ciel qui passe du jour au crépuscule en 1,8 s (la couleur du ciel réel dépend de
l'heure locale). Plume traverse l'écran avec sa lanterne, les lucioles forment le logo.
*Micro-interaction :* secouer le téléphone fait scintiller les étoiles.

### 4.2 Onboarding — « La naissance d'un monde » (7 scènes)
Pas un formulaire : une cérémonie. Chaque réponse transforme visiblement le monde en fond.
1. **Le prénom** — l'enfant (ou le parent) le prononce : les lettres s'écrivent en constellation.
2. **L'âge** — un arbre pousse du nombre d'anneaux correspondant.
3. **Les passions** (cartes à attraper : dinosaures, danse, espace, foot…) — chaque carte attrapée fait apparaître une île au loin.
4. **Les peurs** (posées au parent, formulation douce) — des lanternes s'allument « pour éclairer les endroits sombres ».
5. **La famille & les animaux** — photos optionnelles ; chaque proche devient une étoile nommée.
6. **L'œuf** — l'enfant caresse l'écran pour réchauffer un œuf ; il éclot : **naissance du compagnon unique** (moment signature de l'app, partageable).
7. **Le monde se révèle** — zoom arrière : tout ce qui a été répondu est devenu géographie. Titre : « Le monde de Léo. Épisode 1 ce soir. »

### 4.3 La Carte du Monde (hub enfant)
Une carte picturale vivante (parallaxe 3 couches, cycle jour/nuit réel, météo du monde).
Les lieux débloqués sont éclairés ; les terres inexplorées sont sous brume dorée.
Le compagnon se promène librement sur la carte (idle animations : il bâille, chasse une luciole).
*Micro-interactions :* toucher un lieu → il s'illumine et murmure son nom ; pincer → vue
royaume entier ; les graines de lumière plantées poussent visiblement ici.
*Transition signature :* toute navigation est un **vol de lanterne** (caméra qui glisse), jamais un push latéral.

### 4.4 Le Foyer du Compagnon
Grotte/nid personnalisé. Le compagnon réagit au toucher (12 zones), répète les mots
appris, montre ses souvenirs préférés (mini-flashbacks des épisodes). Jauge discrète de
son humeur — jamais de barres de « faim » anxiogènes : il va toujours bien, il a juste des envies.

### 4.5 Le Rituel du Soir (la Lanterne)
Chorégraphie en 3 temps, sautables individuellement :
1. **Brossage** — le compagnon se brosse les dents en miroir, chanson de 2 min générée avec le prénom.
2. **Calme** — 3 respirations guidées : la lanterne gonfle/dégonfle, l'écran se réchauffe (blue light ↓).
3. **Le seuil** — « On y va ? » — le parent pose le téléphone face cachée si mode voix seule.

### 4.6 Le Théâtre de l'Histoire (player)
L'écran le plus important. Plein écran, une illustration de scène en Ken Burns très lent,
texte optionnel en bas (karaoké doux mot à mot pour les lecteurs débutants).
- **Choix** : 2-3 médaillons illustrés qui flottent ; l'enfant touche, ou *le dit à voix haute*.
- **Pause vivante** : si l'enfant parle, la narration se suspend et le narrateur répond brièvement.
- **Endormissement** : silence détecté → voix -15 % vitesse, musique s'estompe, écran s'éteint
  progressivement, marque-page automatique « endormi à… » pour reprendre demain.
- *Micro-interactions :* toucher une luciole de l'illustration la fait s'envoler ; incliner le
  téléphone déplace légèrement la scène.

### 4.7 La Grande Bibliothèque (album)
Un hall de bibliothèque vu de face, rayonnages = mois. Chaque épisode est un livre relié
dont la tranche porte la date et une pastille d'émotion. Ouvrir un livre = les pages
tournent avec les illustrations, le texte, les choix faits (et « ce qui se serait passé si… » verrouillé — teasing).
Coin **Atelier** : coloriages des illustrations (palette magique, remplissage qui scintille), impression AirPrint.
Coin **Capsules** : sabliers scellés avec date d'ouverture.

### 4.8 L'espace parent (la Lune)
Design inverse : clair (Lin d'Aube), dense, efficace, une main.
- **Aujourd'hui** : check-in émotionnel (4 cartes tactiles, 20 s max, saisie vocale possible),
  heure de coucher cible → durée d'épisode calculée, rêve du matin à écouter à deux.
- **Enfants** : profil vivant (ce que l'IA a appris, *transparent et éditable* — confiance),
  casting familial, moments de vie à venir, peurs en cours de travail avec progression.
- **Album & Livre** : progression du Livre de l'Année (« 214 pages écrites »), commande d'impression.
- **Réglages** : voix (dont voix des proches), abonnement, données & vie privée (export/effacement en un geste).
*Verrou parental :* « maintenez les deux lunes appuyées et tournez » — infranchissable à 4 ans, 1 s pour un adulte.

### 4.9 Écrans systèmes enchantés
Même les états vides sont des scènes : pas de réseau → « Les nuages cachent le monde…
mais la Bibliothèque est toujours ouverte » (mode offline : épisodes téléchargés).
Chargement de génération → Plume écrit à la plume sur un parchemin (jamais de spinner).
Erreur → un Archiviste confus s'excuse avec une plume cassée.

---

## 5. Mouvement — règles d'animation

1. **Rien n'apparaît : tout naît.** Fade + scale 0.96→1 + halo, courbe `easeOutCubic`, 240-400 ms.
2. **La lumière précède l'objet** : le halo arrive 80 ms avant l'élément.
3. **60 fps non négociable** : particules en shader (CustomPainter/Impeller), jamais de widgets par luciole.
4. Transitions de navigation = **vol de lanterne** (Hero + courbe de Bézier spatiale), 450 ms.
5. Haptique : légère à chaque éclosion/choix (`lightImpact`), moyenne à l'éclosion du compagnon.
6. Le soir (après 19 h), toutes les durées d'animation ×1,25 et amplitudes ×0,8 : l'app se calme avec l'enfant.
7. Respect strict de `reduce motion` (parallaxe et particules coupées, fondus conservés).

---

## 6. Accessibilité

- Cibles enfant ≥ 64 px, parent ≥ 44 px.
- Contrastes AA sur tous les textes (le Blanc d'Étoile sur Encre = 13:1).
- VoiceOver/TalkBack : chaque scène a une description narrative (c'est une app d'histoires — l'accessibilité est diégétique).
- Sous-titres complets de la narration, option police dyslexie, option « sans musique ».

---

## 7. Le parcours d'une soirée parfaite (résumé)

18 h 40 — Notification douce au parent : « Le rêve de Pipo parlait d'une grotte qui chante… (épisode prêt à 20 h) ».
19 h 55 — Check-in 20 s pendant que l'enfant met son pyjama (le rituel brossage tourne déjà).
20 h 05 — La Lanterne. Trois respirations. « On y va ? »
20 h 08 — Épisode 214. Le dragon sauvé en janvier revient. L'enfant choisit à voix haute.
20 h 19 — Silence détecté. La voix ralentit, l'écran s'éteint. Marque-page : 20 h 21.
20 h 22 — Sur le téléphone du parent : « Léo s'est endormi à 20 h 21. Le monde continue. À demain. »
