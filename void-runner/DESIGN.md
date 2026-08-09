# VOID RUNNER — note de conception

Ce document explique les décisions. Le code explique le reste.

---

## 1. Le pari

Un « troll platformer » repose sur un cycle très court :

> attente du joueur → trahison → mort → compréhension → nouvel essai → réussite

Ce cycle ne tient qu'à une condition : **le joueur doit toujours pouvoir
reconstituer sa faute après coup.** Une mort incompréhensible ne produit pas
« ahh, j'avais mal compris », elle produit « ce jeu est mal fait ». La
différence entre les deux tient à quelques millisecondes de pré-signal et à
une charte de couleurs tenue.

D'où la règle qui commande tout le projet :

> **Difficile, mais jamais arbitraire.**
> Un joueur qui meurt sans avoir vu d'orange ni de rouge est un joueur trahi —
> c'est le seul cas où le jeu est en tort.

---

## 2. Choix techniques, et pourquoi

**HTML5 Canvas 2D + modules ES, sans framework, empaqueté par Capacitor.**
Un moteur lourd (Unity, Godot) apporterait de la physique 3D, un éditeur, un
pipeline d'assets — rien de ce dont ce jeu a besoin. En échange il coûterait
30 à 60 Mo de binaire, plusieurs secondes de démarrage et l'impossibilité de
tester la logique en dehors de l'éditeur. Ici : ~150 Ko, démarrage instantané,
et un moteur qui tourne aussi bien en Node que dans Safari.

**Pas de moteur physique tiers.** Un jeu de plateforme de précision ne veut pas
d'un solveur de contraintes : il veut des collisions AABB à axes séparés et un
contrôle total sur le sol, l'air, le coyote time et le portage par les
plateformes. Un moteur généraliste rendrait ces réglages impossibles.

**Pas fixe déterministe (1/60 s) avec accumulateur.** Sur un écran 120 Hz, un
pas de temps variable donnerait une hauteur de saut différente de celle d'un
écran 60 Hz. Inacceptable dans un jeu qui reproche ses erreurs au joueur. Le
déterminisme a un second effet, décisif : il rend le jeu **testable**.

**Zéro asset.** Robot, décors, effets, musique et sons sont générés au trait ou
synthétisés à l'exécution. Poids nul, netteté à toutes les résolutions,
huit robots pour zéro octet, et aucun risque de réemployer par accident un
élément d'un jeu existant.

**Niveaux déclaratifs.** Un niveau est un tableau d'objets. Aucune ligne de
code n'est écrite pour une salle particulière — sinon la centième salle
coûterait cent fois plus cher que la première.

---

## 3. Le solveur : la décision dont je suis le plus content

Dans un jeu construit sur la mort, la faute impardonnable est un niveau
**infranchissable**. Elle est indétectable à l'œil : une demi-tuile de trop sur
un gouffre ne se voit pas sur un plan.

`tests/solver.js` joue donc réellement chaque niveau, avec le moteur de
production, par recherche en faisceau sur les actions possibles. `npm test`
renvoie, pour les trente salles, le temps de la solution trouvée.

Il a trouvé cinq défauts pendant le développement — dont trois qui seraient
passés en production :

| Trouvé par le solveur | Nature |
|---|---|
| Niveau 11 — dalles déphasées d'une demi-période | 0,26 s de recouvrement : injouable |
| Niveau 10 — presses hydrauliques | le faisceau s'effondrait ; c'était le solveur, pas le niveau |
| Dalles effondrées jamais réapparues | **bug moteur** : les entités inactives ne recevaient plus de mise à jour |
| Niveau 9 — téléporteur défectueux sur le chemin obligatoire | boucle infinie |
| Niveau 12 — sortie fuyante | s'arrêtait en plein vide si le joueur cessait de la suivre |

Et le navigateur en a trouvé un sixième : la trappe du niveau 1 descendait avec
le robot dessus, qui atterrissait tranquillement au fond au lieu de tomber.
Deux panneaux latéraux ont réglé l'affaire.

Le solveur ne prétend pas jouer *bien*. Il prouve qu'un chemin existe. C'est
exactement ce qu'on lui demande.

---

## 4. Game feel

Les quatre réglages ci-dessous ne sont pas du confort : sans eux, le joueur
accuse les contrôles au lieu d'accuser sa décision — et tout le propos du jeu
s'effondre.

| Réglage | Valeur | Rôle |
|---|---|---|
| Coyote time | 90 ms | sauter juste après avoir quitté le rebord |
| Jump buffer | 120 ms | sauter juste avant de toucher le sol |
| Saut à hauteur variable | ×0,42 au relâcher | le « petit saut » est une commande à part entière |
| Correction de coin | 4 px | on n'accroche jamais un angle de plateforme |

La hitbox mortelle est **3 px plus petite** que la hitbox physique de chaque
côté : on ne meurt jamais d'un frôlement.

Le reste du ressenti tient à trois retours immédiats — visuel (déformation du
robot, particules, secousse), sonore, haptique — et à un unique chiffre :
**320 ms entre la mort et la reprise**, écourtables en appuyant sur saut.

---

## 5. Progression

| Chapitre | Salles | Ce qu'il apprend | Ce qu'il retourne |
|---|---|---|---|
| 1 — Station abandonnée | 1-6 | courir, sauter court, lire un cycle, patienter | le décor ment |
| 2 — Laboratoire | 7-12 | téléporteurs, presses, dalles clignotantes | les **symboles** mentent : porte, panneau, voyant |
| 3 — Zone de gravité | 13-18 | inversion, souffleries, champs magnétiques | la **mémoire** ment (niveau 15) |
| 4 — Centrale | 19-24 | impulsion, lames, glace, séquences | la **maîtrise** ment : ce que tu sais faire devient le piège |
| 5 — Secteur interdit | 25-30 | écho, mémoire persistante, structures absentes | **tu** mens : ton propre trajet te tue |

Trois niveaux portent tout le propos :

- **Niveau 1** — le sol s'ouvre (mort 1) ; on saute, un laser s'allume au-dessus
  (mort 2) ; la bonne réponse est un saut *court*. En vingt secondes, le joueur
  a appris la règle du jeu et une commande fine.
- **Niveau 15** — géométrie identique au niveau 4, mécanismes changés. Le joueur
  exécute sa solution mémorisée et meurt. Le jeu n'a pas triché : la règle
  n'a pas bougé, c'est lui qui n'a pas regardé. Une seule fois dans tout le jeu.
- **Niveau 29** — le niveau 1, à la tuile près, entièrement retourné : la porte
  du fond est mortelle, et la vraie sortie est **dans le trou**. Il faut accepter
  de refaire exactement ce qui nous a tués à la première salle.

---

## 6. Le dosage de l'humour

Une vanne après chaque mort devient un temps de chargement. Les répliques
n'apparaissent donc que dans trois cas : palier rond d'essais, mort ironique
(à moins de 90 px de la sortie), ou tirage rare — environ **18 %** des morts
ordinaires. Le compteur d'essais, lui, est toujours affiché : c'est une
information, pas une plaisanterie.

Le reste de l'humour est porté par le robot. Il n'a qu'un écran pour visage et
sept expressions ; il est surpris pendant la montée d'un saut, inquiet pendant
la chute, et ses yeux deviennent des spirales quand il est électrocuté. Un gag
d'une demi-seconde transforme une punition en spectacle — et un spectacle,
on veut le revoir.

---

## 7. Ergonomie mobile

- Paysage, `viewport-fit=cover`, marges `safe-area` respectées partout.
- **Hauteur virtuelle fixe (360 px), largeur variable selon le ratio.** Un
  téléphone allongé voit plus large, jamais plus haut : personne ne doit
  pouvoir repérer un piège parce qu'il a un meilleur écran.
- Boutons tactiles volontairement **petits à l'œil, larges au toucher** : 56 px
  visibles, 90 px de zone de frappe. Mesuré sur un écran de 390 px de haut,
  au-delà de 68 px le bouton de saut recouvre la ligne de sol — donc la sortie,
  donc le robot.
- Main directrice inversable dans les paramètres.
- Perte de focus (appel, notification) → pause automatique. On ne meurt pas
  parce que le téléphone a sonné.
- Qualité adaptative : si la moyenne dépasse 23,5 ms par image, les effets de
  fond tombent en premier. Le gameplay n'est jamais dégradé.

---

## 8. Ce que je ferais ensuite

1. **Test auprès de joueurs sur les salles 7 à 12.** Le solveur prouve qu'elles
   sont franchissables ; il ne dit rien du nombre de morts avant le déclic. La
   cible est 1 à 5 — il faut la mesurer, pas la supposer.
2. **Rejouabilité fantôme** des records personnels (la trace est déjà
   enregistrée pour les clones : c'est presque gratuit).
3. **Éditeur de niveaux en jeu**, exportant directement le format déclaratif.
   C'est le chemin le plus court vers 100 salles.
4. **Localisation**, puis sauvegarde distante.
