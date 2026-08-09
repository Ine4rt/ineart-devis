/**
 * VOID RUNNER — constantes de simulation.
 *
 * Toutes les valeurs sont en pixels et en secondes, dans l'espace "monde"
 * virtuel (hauteur de référence 360 px). Le rendu se contente de mettre à
 * l'échelle : la physique ne dépend JAMAIS de la résolution de l'écran.
 *
 * Règle absolue du projet : la simulation est déterministe. Pas de Math.random
 * dans la physique, pas de delta-time variable. Deux parties avec les mêmes
 * entrées produisent exactement le même résultat — c'est ce qui garantit
 * qu'une mort est toujours la faute du joueur, et c'est ce qui permet au
 * solveur automatique (tests/) de valider chaque niveau.
 */

// --- Pas de temps -----------------------------------------------------------
export const TILE = 24; // unité de grille utilisée par les niveaux
export const FIXED_DT = 1 / 60; // pas de simulation fixe
export const MAX_FRAME_DT = 0.25; // anti "spirale de la mort" après un lag

// --- Dimensions du monde de référence ---------------------------------------
export const VIEW_H = 360; // hauteur virtuelle constante
export const VIEW_W_MIN = 560; // largeur mini (tablette 4:3)
export const VIEW_W_MAX = 900; // largeur maxi (téléphone très allongé)

// --- Robot ------------------------------------------------------------------
export const PLAYER_W = 20;
export const PLAYER_H = 22;

// Course
export const RUN_SPEED = 190;
export const ACCEL_GROUND = 2400;
export const ACCEL_AIR = 1500;
export const FRICTION_GROUND = 2800;
export const FRICTION_AIR = 260;

// Saut : hauteur d'apex = JUMP_VEL² / (2 * GRAVITY) ≈ 87 px ≈ 3.6 tuiles
// Portée horizontale à pleine vitesse ≈ 118 px ≈ 4.9 tuiles
export const GRAVITY = 1800;
export const JUMP_VEL = 560;
export const JUMP_CUT = 0.42; // saut variable : relâcher coupe l'élan
export const MAX_FALL = 900;

// Confort de contrôle (indispensable pour que le jeu soit "juste")
export const COYOTE_TIME = 0.09; // saut toléré après avoir quitté le sol
export const JUMP_BUFFER = 0.12; // saut mémorisé avant de toucher le sol

// Dash (débloqué au chapitre 3)
export const DASH_SPEED = 430;
export const DASH_TIME = 0.16;
export const DASH_ENDLAG = 0.06;

// Divers
export const CORNER_CORRECT = 4; // "coin arrondi" : évite d'accrocher un angle
export const SNAP_EPS = 0.01;

// --- Identifiants d'état de mort (pilotent l'animation + le son) -------------
export const DEATH = {
  ZAP: 'zap', // électrocution
  CRUSH: 'crush', // écrasement
  BURN: 'burn', // laser
  FALL: 'fall', // chute dans le vide
  BOOM: 'boom', // générique
};
