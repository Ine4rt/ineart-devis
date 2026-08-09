import { Solid, Crumble, Vanish, Ghost, Mirage, Mover, Door, Crusher } from './platforms.js';
import { Zap, Laser } from './hazards.js';
import {
  Exit, Button, Zone, Timer, Logic, Checkpoint, Teleport, Gravity, MemGate,
  Clone, Sign, Deco,
} from './logic.js';
import { Fan, Field } from './fields.js';

/**
 * Registre des mécaniques.
 *
 * Ajouter un piège au jeu = ajouter une classe et une ligne ici. Les niveaux
 * n'ont jamais besoin d'être recodés : ce sont des données.
 */
export const REGISTRY = {
  solid: Solid,
  crumble: Crumble,
  vanish: Vanish,
  ghost: Ghost,
  mirage: Mirage,
  mover: Mover,
  door: Door,
  crusher: Crusher,
  zap: Zap,
  laser: Laser,
  exit: Exit,
  button: Button,
  zone: Zone,
  timer: Timer,
  logic: Logic,
  check: Checkpoint,
  tele: Teleport,
  grav: Gravity,
  memgate: MemGate,
  clone: Clone,
  sign: Sign,
  deco: Deco,
  fan: Fan,
  field: Field,
};

export const ENTITY_TYPES = Object.keys(REGISTRY);

export function createEntity(spec, world) {
  const C = REGISTRY[spec.t];
  if (!C) {
    // Un niveau mal écrit ne doit pas planter le jeu en production ; il doit
    // en revanche faire échouer les tests de validation.
    if (typeof console !== 'undefined') console.warn('Entité inconnue :', spec.t);
    return null;
  }
  return new C(spec, world);
}
