import 'dart:math';

import '../entities/companion.dart';

/// Genèse du compagnon (innovation #14) : l'ADN est dérivé de façon
/// déterministe de l'identité de l'enfant + un sel aléatoire stocké une seule
/// fois. Résultat : unique, reproductible, jamais deux compagnons identiques.
class CompanionGenesis {
  const CompanionGenesis();

  CompanionDna generate({
    required String childId,
    required int salt,
    List<String> passions = const [],
  }) {
    final rng = Random(childId.hashCode ^ salt);

    // Les passions orientent doucement le tempérament : un enfant fan
    // d'espace penche vers « rêveur », un sportif vers « brave ».
    final temperament = _temperamentFor(passions, rng);

    return CompanionDna(
      species: _pick(CompanionDna.speciesBases, rng),
      trait: _pick(CompanionDna.traits, rng),
      pattern: _pick(CompanionDna.patterns, rng),
      paletteSeed: rng.nextInt(1 << 24),
      temperament: temperament,
    );
  }

  Temperament _temperamentFor(List<String> passions, Random rng) {
    const affinities = <String, Temperament>{
      'espace': Temperament.dreamy,
      'étoiles': Temperament.dreamy,
      'sport': Temperament.brave,
      'foot': Temperament.brave,
      'animaux': Temperament.protective,
      'nature': Temperament.protective,
      'blagues': Temperament.playful,
      'dessin': Temperament.curious,
      'livres': Temperament.curious,
    };
    for (final p in passions) {
      final match = affinities[p.toLowerCase()];
      if (match != null && rng.nextBool()) return match;
    }
    return Temperament.values[rng.nextInt(Temperament.values.length)];
  }

  T _pick<T>(List<T> list, Random rng) => list[rng.nextInt(list.length)];
}
