import 'package:flutter/foundation.dart';

/// Le compagnon : unique par enfant, défini par un ADN (innovation #14).
/// L'ADN est généré une seule fois (voir CompanionGenesis) puis le compagnon
/// évolue au fil des épisodes, des émotions et des choix.
@immutable
class Companion {
  const Companion({
    required this.id,
    required this.name,
    required this.dna,
    this.bondLevel = 0,
    this.learnedWords = const [],
    this.mood = CompanionMood.curious,
    this.stage = GrowthStage.hatchling,
  });

  final String id;
  final String name;
  final CompanionDna dna;

  /// Lien avec l'enfant : ne descend jamais (pas de jauges anxiogènes).
  final int bondLevel;

  /// Mots de l'enfant que le compagnon a « appris » et replace dans les épisodes.
  final List<String> learnedWords;
  final CompanionMood mood;
  final GrowthStage stage;

  Companion copyWith({
    int? bondLevel,
    List<String>? learnedWords,
    CompanionMood? mood,
    GrowthStage? stage,
  }) =>
      Companion(
        id: id,
        name: name,
        dna: dna,
        bondLevel: bondLevel ?? this.bondLevel,
        learnedWords: learnedWords ?? this.learnedWords,
        mood: mood ?? this.mood,
        stage: stage ?? this.stage,
      );

  /// Le compagnon grandit avec le lien — jamais de régression.
  Companion afterEpisode({List<String> newWords = const []}) {
    final bond = bondLevel + 1;
    return copyWith(
      bondLevel: bond,
      learnedWords: {...learnedWords, ...newWords}.toList(),
      stage: GrowthStage.forBond(bond),
    );
  }
}

/// ADN : 12 bases × 12 traits × motifs × palette × tempérament.
/// Deux enfants n'auront jamais le même compagnon.
@immutable
class CompanionDna {
  const CompanionDna({
    required this.species,
    required this.trait,
    required this.pattern,
    required this.paletteSeed,
    required this.temperament,
  });

  final String species; // base hybride, ex: « renardeau-lune »
  final String trait; // particularité visuelle, ex: « queue-lanterne »
  final String pattern; // motif de pelage
  final int paletteSeed; // graine de la palette de couleurs
  final Temperament temperament;

  static const speciesBases = [
    'renardeau-lune', 'louveteau-nuage', 'chaton-comète', 'oisillon-braise',
    'faon-brume', 'loutron-étoile', 'hérisson-lueur', 'axolotl-rosée',
    'chouette-velours', 'dragonnet-mousse', 'lapin-aurore', 'panda-lichen',
  ];
  static const traits = [
    'queue-lanterne', 'oreilles-feuilles', 'ailes-de-papier', 'cornes-de-givre',
    'moustaches-étincelles', 'crinière-nébuleuse', 'pattes-nuages', 'yeux-vairons-astraux',
    'écharpe-de-brume', 'taches-constellation', 'plumes-carillon', 'sourcils-broussaille',
  ];
  static const patterns = [
    'rayures-marée', 'pois-lucioles', 'flammèches', 'écailles-pastel',
    'marbrures-encre', 'flocons', 'vagues', 'feuillage',
  ];
}

enum Temperament { curious, shy, playful, protective, dreamy, brave }

enum CompanionMood { curious, sleepy, joyful, thoughtful, mischievous }

/// Étapes de croissance — sur des années, pas des jours.
enum GrowthStage {
  hatchling, // éclosion
  sprout, // ~2 semaines de rituels
  wanderer, // ~2 mois
  guardian, // ~1 an
  ancient; // ~3 ans : le compagnon devient légende du monde

  static GrowthStage forBond(int bond) => switch (bond) {
        < 14 => GrowthStage.hatchling,
        < 60 => GrowthStage.sprout,
        < 365 => GrowthStage.wanderer,
        < 1095 => GrowthStage.guardian,
        _ => GrowthStage.ancient,
      };
}
