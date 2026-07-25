import 'package:flutter/foundation.dart';

/// Une « graine narrative » (innovation #4) : toute décision, peur ou moment de
/// vie est planté avec une date de germination. Le moteur la récolte des
/// semaines ou des mois plus tard — le dragon sauvé en janvier revient en juin.
@immutable
class NarrativeSeed {
  const NarrativeSeed({
    required this.id,
    required this.kind,
    required this.summary,
    required this.plantedAt,
    required this.germinateAfter,
    this.priority = SeedPriority.normal,
    this.harvestedAt,
  });

  final String id;
  final SeedKind kind;

  /// Résumé canonique, ex: « Léo a épargné le dragon Cendreflamme au pont. »
  final String summary;
  final DateTime plantedAt;

  /// Date à partir de laquelle la graine peut germer dans un épisode.
  final DateTime germinateAfter;
  final SeedPriority priority;
  final DateTime? harvestedAt;

  bool get isHarvested => harvestedAt != null;
  bool isRipe(DateTime now) => !isHarvested && !now.isBefore(germinateAfter);

  NarrativeSeed harvested(DateTime when) => NarrativeSeed(
        id: id,
        kind: kind,
        summary: summary,
        plantedAt: plantedAt,
        germinateAfter: germinateAfter,
        priority: priority,
        harvestedAt: when,
      );
}

enum SeedKind {
  choiceConsequence, // conséquence d'un choix de l'enfant
  fearWork, // peur à travailler en filigrane (#7, #8)
  lifeMoment, // rentrée, déménagement… à préparer en avance (#9)
  kindnessSeed, // bonne action réelle → graine de lumière (#11)
  worldEvent, // événement mondial (#6)
}

enum SeedPriority { low, normal, high }

/// Sélectionne les graines à récolter pour l'épisode du soir :
/// les moments de vie imminents d'abord, puis par priorité, puis les plus
/// anciennes (une promesse narrative ne doit jamais être oubliée).
List<NarrativeSeed> pickSeedsForTonight(
  List<NarrativeSeed> seeds,
  DateTime now, {
  int max = 2,
}) {
  final ripe = seeds.where((s) => s.isRipe(now)).toList()
    ..sort((a, b) {
      final lifeMoment = _isLifeMoment(b).compareTo(_isLifeMoment(a));
      if (lifeMoment != 0) return lifeMoment;
      final prio = b.priority.index.compareTo(a.priority.index);
      if (prio != 0) return prio;
      return a.plantedAt.compareTo(b.plantedAt);
    });
  return ripe.take(max).toList();
}

int _isLifeMoment(NarrativeSeed s) => s.kind == SeedKind.lifeMoment ? 1 : 0;
