import '../entities/child_profile.dart';

/// Le Pacte de sommeil (innovation #34) : Plume protège l'heure du coucher.
/// L'app calcule la durée idéale de l'épisode pour que l'enfant s'endorme
/// à l'heure cible — quitte à raccourcir l'histoire. La confiance des
/// parents est le moteur de la rétention.
class EpisodeLengthPlanner {
  const EpisodeLengthPlanner();

  /// Durée cible par tranche d'âge (narration calme).
  static const Map<AgeBand, Duration> _idealByAge = {
    AgeBand.tiny: Duration(minutes: 6),
    AgeBand.explorer: Duration(minutes: 9),
    AgeBand.hero: Duration(minutes: 12),
    AgeBand.legend: Duration(minutes: 15),
  };

  static const Duration _minimum = Duration(minutes: 4);

  /// Marge entre la fin de l'histoire et l'heure cible d'endormissement
  /// (respirations, bisou, extinction).
  static const Duration _windDown = Duration(minutes: 3);

  Duration plan({
    required AgeBand ageBand,
    required DateTime now,
    required DateTime targetSleepTime,
  }) {
    final ideal = _idealByAge[ageBand]!;
    final available = targetSleepTime.difference(now) - _windDown;

    if (available >= ideal) return ideal;
    if (available <= _minimum) return _minimum; // jamais d'histoire au rabais
    return available;
  }

  /// Message honnête au parent quand l'histoire a été raccourcie.
  String? parentNote(Duration planned, AgeBand ageBand) {
    final ideal = _idealByAge[ageBand]!;
    if (planned >= ideal) return null;
    final m = planned.inMinutes;
    return 'Épisode raccourci à $m min ce soir pour protéger l\'heure du coucher. '
        'Le monde reprendra exactement ici demain.';
  }
}
