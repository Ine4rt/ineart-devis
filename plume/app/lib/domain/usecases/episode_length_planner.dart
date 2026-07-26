import '../entities/child_profile.dart';

/// Le Pacte de sommeil (innovation #34) : le parent choisit la durée du
/// récit (1 à 10 minutes) ; à défaut, une durée par âge s'applique. L'app
/// raccourcit encore si l'heure du coucher approche — la confiance des
/// parents est le moteur de la rétention.
class EpisodeLengthPlanner {
  const EpisodeLengthPlanner();

  /// Bornes produit : jamais moins de 1 min, jamais plus de 10 min.
  static const Duration minimum = Duration(minutes: 1);
  static const Duration maximum = Duration(minutes: 10);

  /// Durée par défaut par tranche d'âge, si le parent n'a rien choisi.
  static const Map<AgeBand, Duration> _defaultByAge = {
    AgeBand.tiny: Duration(minutes: 3),
    AgeBand.explorer: Duration(minutes: 5),
    AgeBand.hero: Duration(minutes: 7),
    AgeBand.legend: Duration(minutes: 9),
  };

  /// Marge entre la fin de l'histoire et l'heure cible d'endormissement
  /// (bisou, extinction).
  static const Duration _windDown = Duration(minutes: 3);

  Duration plan({
    required AgeBand ageBand,
    required DateTime now,
    required DateTime targetSleepTime,
    int? parentMinutes,
  }) {
    final ideal = _clamp(
      parentMinutes != null
          ? Duration(minutes: parentMinutes)
          : _defaultByAge[ageBand]!,
    );
    final available = targetSleepTime.difference(now) - _windDown;

    if (available >= ideal) return ideal;
    if (available <= minimum) return minimum; // jamais d'histoire au rabais
    return available;
  }

  Duration _clamp(Duration d) =>
      d < minimum ? minimum : (d > maximum ? maximum : d);

  /// Message honnête au parent quand l'histoire a été raccourcie.
  String? parentNote(Duration planned, Duration requested) {
    if (planned >= requested) return null;
    final m = planned.inMinutes;
    return 'Épisode raccourci à $m min ce soir pour protéger l\'heure du coucher. '
        'Le récit reprendra exactement ici demain.';
  }
}
