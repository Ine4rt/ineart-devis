import 'package:flutter/foundation.dart';

/// Ce que Plume sait de l'enfant. Transparent et éditable par le parent
/// (docs/03-ARCHITECTURE.md §5 — la confiance est une fonctionnalité).
@immutable
class ChildProfile {
  const ChildProfile({
    required this.id,
    required this.firstName,
    required this.birthDate,
    this.passions = const [],
    this.fears = const [],
    this.familyCast = const [],
    this.pets = const [],
    this.readingLevel = ReadingLevel.listener,
    this.storyMinutes,
    this.targetSleepTime = const SleepTime(20, 30),
  });

  final String id;
  final String firstName;
  final DateTime birthDate;

  /// « dinosaures », « danse », « espace »… chaque passion est une île du monde.
  final List<String> passions;

  /// Peurs en cours de travail — tissées en filigrane, jamais frontalement.
  final List<String> fears;

  /// Proches castés en personnages (mamie → la Gardienne des Recettes Magiques).
  final List<FamilyMember> familyCast;
  final List<String> pets;
  final ReadingLevel readingLevel;

  /// Durée du récit voulue par le parent, 1 à 10 min (`children.story_minutes`).
  /// `null` = le parent n'a rien réglé : la durée par âge s'applique.
  final int? storyMinutes;

  /// Heure de coucher cible du Pacte de sommeil (`children.target_sleep_time`).
  final SleepTime targetSleepTime;

  int ageAt(DateTime date) {
    var age = date.year - birthDate.year;
    final hadBirthday = date.month > birthDate.month ||
        (date.month == birthDate.month && date.day >= birthDate.day);
    if (!hadBirthday) age--;
    return age;
  }

  /// Bande d'âge narrative : calibre vocabulaire, thèmes et durée d'épisode.
  AgeBand ageBand(DateTime now) => switch (ageAt(now)) {
        <= 4 => AgeBand.tiny,
        <= 7 => AgeBand.explorer,
        <= 9 => AgeBand.hero,
        _ => AgeBand.legend,
      };

  ChildProfile copyWith({
    List<String>? passions,
    List<String>? fears,
    List<FamilyMember>? familyCast,
    List<String>? pets,
    ReadingLevel? readingLevel,
    int? storyMinutes,
    SleepTime? targetSleepTime,
  }) =>
      ChildProfile(
        id: id,
        firstName: firstName,
        birthDate: birthDate,
        passions: passions ?? this.passions,
        fears: fears ?? this.fears,
        familyCast: familyCast ?? this.familyCast,
        pets: pets ?? this.pets,
        readingLevel: readingLevel ?? this.readingLevel,
        storyMinutes: storyMinutes ?? this.storyMinutes,
        targetSleepTime: targetSleepTime ?? this.targetSleepTime,
      );
}

/// Une heure de la journée, sans date — miroir du type SQL `time`.
/// (Le domaine reste indépendant de Material : pas de `TimeOfDay` ici.)
@immutable
class SleepTime {
  const SleepTime(this.hour, this.minute);

  final int hour;
  final int minute;

  /// Lit un `time` Postgres : « 20:30 » ou « 20:30:00 ».
  static SleepTime parse(String value) {
    final parts = value.split(':');
    return SleepTime(
      int.tryParse(parts.first) ?? 20,
      parts.length > 1 ? (int.tryParse(parts[1]) ?? 0) : 0,
    );
  }

  /// Format `time` attendu par Postgres.
  String toSql() => '${_pad(hour)}:${_pad(minute)}:00';

  /// L'heure de coucher posée sur un jour donné.
  DateTime onDay(DateTime day) =>
      DateTime(day.year, day.month, day.day, hour, minute);

  static String _pad(int v) => v.toString().padLeft(2, '0');

  @override
  bool operator ==(Object other) =>
      other is SleepTime && other.hour == hour && other.minute == minute;

  @override
  int get hashCode => Object.hash(hour, minute);

  @override
  String toString() => '${_pad(hour)}:${_pad(minute)}';
}

enum ReadingLevel { listener, earlyReader, reader }

/// Passage de flambeau (#29) : le ton mûrit avec l'enfant.
enum AgeBand { tiny, explorer, hero, legend }

@immutable
class FamilyMember {
  const FamilyMember({
    required this.realName,
    required this.role,
    required this.heroName,
    this.photoPath,
  });

  final String realName;
  final String role; // papa, maman, mamie, petit frère, chat…
  final String heroName; // son double héroïque canonique
  final String? photoPath;
}
