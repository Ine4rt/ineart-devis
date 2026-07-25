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
      );
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
