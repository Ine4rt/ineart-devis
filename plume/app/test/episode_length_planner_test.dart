import 'package:flutter_test/flutter_test.dart';
import 'package:plume/domain/entities/child_profile.dart';
import 'package:plume/domain/usecases/episode_length_planner.dart';

void main() {
  const planner = EpisodeLengthPlanner();
  final now = DateTime(2026, 7, 25, 20, 0);

  group('EpisodeLengthPlanner — le Pacte de sommeil', () {
    test('donne la durée idéale quand il y a le temps', () {
      final planned = planner.plan(
        ageBand: AgeBand.explorer,
        now: now,
        targetSleepTime: now.add(const Duration(minutes: 30)),
      );
      expect(planned, const Duration(minutes: 9));
    });

    test('raccourcit l\'épisode quand le coucher approche', () {
      final planned = planner.plan(
        ageBand: AgeBand.legend,
        now: now,
        targetSleepTime: now.add(const Duration(minutes: 10)),
      );
      // 10 min - 3 min de marge = 7 min disponibles < 15 min idéales.
      expect(planned, const Duration(minutes: 7));
    });

    test('ne descend jamais sous le minimum : pas d\'histoire au rabais', () {
      final planned = planner.plan(
        ageBand: AgeBand.tiny,
        now: now,
        targetSleepTime: now.add(const Duration(minutes: 2)),
      );
      expect(planned, const Duration(minutes: 4));
    });

    test('explique honnêtement au parent quand l\'épisode est raccourci', () {
      expect(
        planner.parentNote(const Duration(minutes: 9), AgeBand.explorer),
        isNull,
      );
      expect(
        planner.parentNote(const Duration(minutes: 5), AgeBand.explorer),
        contains('5 min'),
      );
    });
  });

  group('ChildProfile.ageBand', () {
    test('le ton mûrit avec l\'enfant (passage de flambeau)', () {
      final child = ChildProfile(
        id: 'c',
        firstName: 'Léo',
        birthDate: DateTime(2020, 3, 14),
      );
      expect(child.ageAt(DateTime(2026, 3, 13)), 5);
      expect(child.ageAt(DateTime(2026, 3, 14)), 6);
      expect(child.ageBand(DateTime(2026, 7, 25)), AgeBand.explorer);
      expect(child.ageBand(DateTime(2030, 7, 25)), AgeBand.legend);
    });
  });
}
