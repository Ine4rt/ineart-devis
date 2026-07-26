import 'package:flutter_test/flutter_test.dart';
import 'package:plume/domain/entities/narrative_seed.dart';

NarrativeSeed seed({
  required String id,
  SeedKind kind = SeedKind.choiceConsequence,
  SeedPriority priority = SeedPriority.normal,
  required DateTime planted,
  required DateTime germinate,
  DateTime? harvested,
}) =>
    NarrativeSeed(
      id: id,
      kind: kind,
      summary: id,
      plantedAt: planted,
      germinateAfter: germinate,
      priority: priority,
      harvestedAt: harvested,
    );

void main() {
  final now = DateTime(2026, 7, 25);

  group('pickSeedsForTonight — le pouvoir des choix à longue portée', () {
    test('ne récolte que les graines mûres et non récoltées', () {
      final seeds = [
        seed(id: 'ripe', planted: DateTime(2026, 1, 1), germinate: DateTime(2026, 6, 1)),
        seed(id: 'green', planted: DateTime(2026, 7, 1), germinate: DateTime(2026, 9, 1)),
        seed(
          id: 'done',
          planted: DateTime(2026, 1, 1),
          germinate: DateTime(2026, 2, 1),
          harvested: DateTime(2026, 3, 1),
        ),
      ];
      final picked = pickSeedsForTonight(seeds, now);
      expect(picked.map((s) => s.id), ['ripe']);
    });

    test('les moments de vie passent devant tout le reste', () {
      final seeds = [
        seed(
          id: 'old-choice',
          priority: SeedPriority.high,
          planted: DateTime(2025, 1, 1),
          germinate: DateTime(2025, 2, 1),
        ),
        seed(
          id: 'rentree',
          kind: SeedKind.lifeMoment,
          planted: DateTime(2026, 7, 20),
          germinate: DateTime(2026, 7, 22),
        ),
      ];
      final picked = pickSeedsForTonight(seeds, now, max: 1);
      expect(picked.single.id, 'rentree');
    });

    test('à priorité égale, une promesse ancienne n\'est jamais oubliée', () {
      final seeds = [
        seed(id: 'recent', planted: DateTime(2026, 6, 1), germinate: DateTime(2026, 7, 1)),
        seed(id: 'ancient', planted: DateTime(2026, 1, 12), germinate: DateTime(2026, 6, 1)),
      ];
      final picked = pickSeedsForTonight(seeds, now, max: 1);
      expect(picked.single.id, 'ancient');
    });

    test('respecte la limite max par épisode', () {
      final seeds = [
        for (var i = 0; i < 5; i++)
          seed(
            id: 's$i',
            planted: DateTime(2026, 1, 1 + i),
            germinate: DateTime(2026, 6, 1),
          ),
      ];
      expect(pickSeedsForTonight(seeds, now).length, 2);
    });
  });
}
