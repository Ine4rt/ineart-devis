import 'package:flutter_test/flutter_test.dart';
import 'package:plume/domain/entities/companion.dart';
import 'package:plume/domain/usecases/companion_genesis.dart';

void main() {
  const genesis = CompanionGenesis();

  group('CompanionGenesis', () {
    test('est déterministe : même enfant + même sel = même compagnon', () {
      final a = genesis.generate(childId: 'child-1', salt: 7);
      final b = genesis.generate(childId: 'child-1', salt: 7);
      expect(a.species, b.species);
      expect(a.trait, b.trait);
      expect(a.pattern, b.pattern);
      expect(a.paletteSeed, b.paletteSeed);
      expect(a.temperament, b.temperament);
    });

    test('deux enfants différents ont des ADN différents (sur un échantillon)', () {
      // L'espace des ADN est vaste : sur 200 enfants, il doit exister une
      // grande diversité (pas de collision massive).
      final signatures = <String>{};
      for (var i = 0; i < 200; i++) {
        final dna = genesis.generate(childId: 'child-$i', salt: i * 31 + 1);
        signatures.add('${dna.species}|${dna.trait}|${dna.pattern}|${dna.paletteSeed}');
      }
      expect(signatures.length, greaterThan(190));
    });

    test('les passions orientent le tempérament sans le forcer', () {
      final temperaments = <Temperament>{};
      for (var salt = 0; salt < 50; salt++) {
        temperaments.add(
          genesis
              .generate(childId: 'space-kid', salt: salt, passions: ['espace'])
              .temperament,
        );
      }
      // « espace » doit produire souvent un rêveur, mais pas toujours.
      expect(temperaments, contains(Temperament.dreamy));
      expect(temperaments.length, greaterThan(1));
    });
  });

  group('GrowthStage', () {
    test('le compagnon grandit avec le lien, sur des années', () {
      expect(GrowthStage.forBond(0), GrowthStage.hatchling);
      expect(GrowthStage.forBond(13), GrowthStage.hatchling);
      expect(GrowthStage.forBond(14), GrowthStage.sprout);
      expect(GrowthStage.forBond(60), GrowthStage.wanderer);
      expect(GrowthStage.forBond(365), GrowthStage.guardian);
      expect(GrowthStage.forBond(1095), GrowthStage.ancient);
    });

    test('afterEpisode augmente le lien et déduplique les mots appris', () {
      final companion = Companion(
        id: 'c1',
        name: 'Pipo',
        dna: genesis.generate(childId: 'x', salt: 1),
        bondLevel: 13,
        learnedWords: const ['chocotruc'],
      );
      final grown = companion.afterEpisode(newWords: ['chocotruc', 'zoubidou']);
      expect(grown.bondLevel, 14);
      expect(grown.stage, GrowthStage.sprout);
      expect(grown.learnedWords, containsAll(['chocotruc', 'zoubidou']));
      expect(grown.learnedWords.length, 2);
    });
  });
}
