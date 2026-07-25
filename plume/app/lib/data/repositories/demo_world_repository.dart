import 'package:uuid/uuid.dart';

import '../../domain/entities/child_profile.dart';
import '../../domain/entities/companion.dart';
import '../../domain/entities/emotion_checkin.dart';
import '../../domain/entities/narrative_seed.dart';
import '../../domain/entities/story_episode.dart';
import '../../domain/repositories/world_repository.dart';
import '../../domain/usecases/companion_genesis.dart';

/// Monde de démonstration : permet de lancer l'app sans backend
/// (revue design, golden tests, démos investisseurs).
class DemoWorldRepository implements WorldRepository {
  DemoWorldRepository() {
    final dna = const CompanionGenesis()
        .generate(childId: 'demo-child', salt: 42, passions: ['espace']);
    _companion = Companion(id: 'demo-companion', name: 'Pipo', dna: dna, bondLevel: 213);
  }

  static const _uuid = Uuid();

  ChildProfile _profile = ChildProfile(
    id: 'demo-child',
    firstName: 'Léo',
    birthDate: DateTime(2020, 3, 14),
    passions: const ['espace', 'dinosaures'],
    fears: const ['le noir'],
    pets: const ['Moustache (chat)'],
  );

  late Companion _companion;
  final List<NarrativeSeed> _seeds = [
    NarrativeSeed(
      id: 'seed-dragon',
      kind: SeedKind.choiceConsequence,
      summary: 'Léo a épargné le dragon Cendreflamme au pont des Brumes.',
      plantedAt: DateTime(2026, 1, 12),
      germinateAfter: DateTime(2026, 6, 1),
    ),
  ];
  final List<StoryEpisode> _archive = [];

  @override
  Future<ChildProfile?> loadProfile() async => _profile;

  @override
  Future<void> saveProfile(ChildProfile profile) async => _profile = profile;

  @override
  Future<Companion?> loadCompanion() async => _companion;

  @override
  Future<void> saveCompanion(Companion companion) async => _companion = companion;

  @override
  Future<StoryEpisode?> tonightEpisode() async => StoryEpisode(
        id: 'demo-ep-214',
        number: 214,
        title: 'Le retour de Cendreflamme',
        date: DateTime.now(),
        emotionalThread: 'courage après une journée difficile',
        isDownloaded: true,
        scenes: const [
          StoryScene(
            index: 0,
            text:
                'Ce soir-là, une ombre douce glissa au-dessus du village de Lunelune. '
                'Pipo dressa ses oreilles-feuilles : il connaissait ce battement '
                'd\'ailes. Quelqu\'un revenait de très, très loin…',
          ),
          StoryScene(
            index: 1,
            text:
                'C\'était Cendreflamme ! Le dragon que Léo avait épargné au pont des '
                'Brumes, il y a si longtemps. Ses écailles avaient poussé, et dans '
                'ses yeux brillait quelque chose de nouveau : de la gratitude.',
            choice: StoryChoice(
              prompt: 'Que fait Léo ?',
              options: [
                ChoiceOption(
                  id: 'opt-welcome',
                  label: 'Courir l\'accueillir',
                  seedSummary: 'Léo et Cendreflamme deviennent alliés du village.',
                ),
                ChoiceOption(
                  id: 'opt-hide',
                  label: 'Observer caché avec Pipo',
                  seedSummary: 'Léo découvre le secret que porte Cendreflamme.',
                ),
              ],
            ),
          ),
          StoryScene(
            index: 2,
            text:
                'Le dragon posa délicatement une écaille dorée dans la main de Léo. '
                '« Tu m\'as sauvé un jour. Cette nuit, c\'est moi qui veille sur toi. » '
                'Et au-dessus de Lunelune, les étoiles se mirent à ronronner.',
          ),
        ],
      );

  @override
  Future<void> recordChoice(String episodeId, ChoiceOption option) async {
    _seeds.add(
      NarrativeSeed(
        id: _uuid.v4(),
        kind: SeedKind.choiceConsequence,
        summary: option.seedSummary,
        plantedAt: DateTime.now(),
        germinateAfter: DateTime.now().add(const Duration(days: 45)),
      ),
    );
  }

  @override
  Future<void> saveBookmark(String episodeId, SleepBookmark bookmark) async {}

  @override
  Future<void> submitCheckin(EmotionCheckin checkin) async {}

  @override
  Future<List<NarrativeSeed>> pendingSeeds() async => List.unmodifiable(_seeds);

  @override
  Future<List<StoryEpisode>> archive({int limit = 30}) async =>
      _archive.take(limit).toList();
}
