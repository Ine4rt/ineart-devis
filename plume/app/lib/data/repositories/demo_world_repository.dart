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
      id: 'seed-herisson',
      kind: SeedKind.choiceConsequence,
      summary: 'Léo a laissé un bol d\'eau au hérisson de la haie.',
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

  // Registre « réalisme doux » : monde réel, poésie du quotidien —
  // la magie vient de la mémoire du monde, pas de créatures fantastiques.
  @override
  Future<StoryEpisode?> tonightEpisode() async => StoryEpisode(
        id: 'demo-ep-214',
        number: 214,
        title: 'Le jardin du soir',
        date: DateTime.now(),
        emotionalThread: 'courage après une journée difficile',
        isDownloaded: true,
        scenes: const [
          StoryScene(
            index: 0,
            text:
                'La journée se rangeait doucement. Dans le jardin, la lumière '
                'était devenue dorée, et le petit renard attendait Léo près de '
                'la porte, comme chaque soir.',
          ),
          StoryScene(
            index: 1,
            text:
                'Au fond du jardin, la haie s\'est mise à remuer. Le petit renard '
                's\'est arrêté net, une oreille levée. Là, sous les feuilles : un '
                'hérisson, tout petit, qui cherchait son chemin.',
            choice: StoryChoice(
              prompt: 'Que fait Léo ?',
              options: [
                ChoiceOption(
                  id: 'opt-water',
                  label: 'Lui laisser un bol d\'eau',
                  seedSummary:
                      'Le hérisson reviendra boire chaque soir de l\'été.',
                ),
                ChoiceOption(
                  id: 'opt-watch',
                  label: 'L\'observer sans faire de bruit',
                  seedSummary:
                      'Léo connaît maintenant le passage secret de la haie.',
                ),
              ],
            ),
          ),
          StoryScene(
            index: 2,
            text:
                'Dedans, la maison sentait le soir. Léo a regardé une dernière '
                'fois le jardin par la fenêtre : la haie, la nuit bleue, et '
                'quelque part là-dessous, un hérisson qui s\'endormait aussi. '
                'Demain, le jardin s\'en souviendrait.',
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
