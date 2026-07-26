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
        // Aucune interaction pendant le récit : le feuilleton se déroule
        // seul, ouvert par un court rappel de l'épisode précédent.
        scenes: const [
          StoryScene(
            index: 0,
            sfx: 'feuilles',
            text:
                'La dernière fois, dans le jardin… Léo avait découvert un '
                'hérisson sous la haie et lui avait laissé un bol d\'eau. '
                'Et ce soir, l\'histoire continue.',
          ),
          StoryScene(
            index: 1,
            sfx: 'oiseau',
            text:
                'La journée se rangeait doucement. Dans le jardin, la lumière '
                'était devenue dorée, et le petit renard attendait Léo près de '
                'la porte, comme chaque soir. Près de la haie… le bol était '
                'vide. Le hérisson était revenu !',
          ),
          StoryScene(
            index: 2,
            sfx: 'craquement',
            text:
                '« Il lui faut un nom », a murmuré Léo. Et comme le hérisson '
                'se grattait le bout du nez, le nom est venu tout seul : '
                'Grattouille. Grattouille a mangé une feuille de salade, puis '
                'il est reparti par un petit passage, tout au fond de la haie.',
          ),
          StoryScene(
            index: 3,
            sfx: 'nuit',
            text:
                'Dedans, la maison sentait le soir. Léo a regardé une dernière '
                'fois le jardin par la fenêtre : la haie, la nuit bleue, et '
                'quelque part là-dessous, Grattouille qui s\'endormait aussi. '
                'Où menait ce passage secret ? Demain soir, le jardin le dirait.',
          ),
        ],
      );

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
