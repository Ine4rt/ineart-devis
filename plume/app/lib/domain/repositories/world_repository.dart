import '../entities/child_profile.dart';
import '../entities/companion.dart';
import '../entities/emotion_checkin.dart';
import '../entities/narrative_seed.dart';
import '../entities/story_episode.dart';

/// Contrat d'accès au monde de l'enfant. Implémentations :
/// - [SupabaseWorldRepository] : production (Postgres + Edge Functions),
/// - [DemoWorldRepository]     : mode démo/offline et tests.
abstract interface class WorldRepository {
  Future<ChildProfile?> loadProfile();
  Future<void> saveProfile(ChildProfile profile);

  Future<Companion?> loadCompanion();
  Future<void> saveCompanion(Companion companion);

  /// L'épisode du soir — pré-généré la nuit, téléchargé en avance.
  Future<StoryEpisode?> tonightEpisode();

  /// Choix de l'enfant → graine narrative plantée dans le canon.
  Future<void> recordChoice(String episodeId, ChoiceOption option);

  /// Marque-page d'endormissement (#19).
  Future<void> saveBookmark(String episodeId, SleepBookmark bookmark);

  Future<void> submitCheckin(EmotionCheckin checkin);

  Future<List<NarrativeSeed>> pendingSeeds();

  /// L'album : tous les épisodes archivés, du plus récent au plus ancien.
  Future<List<StoryEpisode>> archive({int limit = 30});
}
