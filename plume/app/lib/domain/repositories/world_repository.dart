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
  ///
  /// Aucune interaction pendant le récit : il n'existe pas de « choix » à
  /// enregistrer. Les graines narratives sont désormais plantées côté serveur
  /// par `generate-episode`, à partir du canon et de la journée de l'enfant.
  Future<StoryEpisode?> tonightEpisode();

  /// Marque-page d'endormissement (#19).
  Future<void> saveBookmark(String episodeId, SleepBookmark bookmark);

  Future<void> submitCheckin(EmotionCheckin checkin);

  Future<List<NarrativeSeed>> pendingSeeds();

  /// L'album : tous les épisodes archivés, du plus récent au plus ancien.
  Future<List<StoryEpisode>> archive({int limit = 30});
}
