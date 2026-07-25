import 'package:flutter_riverpod/flutter_riverpod.dart';

import 'data/repositories/demo_world_repository.dart';
import 'domain/entities/child_profile.dart';
import 'domain/entities/companion.dart';
import 'domain/entities/story_episode.dart';
import 'domain/repositories/world_repository.dart';
import 'domain/usecases/episode_length_planner.dart';

/// Repository du monde. En production, surchargé au démarrage par
/// SupabaseWorldRepository (voir main.dart) ; par défaut : monde de démo.
final worldRepositoryProvider = Provider<WorldRepository>(
  (ref) => DemoWorldRepository(),
);

final childProfileProvider = FutureProvider<ChildProfile?>(
  (ref) => ref.watch(worldRepositoryProvider).loadProfile(),
);

final companionProvider = FutureProvider<Companion?>(
  (ref) => ref.watch(worldRepositoryProvider).loadCompanion(),
);

final tonightEpisodeProvider = FutureProvider<StoryEpisode?>(
  (ref) => ref.watch(worldRepositoryProvider).tonightEpisode(),
);

final archiveProvider = FutureProvider<List<StoryEpisode>>(
  (ref) => ref.watch(worldRepositoryProvider).archive(),
);

final episodeLengthPlannerProvider =
    Provider<EpisodeLengthPlanner>((ref) => const EpisodeLengthPlanner());

/// Après 19 h, l'app se calme avec l'enfant : animations plus lentes,
/// amplitudes réduites (docs/02-DESIGN.md §5, règle 6).
final eveningPacingProvider = Provider<double>((ref) {
  final hour = DateTime.now().hour;
  return (hour >= 19 || hour < 6) ? 1.25 : 1.0;
});
