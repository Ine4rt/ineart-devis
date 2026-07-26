import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_animate/flutter_animate.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';

import '../../core/theme/app_colors.dart';
import '../../core/theme/app_typography.dart';
import '../../core/widgets/starfield_background.dart';
import '../../domain/entities/story_episode.dart';
import '../../providers.dart';

/// Le Théâtre de l'Histoire — l'écran le plus important de Plume.
/// Plein écran, une scène à la fois, texte en karaoké doux.
/// L'écran s'assombrit scène après scène : l'app accompagne l'endormissement.
///
/// AUCUNE INTERACTION pendant l'histoire (décision produit) : pas de choix,
/// pas de question posée à l'enfant. La seule commande est « avancer » —
/// c'est le canon et les journées réelles de l'enfant qui font l'histoire.
class StoryPlayerScreen extends ConsumerStatefulWidget {
  const StoryPlayerScreen({super.key});

  @override
  ConsumerState<StoryPlayerScreen> createState() => _StoryPlayerScreenState();
}

class _StoryPlayerScreenState extends ConsumerState<StoryPlayerScreen> {
  int _sceneIndex = 0;

  @override
  void initState() {
    super.initState();
    // Immersion totale : plus de barres système pendant l'histoire.
    SystemChrome.setEnabledSystemUIMode(SystemUiMode.immersiveSticky);
  }

  @override
  void dispose() {
    SystemChrome.setEnabledSystemUIMode(SystemUiMode.edgeToEdge);
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final episodeAsync = ref.watch(tonightEpisodeProvider);

    return Scaffold(
      backgroundColor: AppColors.abyss,
      body: episodeAsync.when(
        loading: () => const _ParchmentLoading(),
        error: (_, __) => const _ParchmentLoading(),
        data: (episode) {
          if (episode == null) return const _ParchmentLoading();
          final scene = episode.scenes[_sceneIndex];
          final progress = (_sceneIndex + 1) / episode.scenes.length;

          return StarfieldBackground(
            // Le ciel se vide d'étoiles au fil de l'histoire : la nuit tombe.
            starCount: (90 * (1 - progress * .6)).round(),
            fireflyCount: 3,
            child: SafeArea(
              child: Padding(
                padding: const EdgeInsets.symmetric(horizontal: 28, vertical: 16),
                child: Column(
                  children: [
                    _EpisodeHeader(episode: episode, progress: progress),
                    const Spacer(),
                    // Le texte de la scène, qui naît (jamais n'apparaît).
                    Text(
                      scene.text,
                      textAlign: TextAlign.center,
                      style: AppTypography.storyText,
                    )
                        .animate(key: ValueKey(_sceneIndex))
                        .fadeIn(duration: 900.ms, curve: Curves.easeOutCubic)
                        .slideY(begin: .04),
                    const Spacer(),
                    _NextSceneButton(
                      isLast: _sceneIndex == episode.scenes.length - 1,
                      onTap: () => _advance(episode),
                    ),
                    const SizedBox(height: 24),
                  ],
                ),
              ),
            ),
          );
        },
      ),
    );
  }

  void _advance(StoryEpisode episode) {
    if (_sceneIndex < episode.scenes.length - 1) {
      setState(() => _sceneIndex++);
    } else {
      // Fin d'épisode : retour au monde, qui aura changé demain.
      context.pop();
    }
  }
}

class _EpisodeHeader extends StatelessWidget {
  const _EpisodeHeader({required this.episode, required this.progress});

  final StoryEpisode episode;
  final double progress;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context).textTheme;
    return Column(
      children: [
        Text('Épisode ${episode.number}', style: theme.bodySmall),
        const SizedBox(height: 4),
        Text(episode.title, style: theme.headlineMedium, textAlign: TextAlign.center),
        const SizedBox(height: 12),
        // Progression = un fil d'or qui se tisse, pas une barre de chargement.
        ClipRRect(
          borderRadius: BorderRadius.circular(2),
          child: LinearProgressIndicator(
            value: progress,
            minHeight: 3,
            backgroundColor: AppColors.nightCard,
            valueColor: const AlwaysStoppedAnimation(AppColors.lanternGold),
          ),
        ),
      ],
    );
  }
}

/// La seule commande de l'écran : avancer d'une scène, puis s'endormir.
class _NextSceneButton extends StatelessWidget {
  const _NextSceneButton({required this.isLast, required this.onTap});

  final bool isLast;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return TextButton.icon(
      onPressed: onTap,
      icon: Icon(
        isLast ? Icons.bedtime_rounded : Icons.arrow_forward_rounded,
        color: AppColors.mist,
      ),
      label: Text(
        isLast ? 'Bonne nuit…' : 'Ensuite…',
        style: Theme.of(context)
            .textTheme
            .titleMedium
            ?.copyWith(color: AppColors.mist),
      ),
    ).animate().fadeIn(delay: 2.seconds, duration: 800.ms);
  }
}

/// Jamais de spinner : Plume écrit sur un parchemin (design §4.9).
class _ParchmentLoading extends StatelessWidget {
  const _ParchmentLoading();

  @override
  Widget build(BuildContext context) {
    return StarfieldBackground(
      child: Center(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            const Text('🪶', style: TextStyle(fontSize: 56))
                .animate(onPlay: (c) => c.repeat(reverse: true))
                .moveY(begin: -6, end: 6, duration: 1.2.seconds, curve: Curves.easeInOut),
            const SizedBox(height: 16),
            Text(
              'Les Archivistes terminent la page…',
              style: Theme.of(context).textTheme.bodyLarge,
            ),
          ],
        ),
      ),
    );
  }
}
