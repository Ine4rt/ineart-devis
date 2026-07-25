import 'dart:async';

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
/// Plein écran, une scène à la fois, texte en karaoké doux, choix illustrés.
/// L'écran s'assombrit scène après scène : l'app accompagne l'endormissement.
class StoryPlayerScreen extends ConsumerStatefulWidget {
  const StoryPlayerScreen({super.key});

  @override
  ConsumerState<StoryPlayerScreen> createState() => _StoryPlayerScreenState();
}

class _StoryPlayerScreenState extends ConsumerState<StoryPlayerScreen> {
  int _sceneIndex = 0;
  String? _pendingChoiceId;

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
                    if (scene.choice != null)
                      _ChoiceRow(
                        choice: scene.choice!,
                        selectedId: _pendingChoiceId,
                        onSelect: (option) => _onChoice(episode, option),
                      )
                    else
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

  Future<void> _onChoice(StoryEpisode episode, ChoiceOption option) async {
    setState(() => _pendingChoiceId = option.id);
    HapticFeedback.lightImpact();
    // Le choix plante une graine narrative dans le canon — pour toujours.
    unawaited(
      ref.read(worldRepositoryProvider).recordChoice(episode.id, option),
    );
    await Future<void>.delayed(const Duration(milliseconds: 650));
    if (mounted) _advance(episode);
  }

  void _advance(StoryEpisode episode) {
    if (_sceneIndex < episode.scenes.length - 1) {
      setState(() {
        _sceneIndex++;
        _pendingChoiceId = null;
      });
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

class _ChoiceRow extends StatelessWidget {
  const _ChoiceRow({
    required this.choice,
    required this.selectedId,
    required this.onSelect,
  });

  final StoryChoice choice;
  final String? selectedId;
  final ValueChanged<ChoiceOption> onSelect;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context).textTheme;
    return Column(
      children: [
        Text(choice.prompt,
            style: theme.titleLarge?.copyWith(color: AppColors.lanternGold),),
        const SizedBox(height: 16),
        Wrap(
          spacing: 14,
          runSpacing: 14,
          alignment: WrapAlignment.center,
          children: [
            for (final option in choice.options)
              _ChoiceMedallion(
                option: option,
                selected: option.id == selectedId,
                dimmed: selectedId != null && option.id != selectedId,
                onTap: () => onSelect(option),
              ),
          ],
        ),
      ],
    ).animate().fadeIn(delay: 600.ms, duration: 500.ms);
  }
}

/// Médaillon de choix : flotte doucement, s'illumine quand choisi,
/// les autres s'estompent. Cible ≥ 64 px.
class _ChoiceMedallion extends StatelessWidget {
  const _ChoiceMedallion({
    required this.option,
    required this.selected,
    required this.dimmed,
    required this.onTap,
  });

  final ChoiceOption option;
  final bool selected;
  final bool dimmed;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return AnimatedOpacity(
      opacity: dimmed ? .3 : 1,
      duration: const Duration(milliseconds: 300),
      child: AnimatedContainer(
        duration: const Duration(milliseconds: 300),
        curve: Curves.easeOutCubic,
        constraints: const BoxConstraints(minHeight: 64),
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(32),
          color: selected
              ? AppColors.lanternGold
              : AppColors.nightCard.withOpacity(.85),
          border: Border.all(
            color: AppColors.lanternGold.withOpacity(selected ? 1 : .35),
          ),
          boxShadow:
              selected ? AppColors.glow(AppColors.lanternGold, opacity: .5) : null,
        ),
        child: Material(
          color: Colors.transparent,
          child: InkWell(
            borderRadius: BorderRadius.circular(32),
            onTap: dimmed ? null : onTap,
            child: Padding(
              padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 18),
              child: Text(
                option.label,
                style: Theme.of(context).textTheme.titleMedium?.copyWith(
                      color: selected ? AppColors.abyss : AppColors.starWhite,
                    ),
              ),
            ),
          ),
        ),
      ),
    );
  }
}

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
