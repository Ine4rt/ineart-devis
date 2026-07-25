import 'package:flutter/material.dart';
import 'package:flutter_animate/flutter_animate.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';

import '../../core/router/app_router.dart';
import '../../core/theme/app_colors.dart';
import '../../core/widgets/glow_card.dart';
import '../../core/widgets/lantern_button.dart';
import '../../core/widgets/starfield_background.dart';
import '../../providers.dart';

/// La Carte du Monde — hub enfant. Pas de tab bar : la carte EST le menu.
/// Les lieux débloqués sont éclairés ; l'épisode du soir attend sous la
/// Lanterne. Le compagnon se promène librement (design §4.3).
class WorldMapScreen extends ConsumerWidget {
  const WorldMapScreen({super.key});

  @override
  Widget build(BuildContext context, WidgetRef ref) {
    final profile = ref.watch(childProfileProvider);
    final companion = ref.watch(companionProvider);
    final episode = ref.watch(tonightEpisodeProvider);
    final theme = Theme.of(context).textTheme;

    return Scaffold(
      body: StarfieldBackground(
        child: SafeArea(
          child: Padding(
            padding: const EdgeInsets.all(24),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                // En-tête : le monde de l'enfant + accès parent (la Lune).
                Row(
                  children: [
                    Expanded(
                      child: profile.when(
                        data: (p) => Text(
                          'Le monde de ${p?.firstName ?? '…'}',
                          style: theme.headlineLarge,
                        ).animate().fadeIn(duration: 600.ms).slideY(begin: .1),
                        loading: () => const SizedBox.shrink(),
                        error: (_, __) => Text('Le monde', style: theme.headlineLarge),
                      ),
                    ),
                    IconButton(
                      onPressed: () => context.push(Routes.parent),
                      tooltip: 'Espace parent',
                      iconSize: 32,
                      icon: const Icon(Icons.nightlight_round,
                          color: AppColors.mist,),
                    ),
                  ],
                ),
                const SizedBox(height: 8),
                companion.when(
                  data: (c) => c == null
                      ? const SizedBox.shrink()
                      : Text(
                          '${c.name} le ${c.dna.species} t\'attend '
                          '· lien niveau ${c.bondLevel}',
                          style: theme.bodySmall,
                        ),
                  loading: () => const SizedBox.shrink(),
                  error: (_, __) => const SizedBox.shrink(),
                ),
                const Spacer(),

                // Les lieux du monde (le menu diégétique).
                _WorldPlace(
                  emoji: '🦊',
                  title: 'Le Foyer du Compagnon',
                  subtitle: 'Il a appris un nouveau mot…',
                  glow: AppColors.duskLavender,
                  onTap: () => context.push(Routes.companion),
                ),
                const SizedBox(height: 14),
                _WorldPlace(
                  emoji: '📚',
                  title: 'La Grande Bibliothèque',
                  subtitle: 'Toute ton aventure, épisode par épisode',
                  glow: AppColors.fireflyGreen,
                  onTap: () => context.push(Routes.album),
                ),
                const SizedBox(height: 28),

                // La Lanterne : l'épisode du soir.
                Center(
                  child: episode.when(
                    data: (ep) => Column(
                      children: [
                        if (ep != null)
                          Text(
                            'Épisode ${ep.number} — « ${ep.title} »',
                            style: theme.titleMedium
                                ?.copyWith(color: AppColors.lanternGold),
                            textAlign: TextAlign.center,
                          ).animate().fadeIn(delay: 300.ms),
                        const SizedBox(height: 14),
                        LanternButton(
                          label: ep == null
                              ? 'Le monde se prépare…'
                              : 'C\'est l\'heure de l\'histoire',
                          icon: Icons.auto_stories_rounded,
                          onPressed: () => context.push(Routes.ritual),
                        ),
                      ],
                    ),
                    loading: () =>
                        Text('Les Archivistes écrivent…', style: theme.bodySmall),
                    error: (_, __) => Text(
                      'Les nuages cachent le monde… la Bibliothèque reste ouverte.',
                      style: theme.bodySmall,
                      textAlign: TextAlign.center,
                    ),
                  ),
                ),
                const SizedBox(height: 12),
              ],
            ),
          ),
        ),
      ),
    );
  }
}

class _WorldPlace extends StatelessWidget {
  const _WorldPlace({
    required this.emoji,
    required this.title,
    required this.subtitle,
    required this.glow,
    required this.onTap,
  });

  final String emoji;
  final String title;
  final String subtitle;
  final Color glow;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context).textTheme;
    return GlowCard(
      glowColor: glow,
      onTap: onTap,
      child: Row(
        children: [
          Text(emoji, style: const TextStyle(fontSize: 34)),
          const SizedBox(width: 16),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(title, style: theme.titleMedium),
                const SizedBox(height: 2),
                Text(subtitle, style: theme.bodySmall),
              ],
            ),
          ),
          const Icon(Icons.chevron_right_rounded, color: AppColors.mist),
        ],
      ),
    ).animate().fadeIn(duration: 500.ms).slideY(begin: .06, curve: Curves.easeOutCubic);
  }
}
