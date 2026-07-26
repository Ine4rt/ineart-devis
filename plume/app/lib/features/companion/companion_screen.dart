import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_animate/flutter_animate.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';

import '../../core/theme/app_colors.dart';
import '../../core/widgets/glow_card.dart';
import '../../core/widgets/starfield_background.dart';
import '../../domain/entities/companion.dart';
import '../../providers.dart';

/// Le Foyer du Compagnon (innovation #14) : il réagit au toucher, montre son
/// ADN unique et les mots qu'il a appris. Pas de jauges anxiogènes — il va
/// toujours bien, il a juste des envies.
class CompanionScreen extends ConsumerStatefulWidget {
  const CompanionScreen({super.key});

  @override
  ConsumerState<CompanionScreen> createState() => _CompanionScreenState();
}

class _CompanionScreenState extends ConsumerState<CompanionScreen> {
  int _pats = 0;

  static const _reactions = [
    'ronronne doucement',
    'fait une pirouette !',
    'te regarde avec des yeux pleins d\'étoiles',
    'cache son nez sous sa queue, tout content',
  ];

  @override
  Widget build(BuildContext context) {
    final companionAsync = ref.watch(companionProvider);
    final theme = Theme.of(context).textTheme;

    return Scaffold(
      body: StarfieldBackground(
        fireflyCount: 10,
        child: SafeArea(
          child: companionAsync.when(
            loading: () => const Center(child: Text('🥚', style: TextStyle(fontSize: 64))),
            error: (_, __) => const SizedBox.shrink(),
            data: (companion) {
              if (companion == null) return const SizedBox.shrink();
              return Padding(
                padding: const EdgeInsets.all(24),
                child: Column(
                  children: [
                    Align(
                      alignment: Alignment.centerLeft,
                      child: BackButton(onPressed: context.pop),
                    ),
                    Text(companion.name, style: theme.displayLarge),
                    Text(
                      '${companion.dna.species} · ${companion.dna.trait}\n'
                      '${_stageLabel(companion.stage)} · lien niveau ${companion.bondLevel}',
                      style: theme.bodySmall,
                      textAlign: TextAlign.center,
                    ),
                    const Spacer(),
                    // Le compagnon (placeholder emoji — rendu final : rive/lottie
                    // généré depuis la fiche visuelle de l'ADN).
                    GestureDetector(
                      onTap: () {
                        HapticFeedback.lightImpact();
                        setState(() => _pats++);
                      },
                      child: const Text('🦊', style: TextStyle(fontSize: 120))
                          .animate(key: ValueKey(_pats))
                          .scale(
                            begin: const Offset(1, 1),
                            end: const Offset(1.12, 1.12),
                            duration: 180.ms,
                            curve: Curves.easeOutBack,
                          )
                          .then()
                          .scale(
                            begin: const Offset(1.12, 1.12),
                            end: const Offset(1, 1),
                            duration: 220.ms,
                          ),
                    ),
                    const SizedBox(height: 16),
                    AnimatedSwitcher(
                      duration: const Duration(milliseconds: 300),
                      child: Text(
                        _pats == 0
                            ? '${companion.name} t\'attendait…'
                            : '${companion.name} ${_reactions[(_pats - 1) % _reactions.length]}',
                        key: ValueKey(_pats),
                        style: theme.titleLarge,
                        textAlign: TextAlign.center,
                      ),
                    ),
                    const Spacer(),
                    if (companion.learnedWords.isNotEmpty)
                      GlowCard(
                        glowColor: AppColors.duskLavender,
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.start,
                          children: [
                            Text('Les mots qu\'il a appris de toi',
                                style: theme.titleMedium,),
                            const SizedBox(height: 8),
                            Wrap(
                              spacing: 8,
                              children: [
                                for (final word in companion.learnedWords)
                                  Chip(
                                    label: Text(word),
                                    backgroundColor:
                                        AppColors.duskLavender.withOpacity(.2),
                                  ),
                              ],
                            ),
                          ],
                        ),
                      ),
                    const SizedBox(height: 16),
                  ],
                ),
              );
            },
          ),
        ),
      ),
    );
  }

  String _stageLabel(GrowthStage stage) => switch (stage) {
        GrowthStage.hatchling => 'tout jeune éclos',
        GrowthStage.sprout => 'petit curieux',
        GrowthStage.wanderer => 'grand explorateur',
        GrowthStage.guardian => 'gardien du monde',
        GrowthStage.ancient => 'légende vivante',
      };
}
