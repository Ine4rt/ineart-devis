import 'package:flutter/material.dart';
import 'package:flutter_animate/flutter_animate.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';
import 'package:intl/intl.dart';

import '../../core/theme/app_colors.dart';
import '../../core/widgets/glow_card.dart';
import '../../core/widgets/starfield_background.dart';
import '../../providers.dart';

/// La Grande Bibliothèque (design §4.7) : chaque épisode est un livre relié.
/// À terme : rayonnages par mois, atelier coloriage, capsules temporelles,
/// et le Livre de l'Année imprimé.
class AlbumScreen extends ConsumerWidget {
  const AlbumScreen({super.key});

  @override
  Widget build(BuildContext context, WidgetRef ref) {
    final archive = ref.watch(archiveProvider);
    final theme = Theme.of(context).textTheme;
    final dateFormat = DateFormat('d MMM');

    return Scaffold(
      body: StarfieldBackground(
        fireflyCount: 4,
        child: SafeArea(
          child: Padding(
            padding: const EdgeInsets.all(24),
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                BackButton(onPressed: context.pop),
                Text('La Grande Bibliothèque', style: theme.headlineLarge),
                const SizedBox(height: 4),
                Text(
                  'Toute ton aventure, écrite par les Archivistes.',
                  style: theme.bodySmall,
                ),
                const SizedBox(height: 24),
                Expanded(
                  child: archive.when(
                    loading: () => const SizedBox.shrink(),
                    error: (_, __) => const SizedBox.shrink(),
                    data: (episodes) {
                      if (episodes.isEmpty) {
                        return Center(
                          child: Column(
                            mainAxisSize: MainAxisSize.min,
                            children: [
                              const Text('📖', style: TextStyle(fontSize: 56)),
                              const SizedBox(height: 12),
                              Text(
                                'Les rayonnages attendent ton premier épisode…\n'
                                'Il s\'écrira ce soir.',
                                style: theme.bodyLarge,
                                textAlign: TextAlign.center,
                              ),
                            ],
                          ),
                        ).animate().fadeIn(duration: 600.ms);
                      }
                      return ListView.separated(
                        itemCount: episodes.length,
                        separatorBuilder: (_, __) => const SizedBox(height: 12),
                        itemBuilder: (context, i) {
                          final ep = episodes[i];
                          return GlowCard(
                            glowColor: AppColors.fireflyGreen,
                            child: Row(
                              children: [
                                Text('📕', style: theme.headlineMedium),
                                const SizedBox(width: 14),
                                Expanded(
                                  child: Column(
                                    crossAxisAlignment: CrossAxisAlignment.start,
                                    children: [
                                      Text('Épisode ${ep.number} — ${ep.title}',
                                          style: theme.titleMedium,),
                                      Text(dateFormat.format(ep.date),
                                          style: theme.bodySmall,),
                                    ],
                                  ),
                                ),
                              ],
                            ),
                          ).animate(delay: (60 * i).ms).fadeIn().slideY(begin: .05);
                        },
                      );
                    },
                  ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
