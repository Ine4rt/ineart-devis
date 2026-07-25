import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';

import '../../core/theme/app_colors.dart';
import '../../domain/entities/emotion_checkin.dart';
import '../../providers.dart';

/// La Météo émotionnelle (innovation #7) : 4 questions tactiles, 20 secondes
/// maximum. Ce que le parent saisit ici est tissé en filigrane dans l'épisode
/// du soir — jamais frontalement.
class CheckinScreen extends ConsumerStatefulWidget {
  const CheckinScreen({super.key});

  @override
  ConsumerState<CheckinScreen> createState() => _CheckinScreenState();
}

class _CheckinScreenState extends ConsumerState<CheckinScreen> {
  DayQuality? _quality;
  Emotion? _emotion;
  final _victoryController = TextEditingController();
  final _fearController = TextEditingController();
  bool _sent = false;

  @override
  void dispose() {
    _victoryController.dispose();
    _fearController.dispose();
    super.dispose();
  }

  Future<void> _submit() async {
    final profile = await ref.read(childProfileProvider.future);
    if (profile == null || _quality == null) return;
    await ref.read(worldRepositoryProvider).submitCheckin(
          EmotionCheckin(
            childId: profile.id,
            date: DateTime.now(),
            dayQuality: _quality!,
            dominantEmotion: _emotion,
            victory: _victoryController.text.trim().isEmpty
                ? null
                : _victoryController.text.trim(),
            fearOfTheDay: _fearController.text.trim().isEmpty
                ? null
                : _fearController.text.trim(),
          ),
        );
    if (mounted) setState(() => _sent = true);
  }

  @override
  Widget build(BuildContext context) {
    // Espace parent : thème jour, dense, une main (design §4.8).
    final theme = Theme.of(context).textTheme;
    return Theme(
      data: ThemeData.light(useMaterial3: true),
      child: Scaffold(
        backgroundColor: AppColors.dawnLinen,
        appBar: AppBar(
          backgroundColor: AppColors.dawnLinen,
          leading: BackButton(onPressed: context.pop, color: AppColors.dayInk),
          title: Text('La météo du jour',
              style: theme.titleLarge?.copyWith(color: AppColors.dayInk)),
        ),
        body: _sent
            ? Center(
                child: Padding(
                  padding: const EdgeInsets.all(32),
                  child: Text(
                    'Merci. Les Archivistes tissent tout cela dans l\'épisode '
                    'de ce soir — en filigrane, comme toujours. 🪶',
                    style: theme.titleMedium?.copyWith(color: AppColors.dayInk),
                    textAlign: TextAlign.center,
                  ),
                ),
              )
            : ListView(
                padding: const EdgeInsets.all(20),
                children: [
                  _Question(
                    label: 'Comment s\'est passée sa journée ?',
                    child: Wrap(
                      spacing: 8,
                      children: [
                        for (final q in DayQuality.values)
                          ChoiceChip(
                            label: Text(_qualityLabel(q)),
                            selected: _quality == q,
                            onSelected: (_) {
                              HapticFeedback.selectionClick();
                              setState(() => _quality = q);
                            },
                          ),
                      ],
                    ),
                  ),
                  _Question(
                    label: 'L\'émotion dominante ?',
                    child: Wrap(
                      spacing: 8,
                      children: [
                        for (final e in Emotion.values)
                          ChoiceChip(
                            label: Text(_emotionLabel(e)),
                            selected: _emotion == e,
                            onSelected: (_) => setState(() => _emotion = e),
                          ),
                      ],
                    ),
                  ),
                  _Question(
                    label: 'A-t-il/elle réussi quelque chose ? (optionnel)',
                    child: TextField(
                      controller: _victoryController,
                      decoration: const InputDecoration(
                        hintText: 'Ex. : du vélo sans les petites roues !',
                        border: OutlineInputBorder(),
                      ),
                    ),
                  ),
                  _Question(
                    label: 'Une peur ou un moment difficile ? (optionnel)',
                    child: TextField(
                      controller: _fearController,
                      decoration: const InputDecoration(
                        hintText: 'Ex. : le rendez-vous chez le dentiste',
                        border: OutlineInputBorder(),
                      ),
                    ),
                  ),
                  const SizedBox(height: 12),
                  FilledButton(
                    onPressed: _quality == null ? null : _submit,
                    style: FilledButton.styleFrom(
                      backgroundColor: AppColors.dayInk,
                      minimumSize: const Size.fromHeight(52),
                    ),
                    child: const Text('Tisser dans l\'histoire de ce soir'),
                  ),
                ],
              ),
      ),
    );
  }

  String _qualityLabel(DayQuality q) => switch (q) {
        DayQuality.radiant => '☀️ Radieuse',
        DayQuality.good => '🙂 Bonne',
        DayQuality.mixed => '⛅ Mitigée',
        DayQuality.hard => '🌧 Difficile',
      };

  String _emotionLabel(Emotion e) => switch (e) {
        Emotion.joy => 'Joie',
        Emotion.pride => 'Fierté',
        Emotion.calm => 'Calme',
        Emotion.sadness => 'Tristesse',
        Emotion.anger => 'Colère',
        Emotion.fear => 'Peur',
        Emotion.frustration => 'Frustration',
      };
}

class _Question extends StatelessWidget {
  const _Question({required this.label, required this.child});

  final String label;
  final Widget child;

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 20),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            label,
            style: Theme.of(context)
                .textTheme
                .titleMedium
                ?.copyWith(color: AppColors.dayInk),
          ),
          const SizedBox(height: 10),
          child,
        ],
      ),
    );
  }
}
