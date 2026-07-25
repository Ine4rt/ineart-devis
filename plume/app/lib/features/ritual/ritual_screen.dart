import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';

import '../../core/router/app_router.dart';
import '../../core/theme/app_colors.dart';
import '../../core/widgets/starfield_background.dart';

/// Le Rituel du Soir (innovation #18) — ici, l'étape « Calme » :
/// trois respirations guidées par la lanterne qui gonfle et dégonfle.
/// L'écran se réchauffe progressivement (moins de lumière bleue).
class RitualScreen extends ConsumerStatefulWidget {
  const RitualScreen({super.key});

  @override
  ConsumerState<RitualScreen> createState() => _RitualScreenState();
}

class _RitualScreenState extends ConsumerState<RitualScreen>
    with SingleTickerProviderStateMixin {
  static const _breathsNeeded = 3;
  static const _breathDuration = Duration(seconds: 8); // 4 s inspire, 4 s expire

  late final AnimationController _breath;
  int _breathsDone = 0;

  @override
  void initState() {
    super.initState();
    _breath = AnimationController(vsync: this, duration: _breathDuration)
      ..addStatusListener(_onBreathStatus)
      ..repeat(reverse: true);
  }

  void _onBreathStatus(AnimationStatus status) {
    // Une respiration complète = un aller-retour (fin de l'expiration).
    if (status == AnimationStatus.dismissed) {
      HapticFeedback.lightImpact();
      setState(() => _breathsDone++);
      if (_breathsDone >= _breathsNeeded) {
        _breath.stop();
        context.pushReplacement(Routes.story);
      }
    }
  }

  @override
  void dispose() {
    _breath.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context).textTheme;
    return Scaffold(
      body: StarfieldBackground(
        fireflyCount: 3,
        child: SafeArea(
          child: Column(
            children: [
              const SizedBox(height: 24),
              Text('Le seuil de la nuit', style: theme.headlineMedium),
              const SizedBox(height: 8),
              Text(
                'Trois grandes respirations avec la lanterne…',
                style: theme.bodySmall,
              ),
              Expanded(
                child: Center(
                  child: AnimatedBuilder(
                    animation: _breath,
                    builder: (context, _) {
                      final t = Curves.easeInOut.transform(_breath.value);
                      final size = 120 + 80 * t;
                      final inhaling = _breath.status == AnimationStatus.forward;
                      return Column(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          Container(
                            width: size,
                            height: size,
                            decoration: BoxDecoration(
                              shape: BoxShape.circle,
                              gradient: RadialGradient(
                                colors: [
                                  AppColors.lanternGold.withOpacity(.9),
                                  AppColors.lanternGold.withOpacity(.0),
                                ],
                              ),
                              boxShadow: AppColors.glow(
                                AppColors.lanternGold,
                                opacity: .2 + .3 * t,
                              ),
                            ),
                          ),
                          const SizedBox(height: 40),
                          Text(
                            inhaling ? 'On remplit la lanterne…' : 'On souffle doucement…',
                            style: theme.titleLarge,
                          ),
                        ],
                      );
                    },
                  ),
                ),
              ),
              // Compteur diégétique : des lucioles s'allument, pas des chiffres.
              Padding(
                padding: const EdgeInsets.only(bottom: 40),
                child: Row(
                  mainAxisAlignment: MainAxisAlignment.center,
                  children: [
                    for (var i = 0; i < _breathsNeeded; i++)
                      Padding(
                        padding: const EdgeInsets.symmetric(horizontal: 8),
                        child: AnimatedContainer(
                          duration: const Duration(milliseconds: 400),
                          width: 14,
                          height: 14,
                          decoration: BoxDecoration(
                            shape: BoxShape.circle,
                            color: i < _breathsDone
                                ? AppColors.lanternGold
                                : AppColors.nightCard,
                            boxShadow: i < _breathsDone
                                ? AppColors.glow(AppColors.lanternGold, opacity: .5)
                                : null,
                          ),
                        ),
                      ),
                  ],
                ),
              ),
              TextButton(
                onPressed: () => context.pushReplacement(Routes.story),
                child: Text('Passer', style: theme.bodySmall),
              ),
              const SizedBox(height: 8),
            ],
          ),
        ),
      ),
    );
  }
}
