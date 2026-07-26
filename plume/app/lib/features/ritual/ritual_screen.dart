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

  /// Une respiration complète dure 8 s : une demi-course à l'aller
  /// (4 s d'inspiration) et une demi-course au retour (4 s d'expiration).
  static const _halfBreath = Duration(seconds: 4);

  late final AnimationController _breath;
  int _breathsDone = 0;
  bool _leaving = false;

  @override
  void initState() {
    super.initState();
    // Surtout pas `repeat(reverse: true)` : il alterne en interne et n'émet
    // jamais `completed`/`dismissed` — le compteur resterait bloqué à zéro.
    // On pilote donc le cycle à la main, aller-retour après aller-retour.
    _breath = AnimationController(vsync: this, duration: _halfBreath)
      ..addStatusListener(_onBreathStatus)
      ..forward();
  }

  void _onBreathStatus(AnimationStatus status) {
    if (_leaving) return;

    // Fin de l'inspiration : la lanterne est pleine, on souffle.
    if (status == AnimationStatus.completed) {
      _breath.reverse();
      return;
    }

    // Fin de l'expiration : une respiration complète de plus.
    if (status == AnimationStatus.dismissed) {
      HapticFeedback.lightImpact();
      setState(() => _breathsDone++);
      if (_breathsDone >= _breathsNeeded) {
        _leaving = true;
        _breath.stop();
        _goToStory();
      } else {
        _breath.forward();
      }
    }
  }

  /// Le seuil est franchi : on entre dans l'histoire. Garde `mounted` —
  /// un listener d'animation peut survivre à la sortie de l'écran.
  void _goToStory() {
    if (!mounted) return;
    context.pushReplacement(Routes.story);
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
                onPressed: () {
                  if (_leaving) return;
                  _leaving = true;
                  _breath.stop();
                  _goToStory();
                },
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
