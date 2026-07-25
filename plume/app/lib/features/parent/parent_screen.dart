import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';
import 'package:intl/intl.dart';

import '../../core/router/app_router.dart';
import '../../core/theme/app_colors.dart';
import '../../providers.dart';

/// La Lune — l'espace parent (design §4.8). Verrou par appui long (2 s),
/// thème jour, informations denses : météo du jour, pacte de sommeil,
/// graines narratives en attente (transparence totale sur ce que sait l'IA).
class ParentScreen extends ConsumerStatefulWidget {
  const ParentScreen({super.key});

  @override
  ConsumerState<ParentScreen> createState() => _ParentScreenState();
}

class _ParentScreenState extends ConsumerState<ParentScreen> {
  bool _unlocked = false;

  @override
  Widget build(BuildContext context) {
    return _unlocked ? const _ParentDashboard() : _ParentLock(onUnlocked: () {
      setState(() => _unlocked = true);
    },);
  }
}

/// Verrou parental : « garde les deux lunes appuyées » — un appui long à deux
/// doigts, infranchissable à 4 ans, une seconde pour un adulte.
class _ParentLock extends StatefulWidget {
  const _ParentLock({required this.onUnlocked});

  final VoidCallback onUnlocked;

  @override
  State<_ParentLock> createState() => _ParentLockState();
}

class _ParentLockState extends State<_ParentLock> {
  int _pointersDown = 0;
  DateTime? _bothDownAt;

  void _check() {
    if (_pointersDown >= 2) {
      _bothDownAt ??= DateTime.now();
      Future.delayed(const Duration(seconds: 1), () {
        if (mounted &&
            _pointersDown >= 2 &&
            _bothDownAt != null &&
            DateTime.now().difference(_bothDownAt!) >=
                const Duration(milliseconds: 950)) {
          widget.onUnlocked();
        }
      });
    } else {
      _bothDownAt = null;
    }
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context).textTheme;
    return Scaffold(
      backgroundColor: AppColors.inkBlueHour,
      body: Listener(
        onPointerDown: (_) {
          setState(() => _pointersDown++);
          _check();
        },
        onPointerUp: (_) {
          setState(() => _pointersDown = (_pointersDown - 1).clamp(0, 10));
          _check();
        },
        child: SafeArea(
          child: Column(
            children: [
              Align(
                alignment: Alignment.centerLeft,
                child: BackButton(onPressed: context.pop),
              ),
              const Spacer(),
              Text('L\'espace des grands', style: theme.headlineMedium),
              const SizedBox(height: 12),
              Text(
                'Pose deux doigts sur les deux lunes\net garde-les appuyés.',
                style: theme.bodySmall,
                textAlign: TextAlign.center,
              ),
              const SizedBox(height: 48),
              Row(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  _moon(active: _pointersDown >= 1),
                  const SizedBox(width: 64),
                  _moon(active: _pointersDown >= 2),
                ],
              ),
              const Spacer(flex: 2),
            ],
          ),
        ),
      ),
    );
  }

  Widget _moon({required bool active}) => AnimatedContainer(
        duration: const Duration(milliseconds: 250),
        width: 88,
        height: 88,
        decoration: BoxDecoration(
          shape: BoxShape.circle,
          color: active ? AppColors.lanternGold : AppColors.nightCard,
          boxShadow:
              active ? AppColors.glow(AppColors.lanternGold, opacity: .5) : null,
        ),
        child: Icon(
          Icons.nightlight_round,
          size: 44,
          color: active ? AppColors.abyss : AppColors.mist,
        ),
      );
}

class _ParentDashboard extends ConsumerWidget {
  const _ParentDashboard();

  @override
  Widget build(BuildContext context, WidgetRef ref) {
    final profileAsync = ref.watch(childProfileProvider);
    final theme = ThemeData.light(useMaterial3: true).textTheme;
    final now = DateTime.now();
    final planner = ref.watch(episodeLengthPlannerProvider);

    return Scaffold(
      backgroundColor: AppColors.dawnLinen,
      appBar: AppBar(
        backgroundColor: AppColors.dawnLinen,
        leading: BackButton(onPressed: context.pop, color: AppColors.dayInk),
        title: Text('La Lune',
            style: theme.titleLarge?.copyWith(color: AppColors.dayInk),),
      ),
      body: profileAsync.when(
        loading: () => const Center(child: CircularProgressIndicator()),
        error: (_, __) => const SizedBox.shrink(),
        data: (profile) {
          if (profile == null) {
            return Center(
              child: FilledButton(
                onPressed: () => context.go(Routes.onboarding),
                child: const Text('Commencer la naissance du monde'),
              ),
            );
          }
          // Pacte de sommeil : coucher cible 20 h 30 (paramétrable à terme).
          final target = DateTime(now.year, now.month, now.day, 20, 30);
          final planned = planner.plan(
            ageBand: profile.ageBand(now),
            now: now,
            targetSleepTime: target.isAfter(now)
                ? target
                : target.add(const Duration(days: 1)),
          );
          return ListView(
            padding: const EdgeInsets.all(20),
            children: [
              _card(
                title: 'La météo du jour',
                subtitle:
                    '20 secondes pour que l\'épisode de ce soir parle vraiment '
                    'de la journée de ${profile.firstName}.',
                action: 'Faire le point',
                onTap: () => context.push(Routes.checkin),
                theme: theme,
              ),
              _card(
                title: 'Pacte de sommeil',
                subtitle:
                    'Coucher cible ${DateFormat.Hm().format(target)} — épisode '
                    'de ce soir calibré à ${planned.inMinutes} min.',
                theme: theme,
              ),
              _card(
                title: 'Ce que Plume sait de ${profile.firstName}',
                subtitle: 'Passions : ${profile.passions.join(', ')}\n'
                    'Peurs en cours de travail : '
                    '${profile.fears.isEmpty ? '—' : profile.fears.join(', ')}\n'
                    'Tout est éditable et effaçable, à tout moment.',
                theme: theme,
              ),
              _card(
                title: 'Le Livre de l\'Année',
                subtitle:
                    'Chaque épisode écrit une page. À la fin de l\'année : un vrai '
                    'livre relié de son aventure.',
                theme: theme,
              ),
            ],
          );
        },
      ),
    );
  }

  Widget _card({
    required String title,
    required String subtitle,
    required TextTheme theme,
    String? action,
    VoidCallback? onTap,
  }) =>
      Card(
        color: AppColors.dayCard,
        margin: const EdgeInsets.only(bottom: 14),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
        child: Padding(
          padding: const EdgeInsets.all(18),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(title,
                  style: theme.titleMedium?.copyWith(color: AppColors.dayInk),),
              const SizedBox(height: 6),
              Text(subtitle,
                  style: theme.bodyMedium
                      ?.copyWith(color: AppColors.dayInk.withOpacity(.7)),),
              if (action != null) ...[
                const SizedBox(height: 10),
                FilledButton.tonal(onPressed: onTap, child: Text(action)),
              ],
            ],
          ),
        ),
      );
}
