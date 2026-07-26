import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_animate/flutter_animate.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:go_router/go_router.dart';
import 'package:uuid/uuid.dart';

import '../../core/router/app_router.dart';
import '../../core/theme/app_colors.dart';
import '../../core/widgets/lantern_button.dart';
import '../../core/widgets/starfield_background.dart';
import '../../domain/entities/child_profile.dart';
import '../../domain/entities/companion.dart';
import '../../domain/usecases/companion_genesis.dart';
import '../../providers.dart';

/// Onboarding « La naissance d'un monde » (design §4.2) : pas un formulaire,
/// une cérémonie. Version implémentée : prénom → âge → passions → l'Œuf →
/// révélation du monde. (Peurs & famille : ajoutées ensuite côté parent.)
class OnboardingScreen extends ConsumerStatefulWidget {
  const OnboardingScreen({super.key});

  @override
  ConsumerState<OnboardingScreen> createState() => _OnboardingScreenState();
}

class _OnboardingScreenState extends ConsumerState<OnboardingScreen> {
  final _pageController = PageController();
  final _nameController = TextEditingController();

  int _page = 0;
  int _age = 5;
  final Set<String> _passions = {};
  int _eggWarmth = 0; // caresses sur l'œuf
  CompanionDna? _bornDna;

  static const _passionChoices = [
    'espace', 'dinosaures', 'animaux', 'danse', 'foot',
    'dessin', 'blagues', 'nature', 'livres', 'musique',
  ];

  @override
  void dispose() {
    _pageController.dispose();
    _nameController.dispose();
    super.dispose();
  }

  void _next() {
    HapticFeedback.lightImpact();
    _pageController.nextPage(
      duration: const Duration(milliseconds: 600),
      curve: Curves.easeOutCubic,
    );
  }

  Future<void> _warmEgg() async {
    HapticFeedback.lightImpact();
    setState(() => _eggWarmth++);
    if (_eggWarmth >= 5 && _bornDna == null) {
      // L'éclosion : le moment signature de Plume.
      HapticFeedback.mediumImpact();
      final childId = const Uuid().v4();
      final dna = const CompanionGenesis().generate(
        childId: childId,
        salt: DateTime.now().microsecondsSinceEpoch,
        passions: _passions.toList(),
      );
      setState(() => _bornDna = dna);

      final repo = ref.read(worldRepositoryProvider);
      // On ne demande pas la date de naissance exacte à un enfant de 4 ans.
      // On la dérive du jour courant plutôt que du 1er janvier : l'âge est
      // alors juste aujourd'hui, au lieu de dériver d'un semestre.
      // (Le parent peut l'affiner ensuite depuis la Lune.)
      final now = DateTime.now();
      await repo.saveProfile(
        ChildProfile(
          id: childId,
          firstName: _nameController.text.trim(),
          birthDate: DateTime(now.year - _age, now.month, now.day),
          passions: _passions.toList(),
        ),
      );
      await repo.saveCompanion(
        Companion(id: const Uuid().v4(), name: 'Pipo', dna: dna),
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context).textTheme;
    return Scaffold(
      body: StarfieldBackground(
        // Le monde s'anime à mesure que l'enfant répond.
        fireflyCount: 2 + _page * 2 + _passions.length,
        child: SafeArea(
          child: PageView(
            controller: _pageController,
            physics: const NeverScrollableScrollPhysics(),
            onPageChanged: (page) => setState(() => _page = page),
            children: [
              _Scene(
                title: 'Comment t\'appelles-tu ?',
                subtitle: 'Ton prénom va s\'écrire dans les étoiles.',
                child: Column(
                  children: [
                    TextField(
                      controller: _nameController,
                      textAlign: TextAlign.center,
                      style: theme.displayLarge,
                      cursorColor: AppColors.lanternGold,
                      decoration: const InputDecoration(border: InputBorder.none),
                    ),
                    const SizedBox(height: 32),
                    LanternButton(
                      label: 'C\'est moi !',
                      onPressed: () {
                        if (_nameController.text.trim().isNotEmpty) _next();
                      },
                    ),
                  ],
                ),
              ),
              _Scene(
                title: 'Quel âge as-tu ?',
                subtitle: 'Un anneau de ton arbre pour chaque année.',
                child: Column(
                  children: [
                    Text('$_age ans', style: theme.displayLarge)
                        .animate(key: ValueKey(_age))
                        .scale(begin: const Offset(.8, .8), curve: Curves.easeOutBack),
                    Slider(
                      value: _age.toDouble(),
                      min: 2,
                      max: 11,
                      divisions: 9,
                      activeColor: AppColors.lanternGold,
                      onChanged: (v) => setState(() => _age = v.round()),
                    ),
                    const SizedBox(height: 16),
                    LanternButton(label: 'Continuer', onPressed: _next),
                  ],
                ),
              ),
              _Scene(
                title: 'Qu\'est-ce que tu aimes ?',
                subtitle: 'Chaque passion fera naître une île de ton monde.',
                child: Column(
                  children: [
                    Wrap(
                      spacing: 10,
                      runSpacing: 10,
                      alignment: WrapAlignment.center,
                      children: [
                        for (final passion in _passionChoices)
                          FilterChip(
                            label: Text(passion),
                            selected: _passions.contains(passion),
                            selectedColor: AppColors.lanternGold.withOpacity(.3),
                            checkmarkColor: AppColors.lanternGold,
                            onSelected: (selected) => setState(() {
                              selected
                                  ? _passions.add(passion)
                                  : _passions.remove(passion);
                              HapticFeedback.selectionClick();
                            }),
                          ),
                      ],
                    ),
                    const SizedBox(height: 32),
                    LanternButton(
                      label: 'Continuer',
                      onPressed: () {
                        if (_passions.isNotEmpty) _next();
                      },
                    ),
                  ],
                ),
              ),
              _Scene(
                title: _bornDna == null ? 'Un œuf t\'attend…' : 'Il est né !',
                subtitle: _bornDna == null
                    ? 'Caresse-le pour le réchauffer.'
                    : 'Un ${_bornDna!.species} aux ${_bornDna!.trait}.\n'
                        'Il n\'en existe aucun autre comme lui au monde.',
                child: Column(
                  children: [
                    GestureDetector(
                      onTap: _bornDna == null ? _warmEgg : null,
                      child: Text(
                        _bornDna == null ? '🥚' : '🦊',
                        style: const TextStyle(fontSize: 110),
                      )
                          .animate(key: ValueKey('$_eggWarmth-${_bornDna != null}'))
                          .shake(
                            hz: 3,
                            offset: Offset(2.0 * _eggWarmth, 0),
                            duration: 400.ms,
                          )
                          .scale(
                            begin: const Offset(.95, .95),
                            end: const Offset(1, 1),
                            curve: Curves.easeOutBack,
                          ),
                    ),
                    const SizedBox(height: 32),
                    if (_bornDna != null)
                      LanternButton(
                        label: 'Découvrir mon monde',
                        icon: Icons.public_rounded,
                        onPressed: () => context.go(Routes.worldMap),
                      ).animate().fadeIn(delay: 800.ms),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class _Scene extends StatelessWidget {
  const _Scene({required this.title, required this.subtitle, required this.child});

  final String title;
  final String subtitle;
  final Widget child;

  @override
  Widget build(BuildContext context) {
    final theme = Theme.of(context).textTheme;
    return Padding(
      padding: const EdgeInsets.all(32),
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Text(title, style: theme.headlineLarge, textAlign: TextAlign.center)
              .animate()
              .fadeIn(duration: 600.ms)
              .slideY(begin: .1, curve: Curves.easeOutCubic),
          const SizedBox(height: 8),
          Text(subtitle, style: theme.bodySmall, textAlign: TextAlign.center)
              .animate()
              .fadeIn(delay: 200.ms),
          const SizedBox(height: 40),
          child,
        ],
      ),
    );
  }
}
