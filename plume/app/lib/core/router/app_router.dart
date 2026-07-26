import 'package:flutter/material.dart';
import 'package:go_router/go_router.dart';

import '../../features/album/album_screen.dart';
import '../../features/checkin/checkin_screen.dart';
import '../../features/companion/companion_screen.dart';
import '../../features/onboarding/onboarding_screen.dart';
import '../../features/parent/parent_screen.dart';
import '../../features/ritual/ritual_screen.dart';
import '../../features/story/story_player_screen.dart';
import '../../features/worldmap/worldmap_screen.dart';

abstract final class Routes {
  static const onboarding = '/naissance';
  static const worldMap = '/';
  static const companion = '/compagnon';
  static const ritual = '/rituel';
  static const story = '/histoire';
  static const album = '/bibliotheque';
  static const checkin = '/lune/meteo';
  static const parent = '/lune';
}

final appRouter = GoRouter(
  initialLocation: Routes.worldMap,
  routes: [
    _lanternRoute(Routes.onboarding, const OnboardingScreen()),
    _lanternRoute(Routes.worldMap, const WorldMapScreen()),
    _lanternRoute(Routes.companion, const CompanionScreen()),
    _lanternRoute(Routes.ritual, const RitualScreen()),
    _lanternRoute(Routes.story, const StoryPlayerScreen()),
    _lanternRoute(Routes.album, const AlbumScreen()),
    _lanternRoute(Routes.parent, const ParentScreen()),
    _lanternRoute(Routes.checkin, const CheckinScreen()),
  ],
);

/// Transition signature « vol de lanterne » : fondu + très léger zoom,
/// la lumière arrive avant la scène. Jamais de push latéral (design §5.4).
GoRoute _lanternRoute(String path, Widget screen) => GoRoute(
      path: path,
      pageBuilder: (context, state) => CustomTransitionPage<void>(
        key: state.pageKey,
        transitionDuration: const Duration(milliseconds: 450),
        child: screen,
        transitionsBuilder: (context, animation, secondary, child) {
          final fade = CurvedAnimation(parent: animation, curve: Curves.easeOutCubic);
          final scale = Tween<double>(begin: .96, end: 1).animate(fade);
          return FadeTransition(
            opacity: fade,
            child: ScaleTransition(scale: scale, child: child),
          );
        },
      ),
    );
