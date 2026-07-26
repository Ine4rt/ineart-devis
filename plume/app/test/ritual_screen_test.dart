import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:go_router/go_router.dart';
import 'package:plume/features/ritual/ritual_screen.dart';

void main() {
  // Régression : `AnimationController.repeat(reverse: true)` n'émet jamais
  // `completed`/`dismissed` (il alterne forward/reverse en interne), donc le
  // compteur de respirations restait bloqué à zéro et l'enfant ne quittait
  // jamais cet écran. Ce test verrouille les 3 respirations puis la
  // navigation vers l'histoire.
  testWidgets('les 3 respirations s\'accomplissent puis on rejoint l\'histoire',
      (tester) async {
    var reachedStory = false;
    final router = GoRouter(
      initialLocation: '/rituel',
      routes: [
        GoRoute(path: '/rituel', builder: (_, __) => const RitualScreen()),
        GoRoute(
          path: '/histoire',
          builder: (_, __) {
            reachedStory = true;
            return const SizedBox.shrink();
          },
        ),
      ],
    );

    await tester.pumpWidget(
      ProviderScope(child: MaterialApp.router(routerConfig: router)),
    );
    await tester.pump();

    // Une respiration = 8 s (4 s d'inspiration + 4 s d'expiration).
    for (var i = 0; i < 3; i++) {
      await tester.pump(const Duration(seconds: 4));
      await tester.pump(const Duration(seconds: 4));
    }
    await tester.pumpAndSettle();

    expect(reachedStory, isTrue);
  });

  testWidgets('« Passer » rejoint l\'histoire immédiatement', (tester) async {
    var reachedStory = false;
    final router = GoRouter(
      initialLocation: '/rituel',
      routes: [
        GoRoute(path: '/rituel', builder: (_, __) => const RitualScreen()),
        GoRoute(
          path: '/histoire',
          builder: (_, __) {
            reachedStory = true;
            return const SizedBox.shrink();
          },
        ),
      ],
    );

    await tester.pumpWidget(
      ProviderScope(child: MaterialApp.router(routerConfig: router)),
    );
    await tester.pump();

    await tester.tap(find.text('Passer'));
    await tester.pumpAndSettle();

    expect(reachedStory, isTrue);
  });
}
