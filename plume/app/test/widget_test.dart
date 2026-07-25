import 'package:flutter_riverpod/flutter_riverpod.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:plume/app.dart';

void main() {
  testWidgets('l\'app démarre sur la Carte du Monde en mode démo',
      (tester) async {
    await tester.pumpWidget(const ProviderScope(child: PlumeApp()));
    // Les providers async (profil, compagnon, épisode) se résolvent.
    await tester.pump(const Duration(milliseconds: 300));
    await tester.pump(const Duration(milliseconds: 300));

    expect(find.text('Le monde de Léo'), findsOneWidget);
    expect(find.text('Le Foyer du Compagnon'), findsOneWidget);
    expect(find.text('La Grande Bibliothèque'), findsOneWidget);
    expect(find.text('C\'est l\'heure de l\'histoire'), findsOneWidget);
  });
}
