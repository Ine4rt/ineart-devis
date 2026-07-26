import 'package:flutter/material.dart';
import 'package:flutter_riverpod/flutter_riverpod.dart';

import 'core/router/app_router.dart';
import 'core/theme/app_theme.dart';

class PlumeApp extends ConsumerWidget {
  const PlumeApp({super.key});

  @override
  Widget build(BuildContext context, WidgetRef ref) {
    return MaterialApp.router(
      title: 'Plume',
      debugShowCheckedModeBanner: false,
      // La nuit est le mode principal : Plume est une app du soir.
      theme: AppTheme.dawn(),
      darkTheme: AppTheme.night(),
      themeMode: ThemeMode.dark,
      routerConfig: appRouter,
    );
  }
}
