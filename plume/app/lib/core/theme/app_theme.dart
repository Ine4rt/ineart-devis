import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

import 'app_colors.dart';
import 'app_typography.dart';

abstract final class AppTheme {
  /// Mode principal : la nuit de l'Heure Bleue.
  static ThemeData night() {
    final scheme = ColorScheme.fromSeed(
      seedColor: AppColors.lanternGold,
      brightness: Brightness.dark,
      surface: AppColors.inkBlueHour,
      primary: AppColors.lanternGold,
      secondary: AppColors.duskLavender,
      tertiary: AppColors.dawnPeach,
      onSurface: AppColors.starWhite,
    );
    return _base(scheme).copyWith(
      scaffoldBackgroundColor: AppColors.inkBlueHour,
      cardColor: AppColors.nightCard,
      textTheme: AppTypography.textTheme(AppColors.starWhite, AppColors.mist),
    );
  }

  /// Mode jour : l'espace parent, le matin.
  static ThemeData dawn() {
    final scheme = ColorScheme.fromSeed(
      seedColor: AppColors.lanternGold,
      brightness: Brightness.light,
      surface: AppColors.dawnLinen,
      primary: AppColors.dayInk,
      secondary: AppColors.duskLavender,
      onSurface: AppColors.dayInk,
    );
    return _base(scheme).copyWith(
      scaffoldBackgroundColor: AppColors.dawnLinen,
      cardColor: AppColors.dayCard,
      textTheme: AppTypography.textTheme(AppColors.dayInk, AppColors.dayInk.withOpacity(.6)),
    );
  }

  static ThemeData _base(ColorScheme scheme) => ThemeData(
        useMaterial3: true,
        colorScheme: scheme,
        splashFactory: InkSparkle.splashFactory,
        visualDensity: VisualDensity.standard,
        appBarTheme: const AppBarTheme(
          backgroundColor: Colors.transparent,
          elevation: 0,
          systemOverlayStyle: SystemUiOverlayStyle.light,
        ),
        pageTransitionsTheme: const PageTransitionsTheme(
          builders: {
            TargetPlatform.iOS: FadeUpwardsPageTransitionsBuilder(),
            TargetPlatform.android: FadeUpwardsPageTransitionsBuilder(),
          },
        ),
      );
}
