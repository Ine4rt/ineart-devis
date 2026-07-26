import 'package:flutter/material.dart';
import 'package:google_fonts/google_fonts.dart';

import 'app_colors.dart';

/// Fraunces = la voix du conte (titres). Nunito Sans = l'interface.
abstract final class AppTypography {
  static TextTheme textTheme(Color body, Color soft) {
    const serif = GoogleFonts.fraunces;
    const sans = GoogleFonts.nunitoSans;
    return TextTheme(
      displayLarge: serif(fontSize: 40, fontWeight: FontWeight.w600, color: body, height: 1.15),
      headlineLarge: serif(fontSize: 32, fontWeight: FontWeight.w600, color: body, height: 1.2),
      headlineMedium: serif(fontSize: 24, fontWeight: FontWeight.w600, color: body),
      titleLarge: sans(fontSize: 20, fontWeight: FontWeight.w700, color: body),
      titleMedium: sans(fontSize: 17, fontWeight: FontWeight.w700, color: body),
      bodyLarge: sans(fontSize: 17, color: body, height: 1.45),
      bodyMedium: sans(fontSize: 16, color: body, height: 1.4),
      bodySmall: sans(fontSize: 14, color: soft),
      labelLarge: sans(fontSize: 16, fontWeight: FontWeight.w800, color: body, letterSpacing: .3),
    );
  }

  /// Texte lu par l'enfant : jamais sous 20 pt, interlettrage augmenté.
  static TextStyle storyText = GoogleFonts.fraunces(
    fontSize: 22,
    height: 1.6,
    letterSpacing: .4,
    color: AppColors.starWhite,
  );
}
