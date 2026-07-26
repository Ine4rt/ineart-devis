import 'package:flutter/material.dart';

/// Palette « L'Heure Bleue » — voir docs/02-DESIGN.md.
/// Le mode sombre est le mode principal : Plume est une app du soir.
abstract final class AppColors {
  // Nuit
  static const Color inkBlueHour = Color(0xFF0E1230); // fond principal
  static const Color abyss = Color(0xFF080B20); // scènes d'histoire
  static const Color nightCard = Color(0xFF181D42); // vitraux

  // Lueurs
  static const Color lanternGold = Color(0xFFFFC66E); // action, magie
  static const Color duskLavender = Color(0xFF9D8CFF); // compagnon
  static const Color dawnPeach = Color(0xFFFFA98F); // émotions
  static const Color fireflyGreen = Color(0xFF7FE3A8); // graines de lumière

  // Textes
  static const Color starWhite = Color(0xFFF4F1E8); // jamais de blanc pur
  static const Color mist = Color(0xFF9BA0C0); // texte secondaire

  // Jour (espace parent)
  static const Color dawnLinen = Color(0xFFFAF6EC);
  static const Color dayInk = Color(0xFF2A2E4A);
  static const Color dayCard = Color(0xFFFFFFFF);

  /// Halo standard : la lumière précède l'objet.
  static List<BoxShadow> glow(Color color, {double opacity = .35}) => [
        BoxShadow(
          color: color.withOpacity(opacity),
          blurRadius: 28,
          spreadRadius: 2,
        ),
      ];
}
