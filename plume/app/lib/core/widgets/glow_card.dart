import 'dart:ui';

import 'package:flutter/material.dart';

import '../theme/app_colors.dart';

/// Carte « vitrail » : verre dépoli sombre, bord lumineux 1 px, halo émis.
/// Dans la nuit, ce qui compte émet de la lumière — pas d'ombres portées.
class GlowCard extends StatelessWidget {
  const GlowCard({
    super.key,
    required this.child,
    this.glowColor = AppColors.lanternGold,
    this.onTap,
    this.padding = const EdgeInsets.all(20),
  });

  final Widget child;
  final Color glowColor;
  final VoidCallback? onTap;
  final EdgeInsets padding;

  @override
  Widget build(BuildContext context) {
    return DecoratedBox(
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(24),
        boxShadow: AppColors.glow(glowColor, opacity: .18),
      ),
      child: ClipRRect(
        borderRadius: BorderRadius.circular(24),
        child: BackdropFilter(
          filter: ImageFilter.blur(sigmaX: 14, sigmaY: 14),
          child: Material(
            color: AppColors.nightCard.withOpacity(.72),
            child: InkWell(
              onTap: onTap,
              child: Container(
                padding: padding,
                decoration: BoxDecoration(
                  borderRadius: BorderRadius.circular(24),
                  border: Border.all(color: glowColor.withOpacity(.20)),
                ),
                child: child,
              ),
            ),
          ),
        ),
      ),
    );
  }
}
