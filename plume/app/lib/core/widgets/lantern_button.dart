import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

import '../theme/app_colors.dart';

/// Le bouton principal de Plume : une lanterne qui respire.
/// Cible enfant ≥ 64 px, haptique légère, halo qui précède le geste.
class LanternButton extends StatefulWidget {
  const LanternButton({
    super.key,
    required this.label,
    required this.onPressed,
    this.icon,
    this.breathing = true,
  });

  final String label;
  final VoidCallback onPressed;
  final IconData? icon;

  /// Respiration lente du halo (coupée si reduce motion).
  final bool breathing;

  @override
  State<LanternButton> createState() => _LanternButtonState();
}

class _LanternButtonState extends State<LanternButton>
    with SingleTickerProviderStateMixin {
  late final AnimationController _breath = AnimationController(
    vsync: this,
    duration: const Duration(seconds: 4),
  )..repeat(reverse: true);

  bool _pressed = false;

  @override
  void dispose() {
    _breath.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final reduceMotion = MediaQuery.of(context).disableAnimations;
    final breathe = widget.breathing && !reduceMotion;

    return AnimatedBuilder(
      animation: _breath,
      builder: (context, child) {
        final glow = breathe ? .22 + .16 * _breath.value : .30;
        return AnimatedScale(
          scale: _pressed ? .96 : 1,
          duration: const Duration(milliseconds: 120),
          curve: Curves.easeOutCubic,
          child: Container(
            constraints: const BoxConstraints(minHeight: 64, minWidth: 200),
            decoration: BoxDecoration(
              borderRadius: BorderRadius.circular(32),
              gradient: const LinearGradient(
                colors: [AppColors.lanternGold, Color(0xFFFFB151)],
                begin: Alignment.topLeft,
                end: Alignment.bottomRight,
              ),
              boxShadow: AppColors.glow(AppColors.lanternGold, opacity: glow),
            ),
            child: child,
          ),
        );
      },
      child: Material(
        color: Colors.transparent,
        child: InkWell(
          borderRadius: BorderRadius.circular(32),
          onTapDown: (_) => setState(() => _pressed = true),
          onTapCancel: () => setState(() => _pressed = false),
          onTap: () {
            setState(() => _pressed = false);
            HapticFeedback.lightImpact();
            widget.onPressed();
          },
          child: Padding(
            padding: const EdgeInsets.symmetric(horizontal: 28, vertical: 18),
            child: Row(
              mainAxisSize: MainAxisSize.min,
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                if (widget.icon != null) ...[
                  Icon(widget.icon, color: AppColors.abyss, size: 26),
                  const SizedBox(width: 10),
                ],
                Text(
                  widget.label,
                  style: Theme.of(context).textTheme.labelLarge?.copyWith(
                        color: AppColors.abyss,
                        fontSize: 18,
                      ),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
