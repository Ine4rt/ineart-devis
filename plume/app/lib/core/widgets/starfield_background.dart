import 'dart:math';

import 'package:flutter/material.dart';

import '../theme/app_colors.dart';

/// Ciel de l'Heure Bleue : dégradé vivant + étoiles qui scintillent +
/// lucioles dérivantes. Un seul CustomPainter pour tout (60 fps — jamais
/// un widget par particule, cf. docs/02-DESIGN.md §5 règle 3).
/// Respecte `MediaQuery.disableAnimations` (reduce motion).
class StarfieldBackground extends StatefulWidget {
  const StarfieldBackground({
    super.key,
    this.starCount = 90,
    this.fireflyCount = 7,
    this.child,
  });

  final int starCount;
  final int fireflyCount;
  final Widget? child;

  @override
  State<StarfieldBackground> createState() => _StarfieldBackgroundState();
}

class _StarfieldBackgroundState extends State<StarfieldBackground>
    with SingleTickerProviderStateMixin {
  late final AnimationController _controller = AnimationController(
    vsync: this,
    duration: const Duration(seconds: 40),
  )..repeat();

  @override
  void dispose() {
    _controller.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final reduceMotion = MediaQuery.of(context).disableAnimations;
    return Stack(
      fit: StackFit.expand,
      children: [
        const DecoratedBox(
          decoration: BoxDecoration(
            gradient: LinearGradient(
              begin: Alignment.topCenter,
              end: Alignment.bottomCenter,
              colors: [AppColors.abyss, AppColors.inkBlueHour],
            ),
          ),
        ),
        if (reduceMotion)
          CustomPaint(painter: _SkyPainter(0, widget.starCount, 0))
        else
          AnimatedBuilder(
            animation: _controller,
            builder: (_, __) => CustomPaint(
              painter: _SkyPainter(
                _controller.value,
                widget.starCount,
                widget.fireflyCount,
              ),
            ),
          ),
        if (widget.child != null) widget.child!,
      ],
    );
  }
}

class _SkyPainter extends CustomPainter {
  _SkyPainter(this.t, this.starCount, this.fireflyCount);

  final double t; // 0..1, boucle de 40 s
  final int starCount;
  final int fireflyCount;

  // Graine fixe : le ciel est stable d'une frame à l'autre.
  static final _rng = Random(1211);
  static final List<_Star> _stars = List.generate(400, (_) {
    return _Star(
      dx: _rng.nextDouble(),
      dy: _rng.nextDouble(),
      size: .5 + _rng.nextDouble() * 1.6,
      phase: _rng.nextDouble() * 2 * pi,
    );
  });

  @override
  void paint(Canvas canvas, Size size) {
    final starPaint = Paint();
    for (var i = 0; i < starCount && i < _stars.length; i++) {
      final s = _stars[i];
      // Scintillement doux, désynchronisé par étoile.
      final twinkle = .55 + .45 * sin(t * 2 * pi * 6 + s.phase);
      starPaint.color = AppColors.starWhite.withOpacity(.35 * twinkle);
      canvas.drawCircle(
        Offset(s.dx * size.width, s.dy * size.height * .8),
        s.size,
        starPaint,
      );
    }

    // Lucioles : trajectoire de Lissajous lente, halo doré.
    for (var i = 0; i < fireflyCount; i++) {
      final phase = i * 1.7;
      final x = size.width * (.5 + .42 * sin(t * 2 * pi + phase));
      final y = size.height * (.45 + .35 * sin(t * 4 * pi + phase * 1.3));
      final pulse = .5 + .5 * sin(t * 2 * pi * 10 + phase);
      final glow = Paint()
        ..color = AppColors.lanternGold.withOpacity(.10 + .12 * pulse)
        ..maskFilter = const MaskFilter.blur(BlurStyle.normal, 12);
      canvas.drawCircle(Offset(x, y), 9, glow);
      canvas.drawCircle(
        Offset(x, y),
        2.2,
        Paint()..color = AppColors.lanternGold.withOpacity(.6 + .4 * pulse),
      );
    }
  }

  @override
  bool shouldRepaint(_SkyPainter old) => old.t != t;
}

class _Star {
  const _Star({required this.dx, required this.dy, required this.size, required this.phase});
  final double dx, dy, size, phase;
}
