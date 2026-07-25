import 'package:flutter/foundation.dart';

/// La météo émotionnelle du jour (innovation #7) : 4 questions tactiles,
/// 20 secondes. L'IA tisse cela en filigrane — jamais frontalement.
@immutable
class EmotionCheckin {
  const EmotionCheckin({
    required this.childId,
    required this.date,
    required this.dayQuality,
    this.dominantEmotion,
    this.victory,
    this.fearOfTheDay,
  });

  final String childId;
  final DateTime date;
  final DayQuality dayQuality;
  final Emotion? dominantEmotion;

  /// « Il a réussi à faire du vélo sans les petites roues. »
  final String? victory;

  /// « Il a eu peur chez le dentiste. »
  final String? fearOfTheDay;

  Map<String, dynamic> toJson() => {
        'child_id': childId,
        'date': date.toIso8601String(),
        'day_quality': dayQuality.name,
        'dominant_emotion': dominantEmotion?.name,
        'victory': victory,
        'fear_of_the_day': fearOfTheDay,
      };
}

enum DayQuality { radiant, good, mixed, hard }

enum Emotion { joy, pride, calm, sadness, anger, fear, frustration }
