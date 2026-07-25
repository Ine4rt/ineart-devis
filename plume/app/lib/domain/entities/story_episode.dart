import 'package:flutter/foundation.dart';

/// Un épisode canonique du monde de l'enfant. Pré-généré la nuit côté serveur
/// (texte + audio + illustrations) et téléchargé en avance : l'enfant n'attend
/// jamais (docs/03-ARCHITECTURE.md §1).
@immutable
class StoryEpisode {
  const StoryEpisode({
    required this.id,
    required this.number,
    required this.title,
    required this.scenes,
    required this.date,
    this.emotionalThread,
    this.bookmark,
    this.isDownloaded = false,
  });

  final String id;

  /// Numéro d'épisode dans la saga (l'enfant en est fier : « épisode 214 ! »).
  final int number;
  final String title;
  final List<StoryScene> scenes;
  final DateTime date;

  /// Le fil émotionnel du jour, tissé en filigrane (jamais montré à l'enfant).
  final String? emotionalThread;

  /// Marque-page d'endormissement : on reprend demain exactement ici (#19).
  final SleepBookmark? bookmark;
  final bool isDownloaded;

  Duration get estimatedDuration =>
      scenes.fold(Duration.zero, (sum, s) => sum + s.estimatedDuration);

  StoryEpisode withBookmark(SleepBookmark mark) => StoryEpisode(
        id: id,
        number: number,
        title: title,
        scenes: scenes,
        date: date,
        emotionalThread: emotionalThread,
        bookmark: mark,
        isDownloaded: isDownloaded,
      );
}

@immutable
class StoryScene {
  const StoryScene({
    required this.index,
    required this.text,
    this.illustrationUrl,
    this.audioUrl,
    this.choice,
  });

  final int index;
  final String text;
  final String? illustrationUrl;
  final String? audioUrl;

  /// Choix proposé à la fin de la scène, s'il y en a un.
  final StoryChoice? choice;

  /// ~140 mots/min en narration calme du soir.
  Duration get estimatedDuration {
    final words = text.split(RegExp(r'\s+')).length;
    return Duration(seconds: (words / 140 * 60).round());
  }
}

/// Un choix de l'enfant : 2-3 options illustrées, dicibles à voix haute.
/// Chaque option plante une graine narrative à longue portée.
@immutable
class StoryChoice {
  const StoryChoice({required this.prompt, required this.options});

  final String prompt;
  final List<ChoiceOption> options;
}

@immutable
class ChoiceOption {
  const ChoiceOption({
    required this.id,
    required this.label,
    required this.seedSummary,
    this.iconUrl,
  });

  final String id;
  final String label;

  /// Ce que ce choix plantera dans le monde (résumé canonique).
  final String seedSummary;
  final String? iconUrl;
}

/// « Endormi à 20 h 21, scène 4, mot 213. » Le monde attend l'enfant.
@immutable
class SleepBookmark {
  const SleepBookmark({
    required this.sceneIndex,
    required this.wordIndex,
    required this.fellAsleepAt,
  });

  final int sceneIndex;
  final int wordIndex;
  final DateTime fellAsleepAt;
}
