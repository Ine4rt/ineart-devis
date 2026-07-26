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

/// Une scène du feuilleton. Aucune interaction : le récit se déroule seul,
/// comme une histoire qu'on écoute (décision produit — voir
/// docs/03-ARCHITECTURE.md §2, règle « aucune interaction »).
@immutable
class StoryScene {
  const StoryScene({
    required this.index,
    required this.text,
    this.illustrationUrl,
    this.audioUrl,
    this.sfx = 'silence',
  });

  final int index;
  final String text;
  final String? illustrationUrl;
  final String? audioUrl;

  /// Tag de bruitage choisi par le conteur (`feuilles`, `craquement`, `eau`,
  /// `nuit`, `silence`… voir docs/04-AUDIO.md). Le mixeur y associera le clip.
  final String sfx;

  /// Débit de narration du soir : [wordsPerMinute] mots/min.
  Duration get estimatedDuration {
    final words = text.split(RegExp(r'\s+')).length;
    return Duration(seconds: (words / wordsPerMinute * 60).round());
  }
}

/// Débit de référence de la narration du soir — unique dans tout le produit
/// (app, `generate-episode`, docs/04-AUDIO.md). Voix grave, rythme lent.
const int wordsPerMinute = 115;

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
