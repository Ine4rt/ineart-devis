import 'package:supabase_flutter/supabase_flutter.dart';

import '../../domain/entities/child_profile.dart';
import '../../domain/entities/companion.dart';
import '../../domain/entities/emotion_checkin.dart';
import '../../domain/entities/narrative_seed.dart';
import '../../domain/entities/story_episode.dart';
import '../../domain/repositories/world_repository.dart';

/// Implémentation production : Postgres (RLS par famille) + Edge Functions.
/// Les clés IA ne quittent jamais le serveur — l'app ne parle qu'à Supabase.
class SupabaseWorldRepository implements WorldRepository {
  SupabaseWorldRepository(this._client, {required this.childId});

  final SupabaseClient _client;
  final String childId;

  @override
  Future<ChildProfile?> loadProfile() async {
    final row = await _client
        .from('children')
        .select()
        .eq('id', childId)
        .maybeSingle();
    if (row == null) return null;
    return ChildProfile(
      id: row['id'] as String,
      firstName: row['first_name'] as String,
      birthDate: DateTime.parse(row['birth_date'] as String),
      passions: _stringList(row['passions']),
      fears: _stringList(row['fears']),
      pets: _stringList(row['pets']),
    );
  }

  @override
  Future<void> saveProfile(ChildProfile profile) => _client.from('children').upsert({
        'id': profile.id,
        'first_name': profile.firstName,
        'birth_date': profile.birthDate.toIso8601String(),
        'passions': profile.passions,
        'fears': profile.fears,
        'pets': profile.pets,
      });

  @override
  Future<Companion?> loadCompanion() async {
    final row = await _client
        .from('companions')
        .select()
        .eq('child_id', childId)
        .maybeSingle();
    if (row == null) return null;
    final dna = row['dna'] as Map<String, dynamic>;
    return Companion(
      id: row['id'] as String,
      name: row['name'] as String,
      dna: CompanionDna(
        species: dna['species'] as String,
        trait: dna['trait'] as String,
        pattern: dna['pattern'] as String,
        paletteSeed: dna['palette_seed'] as int,
        temperament: Temperament.values.byName(dna['temperament'] as String),
      ),
      bondLevel: (row['bond_level'] as int?) ?? 0,
      learnedWords: _stringList(row['learned_words']),
      stage: GrowthStage.forBond((row['bond_level'] as int?) ?? 0),
    );
  }

  @override
  Future<void> saveCompanion(Companion c) => _client.from('companions').upsert({
        'id': c.id,
        'child_id': childId,
        'name': c.name,
        'dna': {
          'species': c.dna.species,
          'trait': c.dna.trait,
          'pattern': c.dna.pattern,
          'palette_seed': c.dna.paletteSeed,
          'temperament': c.dna.temperament.name,
        },
        'bond_level': c.bondLevel,
        'learned_words': c.learnedWords,
      });

  @override
  Future<StoryEpisode?> tonightEpisode() async {
    final row = await _client
        .from('episodes')
        .select('*, scenes(*)')
        .eq('child_id', childId)
        .eq('status', 'ready')
        .order('number', ascending: false)
        .limit(1)
        .maybeSingle();
    if (row == null) return null;
    return _episodeFromRow(row);
  }

  @override
  Future<void> recordChoice(String episodeId, ChoiceOption option) =>
      _client.functions.invoke('record-choice', body: {
        'episode_id': episodeId,
        'option_id': option.id,
        'seed_summary': option.seedSummary,
      });

  @override
  Future<void> saveBookmark(String episodeId, SleepBookmark b) =>
      _client.from('episodes').update({
        'bookmark': {
          'scene_index': b.sceneIndex,
          'word_index': b.wordIndex,
          'fell_asleep_at': b.fellAsleepAt.toIso8601String(),
        },
      }).eq('id', episodeId);

  @override
  Future<void> submitCheckin(EmotionCheckin checkin) =>
      _client.from('emotion_checkins').upsert(checkin.toJson());

  @override
  Future<List<NarrativeSeed>> pendingSeeds() async {
    final rows = await _client
        .from('narrative_seeds')
        .select()
        .eq('child_id', childId)
        .isFilter('harvested_at', null)
        .order('planted_at');
    return [
      for (final row in rows as List)
        NarrativeSeed(
          id: row['id'] as String,
          kind: SeedKind.values.byName(row['kind'] as String),
          summary: row['summary'] as String,
          plantedAt: DateTime.parse(row['planted_at'] as String),
          germinateAfter: DateTime.parse(row['germinate_after'] as String),
        ),
    ];
  }

  @override
  Future<List<StoryEpisode>> archive({int limit = 30}) async {
    final rows = await _client
        .from('episodes')
        .select('*, scenes(*)')
        .eq('child_id', childId)
        .eq('status', 'told')
        .order('number', ascending: false)
        .limit(limit);
    return [for (final row in rows as List) _episodeFromRow(row as Map<String, dynamic>)];
  }

  StoryEpisode _episodeFromRow(Map<String, dynamic> row) {
    final sceneRows = (row['scenes'] as List? ?? [])
      ..sort((a, b) => (a['index'] as int).compareTo(b['index'] as int));
    return StoryEpisode(
      id: row['id'] as String,
      number: row['number'] as int,
      title: row['title'] as String,
      date: DateTime.parse(row['episode_date'] as String),
      emotionalThread: row['emotional_thread'] as String?,
      scenes: [
        for (final s in sceneRows)
          StoryScene(
            index: s['index'] as int,
            text: s['text'] as String,
            illustrationUrl: s['illustration_url'] as String?,
            audioUrl: s['audio_url'] as String?,
            choice: _choiceFromJson(s['choice'] as Map<String, dynamic>?),
          ),
      ],
    );
  }

  StoryChoice? _choiceFromJson(Map<String, dynamic>? json) {
    if (json == null) return null;
    return StoryChoice(
      prompt: json['prompt'] as String,
      options: [
        for (final o in json['options'] as List)
          ChoiceOption(
            id: o['id'] as String,
            label: o['label'] as String,
            seedSummary: o['seed_summary'] as String,
            iconUrl: o['icon_url'] as String?,
          ),
      ],
    );
  }

  List<String> _stringList(Object? value) =>
      [for (final v in (value as List? ?? [])) v as String];
}
