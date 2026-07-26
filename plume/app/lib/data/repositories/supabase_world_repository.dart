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
  SupabaseWorldRepository(this._client, {required String childId})
      : _childId = childId;

  final SupabaseClient _client;

  /// Enfant courant. L'onboarding crée un nouvel enfant : [saveProfile] le
  /// mémorise pour que compagnon, monde et épisodes soient bien rattachés à
  /// l'enfant qu'on vient de créer, et pas à celui d'avant.
  String _childId;

  String get childId => _childId;

  @override
  Future<ChildProfile?> loadProfile() async {
    final row = await _client
        .from('children')
        .select()
        .eq('id', _childId)
        .maybeSingle();
    if (row == null) return null;
    final sleep = row['target_sleep_time'] as String?;
    return ChildProfile(
      id: row['id'] as String,
      firstName: row['first_name'] as String,
      birthDate: DateTime.parse(row['birth_date'] as String),
      passions: _stringList(row['passions']),
      fears: _stringList(row['fears']),
      pets: _stringList(row['pets']),
      storyMinutes: row['story_minutes'] as int?,
      targetSleepTime:
          sleep == null ? const SleepTime(20, 30) : SleepTime.parse(sleep),
    );
  }

  @override
  Future<void> saveProfile(ChildProfile profile) async {
    // `children.family_id` est NOT NULL : sans famille, tout l'onboarding
    // échouait silencieusement. On crée (ou retrouve) celle de l'utilisateur.
    final familyId = await _ensureFamily();

    await _client.from('children').upsert({
      'id': profile.id,
      'family_id': familyId,
      'first_name': profile.firstName,
      'birth_date': profile.birthDate.toIso8601String(),
      'passions': profile.passions,
      'fears': profile.fears,
      'pets': profile.pets,
      'reading_level': profile.readingLevel.name,
      if (profile.storyMinutes != null) 'story_minutes': profile.storyMinutes,
      'target_sleep_time': profile.targetSleepTime.toSql(),
    });

    // À partir d'ici, c'est cet enfant-là qui est le monde courant.
    _childId = profile.id;

    // `generate-episode` s'appuie sur `worlds.id` : sans ligne `worlds`,
    // la requête canon partait sur un uuid vide et échouait sans bruit.
    await _ensureWorld(profile);
  }

  /// La famille de l'utilisateur connecté, créée à la première sauvegarde.
  Future<String> _ensureFamily() async {
    final userId = _client.auth.currentUser?.id;
    if (userId == null) {
      throw StateError('Aucune session Supabase : impossible de créer la famille.');
    }
    final existing = await _client
        .from('families')
        .select('id')
        .eq('owner_id', userId)
        .limit(1)
        .maybeSingle();
    if (existing != null) return existing['id'] as String;

    final created = await _client
        .from('families')
        .insert({'owner_id': userId})
        .select('id')
        .single();
    return created['id'] as String;
  }

  /// Le monde de l'enfant — créé une fois, à la naissance du monde.
  Future<void> _ensureWorld(ChildProfile profile) async {
    final existing = await _client
        .from('worlds')
        .select('id')
        .eq('child_id', profile.id)
        .maybeSingle();
    if (existing != null) return;

    await _client.from('worlds').insert({
      'child_id': profile.id,
      'name': 'Le monde de ${profile.firstName}',
      // Registre par défaut : monde réel, poésie du quotidien.
      'style_bible': {'register': 'réalisme doux'},
    });
  }

  @override
  Future<Companion?> loadCompanion() async {
    final row = await _client
        .from('companions')
        .select()
        .eq('child_id', _childId)
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
        // `_childId` est mis à jour par saveProfile : le compagnon est bien
        // rattaché à l'enfant qui vient de naître, pas à un enfant précédent.
        'child_id': _childId,
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
        .eq('child_id', _childId)
        .eq('status', 'ready')
        .order('number', ascending: false)
        .limit(1)
        .maybeSingle();
    if (row == null) return null;
    return _episodeFromRow(row);
  }

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
        .eq('child_id', _childId)
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
        .eq('child_id', _childId)
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
            // Tag de bruitage posé par le conteur (docs/04-AUDIO.md).
            sfx: (s['sfx'] as String?) ?? 'silence',
          ),
      ],
    );
  }

  List<String> _stringList(Object? value) =>
      [for (final v in (value as List? ?? [])) v as String];
}
