-- PLUME — schéma initial du canon narratif.
-- Chaque famille possède un monde ; RLS partout : une famille ne voit que le sien.

create extension if not exists "pgcrypto";

-- ─── Familles & enfants ──────────────────────────────────────────────────────

create table families (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users (id) on delete cascade,
  created_at timestamptz not null default now()
);

create table children (
  id uuid primary key default gen_random_uuid(),
  family_id uuid not null references families (id) on delete cascade,
  first_name text not null,
  birth_date date not null,
  passions text[] not null default '{}',
  fears text[] not null default '{}',
  pets text[] not null default '{}',
  reading_level text not null default 'listener',
  target_sleep_time time not null default '20:30',
  created_at timestamptz not null default now()
);

-- Proches castés en personnages (mamie → la Gardienne des Recettes Magiques).
create table family_cast (
  id uuid primary key default gen_random_uuid(),
  child_id uuid not null references children (id) on delete cascade,
  real_name text not null,
  role text not null,
  hero_name text not null,
  photo_path text
);

-- ─── Le compagnon ────────────────────────────────────────────────────────────

create table companions (
  id uuid primary key default gen_random_uuid(),
  child_id uuid not null unique references children (id) on delete cascade,
  name text not null,
  dna jsonb not null, -- species, trait, pattern, palette_seed, temperament
  bond_level int not null default 0,
  learned_words text[] not null default '{}',
  visual_ref_path text, -- fiche de référence pour la cohérence des illustrations
  created_at timestamptz not null default now()
);

-- ─── Le monde & son canon ────────────────────────────────────────────────────

create table worlds (
  id uuid primary key default gen_random_uuid(),
  child_id uuid not null unique references children (id) on delete cascade,
  name text not null,
  season text not null default 'printemps',
  style_bible jsonb not null default '{}', -- direction artistique des illustrations
  created_at timestamptz not null default now()
);

-- Entités canoniques : personnages, lieux, objets, légendes.
create table world_entities (
  id uuid primary key default gen_random_uuid(),
  world_id uuid not null references worlds (id) on delete cascade,
  kind text not null check (kind in ('character', 'place', 'object', 'legend')),
  name text not null,
  traits jsonb not null default '{}',
  state jsonb not null default '{}', -- état courant (le village grandit vraiment)
  first_seen_episode int,
  created_at timestamptz not null default now()
);

-- Journal append-only : la source de vérité du monde.
create table world_events (
  id uuid primary key default gen_random_uuid(),
  world_id uuid not null references worlds (id) on delete cascade,
  episode_number int,
  summary text not null,
  entities uuid[] not null default '{}',
  created_at timestamptz not null default now()
);

-- Les graines narratives : promesses à longue portée (#4, #7, #9, #11).
create table narrative_seeds (
  id uuid primary key default gen_random_uuid(),
  child_id uuid not null references children (id) on delete cascade,
  kind text not null check (kind in
    ('choiceConsequence', 'fearWork', 'lifeMoment', 'kindnessSeed', 'worldEvent')),
  summary text not null,
  priority text not null default 'normal' check (priority in ('low', 'normal', 'high')),
  planted_at timestamptz not null default now(),
  germinate_after timestamptz not null,
  harvested_at timestamptz,
  harvested_episode int
);

-- Le Fil d'Or : arcs pluriannuels planifiés comme une série (#3).
create table story_arcs (
  id uuid primary key default gen_random_uuid(),
  world_id uuid not null references worlds (id) on delete cascade,
  season_number int not null,
  act int not null,
  theme text not null, -- ex: « apprendre que demander de l'aide est un courage »
  planned_beats jsonb not null default '[]',
  status text not null default 'active' check (status in ('active', 'done'))
);

-- ─── Épisodes ────────────────────────────────────────────────────────────────

create table episodes (
  id uuid primary key default gen_random_uuid(),
  child_id uuid not null references children (id) on delete cascade,
  number int not null,
  title text not null,
  episode_date date not null default current_date,
  emotional_thread text, -- le fil du jour, jamais montré à l'enfant
  status text not null default 'generating'
    check (status in ('generating', 'ready', 'told', 'failed')),
  bookmark jsonb, -- { scene_index, word_index, fell_asleep_at } (#19)
  planned_minutes int not null default 9,
  unique (child_id, number)
);

create table scenes (
  id uuid primary key default gen_random_uuid(),
  episode_id uuid not null references episodes (id) on delete cascade,
  index int not null,
  text text not null,
  illustration_url text,
  audio_url text,
  choice jsonb, -- { prompt, options: [{ id, label, seed_summary }] }
  unique (episode_id, index)
);

-- ─── Météo émotionnelle & capsules ───────────────────────────────────────────

create table emotion_checkins (
  child_id uuid not null references children (id) on delete cascade,
  date date not null,
  day_quality text not null,
  dominant_emotion text,
  victory text,
  fear_of_the_day text,
  primary key (child_id, date)
);

create table time_capsules (
  id uuid primary key default gen_random_uuid(),
  child_id uuid not null references children (id) on delete cascade,
  audio_path text not null,
  sealed_at timestamptz not null default now(),
  opens_at timestamptz not null
);

-- ─── RLS : une famille ne voit que son monde ─────────────────────────────────

alter table families enable row level security;
alter table children enable row level security;
alter table family_cast enable row level security;
alter table companions enable row level security;
alter table worlds enable row level security;
alter table world_entities enable row level security;
alter table world_events enable row level security;
alter table narrative_seeds enable row level security;
alter table story_arcs enable row level security;
alter table episodes enable row level security;
alter table scenes enable row level security;
alter table emotion_checkins enable row level security;
alter table time_capsules enable row level security;

create policy "own family" on families
  for all using (owner_id = auth.uid());

create policy "own children" on children
  for all using (family_id in (select id from families where owner_id = auth.uid()));

-- Helper : les enfants de la famille de l'utilisateur courant.
create or replace function my_child_ids()
returns setof uuid language sql stable security definer set search_path = public as $$
  select c.id from children c
  join families f on f.id = c.family_id
  where f.owner_id = auth.uid()
$$;

create policy "own cast" on family_cast for all using (child_id in (select my_child_ids()));
create policy "own companion" on companions for all using (child_id in (select my_child_ids()));
create policy "own world" on worlds for all using (child_id in (select my_child_ids()));
create policy "own entities" on world_entities for all
  using (world_id in (select w.id from worlds w where w.child_id in (select my_child_ids())));
create policy "own events" on world_events for all
  using (world_id in (select w.id from worlds w where w.child_id in (select my_child_ids())));
create policy "own seeds" on narrative_seeds for all using (child_id in (select my_child_ids()));
create policy "own arcs" on story_arcs for all
  using (world_id in (select w.id from worlds w where w.child_id in (select my_child_ids())));
create policy "own episodes" on episodes for all using (child_id in (select my_child_ids()));
create policy "own scenes" on scenes for all
  using (episode_id in (select e.id from episodes e where e.child_id in (select my_child_ids())));
create policy "own checkins" on emotion_checkins for all using (child_id in (select my_child_ids()));
create policy "own capsules" on time_capsules for all using (child_id in (select my_child_ids()));

-- ─── Index ───────────────────────────────────────────────────────────────────

create index idx_seeds_harvest on narrative_seeds (child_id, germinate_after)
  where harvested_at is null;
create index idx_episodes_child on episodes (child_id, number desc);
create index idx_events_world on world_events (world_id, created_at desc);
