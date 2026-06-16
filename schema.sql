-- ── Проекты ──────────────────────────────────────────────────────────────────
create table if not exists projects (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  description text,
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

-- ── Запуски анализа ───────────────────────────────────────────────────────────
create table if not exists analyses (
  id            uuid primary key default gen_random_uuid(),
  project_id    uuid not null references projects(id) on delete cascade,
  filename      text not null,
  status        text not null default 'pending',   -- pending | completed | error
  error_message text,
  stats_total   int  default 0,
  stats_auto    int  default 0,
  stats_request int  default 0,
  created_at    timestamptz default now()
);

-- ── Файлы ─────────────────────────────────────────────────────────────────────
create table if not exists files (
  id           uuid primary key default gen_random_uuid(),
  analysis_id  uuid not null references analyses(id) on delete cascade,
  file_type    text not null,   -- source | analysis | request
  storage_path text not null,
  created_at   timestamptz default now()
);

-- ── Индексы ───────────────────────────────────────────────────────────────────
create index if not exists analyses_project_id_idx on analyses(project_id);
create index if not exists files_analysis_id_idx   on files(analysis_id);

-- ── Auto-update updated_at ────────────────────────────────────────────────────
create or replace function update_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists projects_updated_at on projects;
create trigger projects_updated_at
  before update on projects
  for each row execute function update_updated_at();

-- ── Storage bucket ────────────────────────────────────────────────────────────
insert into storage.buckets (id, name, public)
values ('analysis-files', 'analysis-files', false)
on conflict (id) do nothing;
