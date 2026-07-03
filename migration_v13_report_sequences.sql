-- ════════════════════════════════════════════════════════════════════════════
-- Migration V13 — Numérotation des rapports par province (saisie.html)
-- ════════════════════════════════════════════════════════════════════════════
-- À exécuter après migration_v10_supabase_auth.sql dans l'éditeur SQL Supabase.
-- Fournit un numéro séquentiel par province, utilisé par le formulaire de
-- saisie du rôle "province" (saisie.html) lors de la création d'un rapport.
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.report_sequences (
  province text primary key,
  last_seq int not null default 0
);

alter table public.report_sequences enable row level security;

drop policy if exists "report_sequences_auth" on public.report_sequences;
create policy "report_sequences_auth" on public.report_sequences for all to authenticated
  using (true) with check (true);

create or replace function public.next_report_seq(p_province text)
returns int language plpgsql security definer set search_path = public as $$
declare v_seq int;
begin
  insert into report_sequences(province, last_seq) values (p_province, 1)
  on conflict (province) do update
    set last_seq = report_sequences.last_seq + 1
  returning last_seq into v_seq;
  return v_seq;
end;
$$;

revoke execute on function public.next_report_seq(text) from anon, public;
grant  execute on function public.next_report_seq(text) to authenticated;
