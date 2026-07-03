-- ════════════════════════════════════════════════════════════════════════════
-- Migration V10 — Authentification Supabase Auth + verrouillage des RLS
-- ════════════════════════════════════════════════════════════════════════════
-- À exécuter après migration_v9_province_comptables.sql dans l'éditeur SQL
-- Supabase (SQL Editor).
--
-- Contexte : jusqu'à cette migration, l'application n'avait AUCUN écran de
-- connexion fonctionnel (aucune page ne créait la session locale attendue) et
-- la table `users` était lisible par n'importe qui via la clé anon publique,
-- exposant emails + mots de passe (encodage base64 réversible) de tous les
-- comptes. Cette migration :
--   1. Crée un vrai compte Supabase Auth pour chaque ligne de `users`, en
--      réutilisant le mot de passe existant (décodé depuis password_b64) ;
--   2. Crée la table `profiles`, source de vérité pour rôle/provinces,
--      liée à auth.users(id) ;
--   3. Ajoute des fonctions SECURITY DEFINER pour évaluer les RLS sans
--      récursion (is_super_admin, has_province_access, ...) ;
--   4. Retire tout accès anonyme (rôle `anon`) sur les tables métier et migre
--      les policies vers `authenticated`, avec scoping par province/rôle ;
--   5. Verrouille l'ancienne table `users` (conservée pour historique,
--      réservée au super_admin, n'est plus utilisée pour l'authentification).
--
-- La création des comptes (étape 1) et le déploiement de la fonction Edge
-- `admin-users` (utilisée par superadmin.html pour créer/éditer/supprimer des
-- comptes avec les privilèges service_role) ont été réalisés directement sur
-- le projet Supabase de production. Ce script reproduit fidèlement ces
-- opérations pour un déploiement complet à partir de zéro (voir README.md).
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1. Comptes Supabase Auth à partir de la table users existante ────────────
create extension if not exists pgcrypto;

-- Corriger l'email invalide "se_uncp" (sans domaine) s'il existe encore
update public.users set email = 'se_uncp@pnda.cd' where email = 'se_uncp';

insert into auth.users (
  instance_id, id, aud, role, email, encrypted_password,
  email_confirmed_at, raw_app_meta_data, raw_user_meta_data,
  created_at, updated_at, confirmation_token, recovery_token,
  email_change_token_new, email_change, is_sso_user, is_anonymous
)
select
  '00000000-0000-0000-0000-000000000000',
  u.id,
  'authenticated',
  'authenticated',
  lower(u.email),
  crypt(convert_from(decode(u.password_b64, 'base64'), 'UTF8'), gen_salt('bf')),
  now(),
  jsonb_build_object('provider','email','providers', jsonb_build_array('email')),
  '{}'::jsonb,
  now(), now(), '', '', '', '', false, false
from public.users u
where not exists (select 1 from auth.users au where au.id = u.id);

insert into auth.identities (provider_id, user_id, identity_data, provider, last_sign_in_at, created_at, updated_at)
select
  u.id::text, u.id,
  jsonb_build_object('sub', u.id::text, 'email', lower(u.email)),
  'email', now(), now(), now()
from public.users u
where not exists (select 1 from auth.identities ai where ai.user_id = u.id and ai.provider = 'email');

-- ── 2. Table profiles (rôle/provinces liés à auth.users) ─────────────────────
create table if not exists public.profiles (
  id           uuid primary key references auth.users(id) on delete cascade,
  email        text unique not null,
  display_name text,
  role         text not null default 'province' check (role in ('province','comptable','admin','super_admin')),
  provinces    text[],              -- NULL = accès à toutes les provinces
  is_active    boolean not null default true,
  created_at   timestamptz not null default now(),
  updated_at   timestamptz not null default now()
);

insert into public.profiles (id, email, display_name, role, provinces, is_active)
select u.id, lower(u.email), u.display_name, u.role, u.provinces, true
from public.users u
where not exists (select 1 from public.profiles p where p.id = u.id);

alter table public.profiles enable row level security;

-- ── 3. Fonctions utilitaires SECURITY DEFINER (évitent la récursion RLS) ─────
create or replace function public.current_profile_role()
returns text language sql security definer set search_path = public stable as $$
  select role from public.profiles where id = auth.uid()
$$;

create or replace function public.current_profile_provinces()
returns text[] language sql security definer set search_path = public stable as $$
  select provinces from public.profiles where id = auth.uid()
$$;

create or replace function public.is_super_admin()
returns boolean language sql security definer set search_path = public stable as $$
  select exists(select 1 from public.profiles where id = auth.uid() and role = 'super_admin')
$$;

create or replace function public.has_province_access(p_province text)
returns boolean language sql security definer set search_path = public stable as $$
  select
    coalesce(public.is_super_admin(), false)
    or coalesce(public.current_profile_role(), '') = 'admin'
    or p_province = any(coalesce(public.current_profile_provinces(), array[]::text[]))
$$;

-- ── 4. Policies profiles ──────────────────────────────────────────────────────
drop policy if exists "profiles_select_own_or_admin" on public.profiles;
create policy "profiles_select_own_or_admin" on public.profiles
  for select to authenticated using (auth.uid() = id or public.is_super_admin());

drop policy if exists "profiles_update_super_admin" on public.profiles;
create policy "profiles_update_super_admin" on public.profiles
  for update to authenticated using (public.is_super_admin()) with check (public.is_super_admin());

drop policy if exists "profiles_insert_super_admin" on public.profiles;
create policy "profiles_insert_super_admin" on public.profiles
  for insert to authenticated with check (public.is_super_admin());

drop policy if exists "profiles_delete_super_admin" on public.profiles;
create policy "profiles_delete_super_admin" on public.profiles
  for delete to authenticated using (public.is_super_admin());

-- ── 5. Verrouillage des RLS : retrait de l'accès anonyme ─────────────────────
drop policy if exists "provinces_select" on provinces;
drop policy if exists "provinces_insert" on provinces;
drop policy if exists "provinces_update" on provinces;
drop policy if exists "provinces_delete" on provinces;
create policy "provinces_select_auth" on provinces for select to authenticated using (true);
create policy "provinces_write_super_admin" on provinces for all to authenticated
  using (public.is_super_admin()) with check (public.is_super_admin());

drop policy if exists "missions_ref_select" on missions_ref;
drop policy if exists "missions_ref_insert" on missions_ref;
drop policy if exists "missions_ref_update" on missions_ref;
drop policy if exists "missions_ref_delete" on missions_ref;
create policy "missions_ref_select_auth" on missions_ref for select to authenticated using (true);
create policy "missions_ref_write_super_admin" on missions_ref for all to authenticated
  using (public.is_super_admin()) with check (public.is_super_admin());

drop policy if exists "territoires_select" on territoires;
drop policy if exists "territoires_insert" on territoires;
drop policy if exists "territoires_update" on territoires;
drop policy if exists "territoires_delete" on territoires;
create policy "territoires_select_auth" on territoires for select to authenticated using (true);
create policy "territoires_write_super_admin" on territoires for all to authenticated
  using (public.is_super_admin()) with check (public.is_super_admin());

drop policy if exists "secteurs_select" on secteurs;
drop policy if exists "secteurs_insert" on secteurs;
drop policy if exists "secteurs_update" on secteurs;
drop policy if exists "secteurs_delete" on secteurs;
create policy "secteurs_select_auth" on secteurs for select to authenticated using (true);
create policy "secteurs_write_super_admin" on secteurs for all to authenticated
  using (public.is_super_admin()) with check (public.is_super_admin());

drop policy if exists "reports_select" on reports;
drop policy if exists "reports_insert" on reports;
drop policy if exists "reports_update" on reports;
drop policy if exists "reports_delete" on reports;
-- Une policy "Allow update for authenticated users" (USING true) créée
-- manuellement hors migration a été retrouvée en production et supprimée :
-- elle autorisait tout utilisateur connecté à modifier n'importe quel rapport.
drop policy if exists "Allow update for authenticated users" on public.reports;
create policy "reports_select_auth" on reports for select to authenticated
  using (public.has_province_access(province));
create policy "reports_insert_auth" on reports for insert to authenticated
  with check (public.has_province_access(province));
create policy "reports_update_auth" on reports for update to authenticated
  using (public.has_province_access(province)) with check (public.has_province_access(province));
create policy "reports_delete_auth" on reports for delete to authenticated
  using (public.has_province_access(province));

drop policy if exists "liquidation_sequences_all" on liquidation_sequences;
create policy "liquidation_sequences_auth" on liquidation_sequences for all to authenticated
  using (public.current_profile_role() in ('comptable','admin','super_admin'))
  with check (public.current_profile_role() in ('comptable','admin','super_admin'));

-- Ancienne table users : conservée pour historique/traçabilité, réservée super_admin
drop policy if exists "users_select" on users;
drop policy if exists "users_insert" on users;
drop policy if exists "users_update" on users;
drop policy if exists "users_delete" on users;
create policy "users_super_admin_only" on users for all to authenticated
  using (public.is_super_admin()) with check (public.is_super_admin());

-- ── 6. Fonctions RPC : exécution réservée aux utilisateurs authentifiés ──────
create or replace function public.next_liquidation_seq(p_province text, p_year int)
returns int language plpgsql security definer set search_path = public as $$
declare v_seq int;
begin
  insert into liquidation_sequences(province, year, last_seq)
  values (p_province, p_year, 1)
  on conflict (province, year) do update
    set last_seq = liquidation_sequences.last_seq + 1
  returning last_seq into v_seq;
  return v_seq;
end;
$$;

revoke execute on function next_liquidation_seq(text, int) from anon, public;
grant  execute on function next_liquidation_seq(text, int) to authenticated;

-- ── 7. Défense en profondeur : retrait des GRANTs de table pour anon ─────────
revoke select on public.provinces, public.missions_ref, public.territoires, public.secteurs,
                   public.reports, public.users, public.profiles, public.liquidation_sequences
  from anon;

revoke execute on function public.current_profile_role()        from anon, public;
revoke execute on function public.current_profile_provinces()   from anon, public;
revoke execute on function public.is_super_admin()               from anon, public;
revoke execute on function public.has_province_access(text)      from anon, public;

grant execute on function public.current_profile_role()        to authenticated;
grant execute on function public.current_profile_provinces()   to authenticated;
grant execute on function public.is_super_admin()               to authenticated;
grant execute on function public.has_province_access(text)      to authenticated;

-- ── 8. Vérification ───────────────────────────────────────────────────────────
-- SELECT p.email, p.role, p.provinces, (au.encrypted_password IS NOT NULL) AS has_pw
--   FROM public.profiles p JOIN auth.users au ON au.id = p.id ORDER BY p.role, p.email;

-- ── 9. Étape manuelle requise ─────────────────────────────────────────────────
-- Déployer la fonction Edge `admin-users` (dossier supabase/functions/admin-users
-- si vous utilisez la CLI Supabase, ou via le Dashboard > Edge Functions) : elle
-- utilise la clé service_role pour créer/modifier/supprimer des comptes Auth +
-- profils depuis superadmin.html, avec vérification que l'appelant est
-- super_admin. Voir README.md pour le code source de la fonction.
