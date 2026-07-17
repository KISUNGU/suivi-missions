-- ════════════════════════════════════════════════════════════════════════════
-- Migration V23 — Corriger le contournement de has_province_access()
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte : has_province_access(p_province) autorisait tout accès (SELECT/
-- INSERT/UPDATE/DELETE) dès que p_province ne correspondait pas EXACTEMENT
-- (casse + espaces) à une ligne de la table `provinces`. Un compte comptable
-- restreint à ['UN'] pouvait donc saisir/valider un rapport avec
-- province = "KWILU" (majuscules) : cette valeur ne matchait pas
-- exactement "Kwilu" en base, donc la clause `not exists (...)` renvoyait
-- vrai et l'accès était accordé à tort. Résultat concret observé : un
-- rapport saisi par compte_uncp@pnda.cd (TSHIENDA Benjamin) s'est retrouvé
-- tagué province="KWILU" et visible dans le compte du comptable Kwilu.
--
-- Correctif : comparer les provinces en case-insensitive/trim des deux
-- côtés (province de l'utilisateur ET vérification d'existence), pour ne
-- garder le contournement que pour de véritables provinces inconnues/
-- nouvelles, pas pour de simples variantes de casse d'une province existante.
-- ════════════════════════════════════════════════════════════════════════════

create or replace function public.has_province_access(p_province text)
returns boolean language sql stable security definer set search_path to 'public'
as $$
  select
    coalesce(public.is_super_admin(), false)
    or coalesce(public.current_profile_role(), '') = 'admin'
    or lower(trim(p_province)) = any(
         select lower(trim(x))
         from unnest(coalesce(public.current_profile_provinces(), array[]::text[])) as x
       )
    or not exists (
         select 1 from public.provinces where lower(trim(name)) = lower(trim(p_province))
       )
$$;

-- Correction de données associée (déjà appliquée en base) :
-- update reports set province = 'UN'
--   where id = '08d617f5-a69c-42ac-955d-3570f900559d';
