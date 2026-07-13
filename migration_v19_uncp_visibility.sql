-- ════════════════════════════════════════════════════════════════════════════
-- Migration V19 — Visibilité UNCP : RAF, RNSE & missionnaires Kinshasa
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte : le compte comptable UNCP (compte_uncp@pnda.cd, provinces = ['UN'])
-- ne voit aujourd'hui, via la RLS "reports_select_auth", que les rapports dont
-- la colonne province = 'UN' (donc les missionnaires rattachés à Kinshasa).
--
-- Objectif : les rapports élaborés par le RAF (role = 'admin') et le RNSE
-- (role = 'super_admin') doivent aussi être visibles par le compte UNCP,
-- quelle que soit la province choisie sur le rapport (ces rôles ne sont pas
-- rattachés à une province unique). Les autres comptables provinciaux
-- (Kwilu, Kasaï, Kasaï Central) ne sont PAS concernés par cet élargissement.
--
-- À exécuter après migration_v10_supabase_auth.sql dans l'éditeur SQL Supabase.
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1. Fonction utilitaire : le profil courant est-il le comptable UNCP ? ────
create or replace function public.is_uncp_comptable()
returns boolean language sql security definer set search_path = public stable as $$
  select
    coalesce(public.current_profile_role(), '') = 'comptable'
    and 'UN' = any(coalesce(public.current_profile_provinces(), array[]::text[]))
$$;

-- ── 2. Policy de lecture des rapports : ajout du cas RAF / RNSE pour UNCP ────
drop policy if exists "reports_select_auth" on reports;
create policy "reports_select_auth" on reports for select to authenticated
  using (
    public.has_province_access(province)
    or (public.is_uncp_comptable() and rapporteur_role in ('admin', 'super_admin'))
  );

-- Remarque : seule la policy de SELECT est élargie. Les policies d'insertion,
-- mise à jour et suppression restent inchangées (has_province_access) : le
-- compte UNCP peut donc voir ces rapports, mais ne peut pas les modifier ou
-- les supprimer si leur province n'est pas 'UN'.

-- Vérification conseillée après exécution :
-- SELECT id, province, rapporteur, rapporteur_role FROM reports
-- WHERE rapporteur_role IN ('admin','super_admin') ORDER BY saved_at DESC;
