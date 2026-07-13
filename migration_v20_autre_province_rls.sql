-- ════════════════════════════════════════════════════════════════════════════
-- Migration V20 — Autoriser les rapports "Autre" (province hors liste fixe)
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte : saisie.html propose désormais un choix de province limité aux 4
-- provinces de référence (UN/Kinshasa, Kwilu, Kasaï, Kasaï Central) + une
-- option "Autre" qui laisse saisir librement un nom de province (ex. mission
-- ponctuelle à Kongo Central). Ce texte libre ne correspond à aucune ligne de
-- la table `provinces`, or la policy RLS "reports_select_auth" /
-- "reports_insert_auth" (etc.) s'appuie sur has_province_access(province),
-- qui exige que la province soit dans la liste des provinces assignées au
-- profil (ou que le rôle soit admin/super_admin). Résultat : tout compte
-- "province" ou "comptable" qui soumet un rapport en mode "Autre" reçoit
-- l'erreur 403 "new row violates row-level security policy for table reports".
--
-- Correction : has_province_access() accorde aussi l'accès lorsque la
-- province indiquée ne correspond à AUCUNE ligne de la table `provinces`
-- (donc ne peut être qu'une saisie libre "Autre"). Les 4 provinces de
-- référence restent scopées comme avant (un compte Kwilu ne voit toujours
-- que les rapports Kwilu, etc.) ; seuls les rapports "Autre" deviennent
-- visibles/gérables par tout utilisateur authentifié (en plus de RAF/RNSE qui
-- voyaient déjà tout).
--
-- À exécuter après migration_v10_supabase_auth.sql (et migration_v19) dans
-- l'éditeur SQL Supabase.
-- ════════════════════════════════════════════════════════════════════════════

create or replace function public.has_province_access(p_province text)
returns boolean language sql security definer set search_path = public stable as $$
  select
    coalesce(public.is_super_admin(), false)
    or coalesce(public.current_profile_role(), '') = 'admin'
    or p_province = any(coalesce(public.current_profile_provinces(), array[]::text[]))
    or not exists (select 1 from public.provinces where name = p_province)
$$;

-- Vérification conseillée après exécution : se connecter avec un compte
-- "province" ou "comptable" (non admin/super_admin) et soumettre un rapport
-- via saisie.html en choisissant "Autre" pour la province -> doit réussir.
