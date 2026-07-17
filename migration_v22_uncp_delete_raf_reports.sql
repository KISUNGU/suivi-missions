-- ════════════════════════════════════════════════════════════════════════════
-- Migration V22 — Autoriser UNCP à supprimer les rapports RAF/RNSE visibles
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte : après la V21, le compte UNCP peut voir les rapports du RAF/RNSE
-- même hors province UN, mais la policy DELETE restait sur le critère générique
-- has_province_access(province). Résultat : le bouton Supprimer s'affiche,
-- mais Supabase refuse le DELETE pour un rapport RAF/RNSE hors UN, puis la
-- ligne réapparaît au rafraîchissement.
--
-- Règle métier appliquée ici : le comptable UNCP peut supprimer les rapports
-- RAF/RNSE qu'il est autorisé à consulter. Les autres comptables provinciaux
-- ne gagnent aucun droit supplémentaire.
-- ════════════════════════════════════════════════════════════════════════════

create or replace function public.is_uncp_comptable()
returns boolean language sql security definer set search_path = public stable as $$
  select
    coalesce(public.current_profile_role(), '') = 'comptable'
    and 'UN' = any(coalesce(public.current_profile_provinces(), array[]::text[]))
$$;

drop policy if exists "reports_delete_auth" on reports;
create policy "reports_delete_auth" on reports for delete to authenticated
  using (
    case
      when rapporteur_role in ('admin', 'super_admin') then
        coalesce(public.is_super_admin(), false)
        or coalesce(public.current_profile_role(), '') = 'admin'
        or public.is_uncp_comptable()
      else
        public.has_province_access(province)
    end
  );

-- Vérification conseillée :
-- 1) compte_uncp@pnda.cd supprime bien un rapport RAF/RNSE visible
-- 2) compte_kwilu@pnda.cd ne peut toujours pas supprimer un rapport RAF/RNSE
-- 3) les suppressions classiques par province restent inchangées