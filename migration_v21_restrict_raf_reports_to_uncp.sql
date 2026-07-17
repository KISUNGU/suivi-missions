-- ════════════════════════════════════════════════════════════════════════════
-- Migration V21 — Réserver les rapports RAF/RNSE au RAF, RNSE et UNCP
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte : après la V19, le compte UNCP devait voir les rapports élaborés
-- par le RAF (admin) et le RNSE (super_admin), quelle que soit leur province.
-- Mais la policy gardait aussi le chemin générique has_province_access(province),
-- ce qui laissait un comptable provincial voir un rapport RAF dès lors que la
-- province du rapport correspondait à sa province (ex. Kwilu).
--
-- Règle métier : les rapports saisis par le RAF/RNSE sont traités pour le
-- compte de l'UNCP. Ils ne doivent donc être visibles que par :
--   - le RAF (role = admin)
--   - le RNSE (role = super_admin)
--   - le comptable UNCP (profiles.provinces contient 'UN')
-- Les autres comptables/provinces ne doivent plus les voir.
-- ════════════════════════════════════════════════════════════════════════════

create or replace function public.is_uncp_comptable()
returns boolean language sql security definer set search_path = public stable as $$
  select
    coalesce(public.current_profile_role(), '') = 'comptable'
    and 'UN' = any(coalesce(public.current_profile_provinces(), array[]::text[]))
$$;

drop policy if exists "reports_select_auth" on reports;
create policy "reports_select_auth" on reports for select to authenticated
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
-- 1) compte_kwilu@pnda.cd ne doit plus voir de lignes avec rapporteur_role admin/super_admin
-- 2) compte_uncp@pnda.cd doit toujours les voir
-- 3) le RAF et le RNSE conservent leur accès complet