-- ════════════════════════════════════════════════════════════════════════════
-- Migration V24 — Autoriser UNCP à activer/mettre à jour les rapports RAF/RNSE
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte : le compte comptable UNCP peut consulter (V21) et supprimer (V22)
-- les rapports créés par RAF/RNSE, mais la policy UPDATE restait limitée à
-- has_province_access(province). Donc, pour un rapport RAF/RNSE hors province
-- UN, l'activation échoue (RLS), et le statut reste "en_attente".
--
-- Règle métier : UNCP doit pouvoir piloter le workflow (activation/statut)
-- des rapports RAF/RNSE qu'il est autorisé à consulter.
-- ════════════════════════════════════════════════════════════════════════════

create or replace function public.is_uncp_comptable()
returns boolean language sql security definer set search_path = public stable as $$
  select
    coalesce(public.current_profile_role(), '') = 'comptable'
    and 'UN' = any(coalesce(public.current_profile_provinces(), array[]::text[]))
$$;

drop policy if exists "reports_update_auth" on reports;
create policy "reports_update_auth" on reports for update to authenticated
  using (
    case
      when rapporteur_role in ('admin', 'super_admin') then
        coalesce(public.is_super_admin(), false)
        or coalesce(public.current_profile_role(), '') = 'admin'
        or public.is_uncp_comptable()
      else
        public.has_province_access(province)
    end
  )
  with check (
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
-- 1) compte_uncp@pnda.cd active un rapport RAF/RNSE hors UN -> succès
-- 2) compte_kwilu@pnda.cd ne peut pas modifier un rapport RAF/RNSE
-- 3) RAF/RNSE conservent leurs droits de mise à jour
