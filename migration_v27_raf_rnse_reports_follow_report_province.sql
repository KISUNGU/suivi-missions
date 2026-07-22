-- ============================================================================
-- Migration V27 - RAF/RNSE: gestion des rapports selon la province du rapport
-- ============================================================================
-- Nouvelle regle metier:
-- - RAF et RNSE peuvent effectuer des missions dans toutes les provinces.
-- - Le traitement comptable suit la province du rapport (pas UN force).
-- - La suppression reste reservee RAF/RNSE (deja geree par la migration V26).
--
-- Cette migration remet les policies SELECT/UPDATE des rapports sur une logique
-- purement province-based via has_province_access(province).
-- ============================================================================

drop policy if exists "reports_select_auth" on public.reports;
create policy "reports_select_auth"
  on public.reports
  for select
  to authenticated
  using (public.has_province_access(province));

drop policy if exists "reports_update_auth" on public.reports;
create policy "reports_update_auth"
  on public.reports
  for update
  to authenticated
  using (public.has_province_access(province))
  with check (public.has_province_access(province));

-- Verification conseillee:
-- 1) rapport RAF/RNSE avec province='Kwilu' visible/traitable par comptable Kwilu
-- 2) meme rapport non visible pour un comptable hors province (sauf admin/super_admin)
-- 3) suppression directe toujours refusee pour comptable (V26)
