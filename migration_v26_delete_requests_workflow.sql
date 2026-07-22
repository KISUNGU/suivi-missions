-- ============================================================================
-- Migration V26 - Workflow de demande de suppression des rapports
-- ============================================================================
-- Regle metier:
-- - Seuls RAF (admin) et RNSE (super_admin) peuvent supprimer un rapport.
-- - Les comptables envoient une demande de suppression.
-- - La demande reste visible (statut pending) jusqu'a execution/rejet.
-- ============================================================================

create table if not exists public.report_delete_requests (
  id uuid primary key default gen_random_uuid(),
  report_id text not null,
  province text,
  province_label text,
  rapporteur text,
  requested_by_email text not null,
  requested_by_role text not null,
  requested_at timestamptz not null default now(),
  reason text,
  status text not null default 'pending' check (status in ('pending', 'executed', 'rejected')),
  processed_by_email text,
  processed_at timestamptz,
  decision_note text
);

create index if not exists idx_report_delete_requests_report_id
  on public.report_delete_requests (report_id);

create index if not exists idx_report_delete_requests_status_requested_at
  on public.report_delete_requests (status, requested_at desc);

create unique index if not exists uq_report_delete_requests_pending_report
  on public.report_delete_requests (report_id)
  where status = 'pending';

alter table public.report_delete_requests enable row level security;

-- Lecture:
-- - RAF/RNSE voient toutes les demandes
-- - le demandeur voit ses demandes

drop policy if exists "report_delete_requests_select_auth" on public.report_delete_requests;
create policy "report_delete_requests_select_auth"
  on public.report_delete_requests
  for select
  to authenticated
  using (
    coalesce(public.current_profile_role(), '') in ('admin', 'super_admin')
    or lower(coalesce(requested_by_email, '')) = lower(coalesce(auth.jwt() ->> 'email', ''))
  );

-- Insertion:
-- - uniquement comptable
-- - l'email demandeur doit correspondre au token

drop policy if exists "report_delete_requests_insert_comptable" on public.report_delete_requests;
create policy "report_delete_requests_insert_comptable"
  on public.report_delete_requests
  for insert
  to authenticated
  with check (
    coalesce(public.current_profile_role(), '') = 'comptable'
    and status = 'pending'
    and lower(coalesce(requested_by_email, '')) = lower(coalesce(auth.jwt() ->> 'email', ''))
  );

-- Mise a jour / suppression:
-- - RAF/RNSE seulement

drop policy if exists "report_delete_requests_update_admin" on public.report_delete_requests;
create policy "report_delete_requests_update_admin"
  on public.report_delete_requests
  for update
  to authenticated
  using (coalesce(public.current_profile_role(), '') in ('admin', 'super_admin'))
  with check (coalesce(public.current_profile_role(), '') in ('admin', 'super_admin'));

drop policy if exists "report_delete_requests_delete_admin" on public.report_delete_requests;
create policy "report_delete_requests_delete_admin"
  on public.report_delete_requests
  for delete
  to authenticated
  using (coalesce(public.current_profile_role(), '') in ('admin', 'super_admin'));

-- Durcissement de la suppression des rapports:
-- RAF/RNSE uniquement (les comptables passent par report_delete_requests).

drop policy if exists "reports_delete_auth" on public.reports;
create policy "reports_delete_auth"
  on public.reports
  for delete
  to authenticated
  using (
    coalesce(public.is_super_admin(), false)
    or coalesce(public.current_profile_role(), '') = 'admin'
  );

-- Verifications conseillees:
-- 1) comptable: insert public.report_delete_requests -> OK
-- 2) comptable: delete public.reports -> REFUSE (RLS)
-- 3) admin/super_admin: select/update/delete public.report_delete_requests -> OK
-- 4) admin/super_admin: delete public.reports -> OK
