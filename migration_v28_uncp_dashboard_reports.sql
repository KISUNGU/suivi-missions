-- ============================================================================
-- Migration V28 — Correction province des rapports comptable UNCP
-- ============================================================================
-- Contexte :
--   Le compte comptable UNCP (compte_uncp@pnda.cd) voit 60 rapports dans
--   « Mes rapports » (saisie.html) mais seulement 11 dans le tableau de bord
--   (index.html).
--
-- Cause :
--   - saisie.html charge les rapports par rapporteur (tous les rapports du
--     compte sont visibles).
--   - index.html filtre les rapports par province pour les comptables ; UNCP
--     ne voit que les rapports dont la colonne province = 'UN'.
--   - Certains rapports saisis par compte_uncp ont une province différente
--     de 'UN' en base (anciennes saisies, imports, alias non harmonisés, etc.).
--
-- Action :
--   Réattribuer la province 'UN' aux rapports du comptable UNCP qui n'y sont
--   pas rattachés, afin qu'ils apparaissent dans son tableau de bord.
--
-- Cette migration est idempotente (les lignes déjà à 'UN' ne sont pas touchées).
-- ============================================================================

-- ── 1. Fonction utilitaire de normalisation (reprise de V25, sans danger) ────
create or replace function public.normalize_province_key(p_value text)
returns text
language sql
stable
security definer
set search_path = public
as $$
  select case
    when p_value is null then ''
    when lower(trim(p_value)) in ('un', 'kinshasa', 'uncp', 'unite nationale de coordination') then 'un'
    else lower(trim(p_value))
  end
$$;

-- ── 2. Diagnostic avant correction (optionnel, à exécuter séparément) ───────
-- select
--   province,
--   count(*) as nb_rapports
-- from public.reports
-- where rapporteur = 'compte_uncp@pnda.cd'
-- group by province
-- order by nb_rapports desc;

-- ── 3. Correction : rattacher les rapports UNCP à la province UN ─────────────
-- On cible les rapports créés par le compte UNCP, que le champ rapporteur
-- contienne son email ou son display_name.
with uncp_profile as (
  select display_name
  from public.profiles
  where email = 'compte_uncp@pnda.cd'
  limit 1
)
update public.reports
set
  province       = 'UN',
  province_label = case
                     when coalesce(trim(province_label), '') = '' then 'UNCP (Unité Nationale de Coordination)'
                     else province_label
                   end,
  province_code  = case
                     when coalesce(trim(province_code), '') = '' then '1'
                     else province_code
                   end,
  meta = jsonb_set(
           coalesce(meta, '{}'::jsonb),
           '{province}',
           '"UN"'::jsonb,
           true
         )
where (
        rapporteur = 'compte_uncp@pnda.cd'
        or rapporteur = (select display_name from uncp_profile)
      )
  and province is not null
  and public.normalize_province_key(province) <> 'un';

-- ── 4. Vérification après correction ─────────────────────────────────────────
-- select
--   province,
--   count(*) as nb_rapports
-- from public.reports
-- where rapporteur = 'compte_uncp@pnda.cd'
-- group by province;
