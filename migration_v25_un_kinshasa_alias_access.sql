-- ============================================================================
-- Migration V25 - Robustesse alias province UN/Kinshasa/UNCP
-- ============================================================================
-- Objectif:
-- 1) Eviter les pertes de visibilite lorsque les donnees utilisent tantot
--    'UN', tantot 'Kinshasa' (ou variantes) pour la meme province.
-- 2) Garantir que le compte comptable UNCP reste reconnu meme si son profil
--    contient un alias ('Kinshasa', 'UNCP', etc.) au lieu de 'UN'.
--
-- Cette migration est idempotente (CREATE OR REPLACE + UPDATE conditionnels).
-- ============================================================================

-- Normalise une valeur province vers une cle comparable.
-- Les alias centraux convergent vers 'un'.
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

-- Reconnaitre UNCP meme si profiles.provinces contient un alias de UN.
create or replace function public.is_uncp_comptable()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select
    coalesce(public.current_profile_role(), '') = 'comptable'
    and exists (
      select 1
      from unnest(coalesce(public.current_profile_provinces(), array[]::text[])) as p
      where public.normalize_province_key(p) = 'un'
    )
$$;

-- Controle d'acces province robuste aux alias et a la casse.
create or replace function public.has_province_access(p_province text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select
    coalesce(public.is_super_admin(), false)
    or coalesce(public.current_profile_role(), '') = 'admin'
    or public.normalize_province_key(p_province) = any (
      select public.normalize_province_key(x)
      from unnest(coalesce(public.current_profile_provinces(), array[]::text[])) as x
    )
    or not exists (
      select 1
      from public.provinces
      where public.normalize_province_key(name) = public.normalize_province_key(p_province)
    )
$$;

-- Harmonisation douce des profils contenant un alias de UN.
-- Exemples: ['Kinshasa'] -> ['UN'], ['UNCP','Kwilu'] -> ['UN','Kwilu']
update public.profiles
set provinces = (
  select coalesce(array_agg(v order by v), array[]::text[])
  from (
    select distinct
      case
        when public.normalize_province_key(x) = 'un' then 'UN'
        else trim(x)
      end as v
    from unnest(coalesce(public.profiles.provinces, array[]::text[])) as x
  ) normalized
)
where provinces is not null
  and exists (
    select 1
    from unnest(provinces) as x
    where public.normalize_province_key(x) = 'un' and trim(x) <> 'UN'
  );

-- Harmonisation douce des rapports deja enregistres avec alias de UN.
update public.reports
set province = 'UN',
    province_label = case
      when coalesce(trim(province_label), '') = '' then 'Kinshasa'
      else province_label
    end,
    province_code = case
      when coalesce(trim(province_code), '') = '' then 'UN'
      else province_code
    end
where public.normalize_province_key(province) = 'un'
  and province <> 'UN';

-- Verification conseillee:
-- 1) select email, role, provinces from public.profiles where email in ('rnse@pnda.cd','compte_uncp@pnda.cd');
-- 2) select id, province, province_label, rapporteur, rapporteur_role from public.reports
--    where rapporteur ilike '%rnse%' or public.normalize_province_key(province) = 'un'
--    order by saved_at desc limit 50;
