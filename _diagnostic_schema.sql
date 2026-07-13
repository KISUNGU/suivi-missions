-- Requête de diagnostic (lecture seule) — à exécuter dans l'éditeur SQL Supabase
-- et à coller le résultat complet dans le chat.

-- 1) Tables existantes dans le schéma public
select table_name
from information_schema.tables
where table_schema = 'public'
order by table_name;

-- 2) Fonctions existantes liées à l'auth/RLS
select routine_name, data_type as return_type
from information_schema.routines
where routine_schema = 'public'
  and (routine_name ilike '%admin%' or routine_name ilike '%province%' or routine_name ilike '%profile%' or routine_name ilike '%uncp%')
order by routine_name;

-- 3) Policies RLS actuellement actives sur la table reports
select policyname, cmd, qual, with_check
from pg_policies
where schemaname = 'public' and tablename = 'reports';

-- 4) Colonnes de la table utilisée pour le rôle/provinces de l'utilisateur
-- (adapter le nom de table selon le résultat de la requête 1 si ce n'est pas "profiles")
select column_name, data_type
from information_schema.columns
where table_schema = 'public' and table_name in ('profiles','users')
order by table_name, ordinal_position;
