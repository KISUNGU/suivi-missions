-- ─────────────────────────────────────────────────────────────────────────────
-- PNDA-SE — Migration v2 : Plateforme de gestion des missions
-- Exécuter dans l'éditeur SQL Supabase APRÈS supabase_db.sql
--
-- ⚠ NON UTILISÉE EN PRODUCTION : les tables `missions` et `mission_sequences`
-- créées ici ne sont référencées par aucune page de l'application (index.html,
-- admin.html, superadmin.html utilisent exclusivement la table `reports` avec
-- une colonne JSONB `missions`, définie dans supabase_db.sql). Ce script n'a
-- jamais été exécuté sur le projet Supabase de production. Conservé à titre
-- historique / pour une éventuelle migration future vers un schéma normalisé.
-- Corrigé le 03/07/2026 : la syntaxe `CREATE POLICY IF NOT EXISTS` n'existe pas
-- en PostgreSQL et aurait fait échouer ce script s'il avait été exécuté.
-- ─────────────────────────────────────────────────────────────────────────────

-- ── 1. MISE À JOUR TABLE USERS ────────────────────────────────────────────────

ALTER TABLE users ADD COLUMN IF NOT EXISTS display_name TEXT;
ALTER TABLE users ADD COLUMN IF NOT EXISTS is_active    BOOLEAN NOT NULL DEFAULT TRUE;
ALTER TABLE users ADD COLUMN IF NOT EXISTS created_at   TIMESTAMPTZ NOT NULL DEFAULT now();
ALTER TABLE users ADD COLUMN IF NOT EXISTS created_by   TEXT;

-- Promouvoir le compte rnse en super_admin
UPDATE users SET role = 'super_admin' WHERE email = 'rnse@pnda.cd';

-- ── 2. TABLE DES MISSIONS INDIVIDUELLES ──────────────────────────────────────
-- Chaque missionary enregistre ses missions ici avec un numéro séquentiel.

CREATE TABLE IF NOT EXISTS missions (
  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  seq_num          INT  NOT NULL,
  user_email       TEXT NOT NULL,
  province         TEXT NOT NULL,
  province_label   TEXT,
  created_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
  mission_date     DATE,
  nature_category  TEXT NOT NULL,   -- 'implementation' | 'activite' | 'atelier' | 'autre'
  nature_label     TEXT NOT NULL,
  objectif         TEXT NOT NULL DEFAULT '',
  lieu             TEXT,
  section          TEXT,
  trimestre        TEXT,
  annee            INT,
  rapporteur       TEXT,
  rapporteur_role  TEXT,
  hors_projet      INT  NOT NULL DEFAULT 0,
  projet           INT  NOT NULL DEFAULT 1,
  montant_usd      NUMERIC(12,2) NOT NULL DEFAULT 0,
  avance_usd       NUMERIC(12,2) NOT NULL DEFAULT 0,
  statut           TEXT NOT NULL DEFAULT 'saisie',
  UNIQUE (user_email, seq_num)
);

-- ── 3. TABLE DES SÉQUENCES (compteur par compte) ─────────────────────────────

CREATE TABLE IF NOT EXISTS mission_sequences (
  user_email TEXT PRIMARY KEY,
  last_seq   INT  NOT NULL DEFAULT 0
);

-- ── 4. FUNCTION : incrémenter le numéro séquentiel (sûr en concurrent) ────────

CREATE OR REPLACE FUNCTION next_mission_seq(p_email TEXT)
RETURNS INT LANGUAGE plpgsql AS $$
DECLARE v_seq INT;
BEGIN
  INSERT INTO mission_sequences(user_email, last_seq) VALUES (p_email, 1)
  ON CONFLICT (user_email) DO UPDATE
    SET last_seq = mission_sequences.last_seq + 1
  RETURNING last_seq INTO v_seq;
  RETURN v_seq;
END;
$$;

-- ── 5. RLS POUR LES NOUVELLES TABLES ─────────────────────────────────────────

ALTER TABLE missions          ENABLE ROW LEVEL SECURITY;
ALTER TABLE mission_sequences ENABLE ROW LEVEL SECURITY;

-- Missions : tout le monde peut lire ; seul l'auteur (par email) peut insérer/modifier
CREATE POLICY "missions_select" ON missions FOR SELECT TO anon USING (true);
CREATE POLICY "missions_insert" ON missions FOR INSERT TO anon WITH CHECK (true);
CREATE POLICY "missions_update" ON missions FOR UPDATE TO anon USING (true);
CREATE POLICY "missions_delete" ON missions FOR DELETE TO anon USING (true);

-- Sequences : accès complet via clé anon (gérée par l'app)
CREATE POLICY "mission_seq_all" ON mission_sequences
  FOR ALL TO anon USING (true) WITH CHECK (true);

-- Users : autoriser insert/update/delete pour la gestion par super_admin
-- (CREATE POLICY ne supporte pas IF NOT EXISTS en PostgreSQL ; on passe par
-- un bloc DO conditionnel, comme dans les autres scripts de migration.)
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='users' AND policyname='users_insert') THEN
    CREATE POLICY "users_insert" ON users FOR INSERT TO anon WITH CHECK (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='users' AND policyname='users_update') THEN
    CREATE POLICY "users_update" ON users FOR UPDATE TO anon USING (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='users' AND policyname='users_delete') THEN
    CREATE POLICY "users_delete" ON users FOR DELETE TO anon USING (true);
  END IF;
END $$;

-- ── 6. EXPOSER LA FONCTION next_mission_seq À LA CLÉ ANON ────────────────────
-- Obligatoire pour appeler la fonction depuis le client JS

GRANT EXECUTE ON FUNCTION next_mission_seq(TEXT) TO anon;
