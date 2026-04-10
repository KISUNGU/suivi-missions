-- ─────────────────────────────────────────────────────────────────────────────
-- PNDA-SE — Schéma Supabase
-- Copiez et exécutez ce script dans l'éditeur SQL de votre projet Supabase.
-- https://supabase.com/dashboard/project/YOUR_PROJECT/editor
-- ─────────────────────────────────────────────────────────────────────────────

-- ── 1. TABLES ─────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS provinces (
  id    UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name  TEXT UNIQUE NOT NULL,
  label TEXT NOT NULL,
  code  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS missions_ref (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  province_name TEXT NOT NULL REFERENCES provinces(name) ON DELETE CASCADE,
  ref_id        TEXT NOT NULL,
  nature        TEXT NOT NULL,
  objectif      TEXT NOT NULL DEFAULT '',
  ordre         INT  NOT NULL DEFAULT 0,
  category      TEXT NOT NULL DEFAULT 'implementation',   -- implementation | activite | atelier | autre
  UNIQUE (province_name, ref_id)
);

CREATE TABLE IF NOT EXISTS users (
  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  email        TEXT UNIQUE NOT NULL,
  password_b64 TEXT NOT NULL,
  display_name TEXT,
  provinces    TEXT[],          -- NULL = accès à toutes (admin / super_admin)
  role         TEXT NOT NULL DEFAULT 'province'
);

CREATE TABLE IF NOT EXISTS reports (
  id               TEXT PRIMARY KEY,
  saved_at         TIMESTAMPTZ NOT NULL DEFAULT now(),
  source           TEXT DEFAULT 'formulaire-html',
  province         TEXT NOT NULL,
  province_label   TEXT,
  province_code    TEXT,
  trimestre        TEXT,
  annee            INT,
  rapporteur       TEXT,
  rapporteur_role  TEXT,
  date_generation  DATE,
  seq_number       INT,
  missions         JSONB NOT NULL DEFAULT '[]',
  meta             JSONB NOT NULL DEFAULT '{}'
);

-- ── 2. ROW LEVEL SECURITY ─────────────────────────────────────────────────────

ALTER TABLE provinces    ENABLE ROW LEVEL SECURITY;
ALTER TABLE missions_ref ENABLE ROW LEVEL SECURITY;
ALTER TABLE users        ENABLE ROW LEVEL SECURITY;
ALTER TABLE reports      ENABLE ROW LEVEL SECURITY;

-- Lecture publique (clé anon) pour provinces, missions_ref, users
CREATE POLICY "provinces_select"    ON provinces    FOR SELECT TO anon USING (true);
CREATE POLICY "missions_ref_select" ON missions_ref FOR SELECT TO anon USING (true);
CREATE POLICY "users_select"        ON users        FOR SELECT TO anon USING (true);

-- Gestion des utilisateurs (super_admin via clé anon)
CREATE POLICY "users_insert" ON users FOR INSERT TO anon WITH CHECK (true);
CREATE POLICY "users_update" ON users FOR UPDATE TO anon USING (true);
CREATE POLICY "users_delete" ON users FOR DELETE TO anon USING (true);

-- Rapports : lecture, insertion, mise à jour, suppression
CREATE POLICY "reports_select" ON reports FOR SELECT TO anon USING (true);
CREATE POLICY "reports_insert" ON reports FOR INSERT TO anon WITH CHECK (true);
CREATE POLICY "reports_update" ON reports FOR UPDATE TO anon USING (true);
CREATE POLICY "reports_delete" ON reports FOR DELETE TO anon USING (true);

-- ── 3. DONNÉES DE RÉFÉRENCE — PROVINCES ───────────────────────────────────────

INSERT INTO provinces (name, label, code) VALUES
  ('UN',            'UNCP (Unité Nationale de Coordination)', '1'),
  ('Kasaï Central', 'Kasaï Central',                         '2'),
  ('Kwilu',         'Kwilu',                                  '3'),
  ('Kasaï',         'Kasaï',                                  '4')
ON CONFLICT (name) DO UPDATE SET label = EXCLUDED.label, code = EXCLUDED.code;

-- ── 4. DONNÉES DE RÉFÉRENCE — MISSIONS ────────────────────────────────────────

-- Province UN
INSERT INTO missions_ref (province_name, ref_id, nature, objectif, ordre) VALUES
  ('UN','2_1_UN',  'Atelier formation prévention & sécurité routière (Kananga)',       'Former les utilisateurs des engins roulants', 1),
  ('UN','2_2_UN',  'Enquête production PNDA (Kwilu, Kasaï Central, Kasaï)',             'Collecte données indicateurs ODP PNDA',       2),
  ('UN','2_3_UN',  'Participation ateliers restitution EIES (Kananga)',                 'Participer à la restitution des EIES',         3),
  ('UN','2_4_UN',  'Enquête production PNDA (répétition ligne 2.2)',                   'Collecte données indicateurs ODP PNDA',        4),
  ('UN','2_5_UN',  'Appui technique terrain – préparation supervision',                'Appuyer consultant pour rapport mi-parcours',  5),
  ('UN','2_6_UN',  'Appui technique – préparation supervision conjointe',              'Appui organisation matérielle mission technique',6),
  ('UN','2_7_UN',  'Supervision terrain Kikwit & Kasaï',                               'Accompagner TTL & vérifier aspects environnementaux',7),
  ('UN','2_8_UN',  'Remise officielle motos – Kongo Central',                          'Renforcer capacités agents de l''État',        8),
  ('UN','2_9_UN',  'Examen & validation PTBA 2026 (COPIL)',                            'Examiner & approuver PTBA avant transmission au bailleur',9),
  ('UN','8_1_UN',  'Atelier revue à mi-parcours (Kinshasa)',                           'Participer à la revue à mi-parcours du PNDA', 10)
ON CONFLICT (province_name, ref_id) DO UPDATE SET nature = EXCLUDED.nature, objectif = EXCLUDED.objectif;

-- Province Kwilu
INSERT INTO missions_ref (province_name, ref_id, nature, objectif, ordre) VALUES
  ('Kwilu','2_1_Kwilu',         'Mise en œuvre activités','Collecte données enquêtes production PEA',1),
  ('Kwilu','2_2_Kwilu',         'Mise en œuvre activités','Suivi contentieux campagne agricole A2025 (Gungu)',2),
  ('Kwilu','2_3_Kwilu',         'Mise en œuvre activités','Sensibilisation + supervision campagne A & vente mucuna',3),
  ('Kwilu','2_4_Kwilu',         'Mise en œuvre activités','Appui & supervision commercialisation boutures manioc/mucuna/maïs (Bulungu, Gungu, Idiofa)',4),
  ('Kwilu','2_5_Kwilu',         'Mise en œuvre activités','Appui logistique distribution carburant & huile moteur',5),
  ('Kwilu','2_6_Kwilu',         'Mise en œuvre activités','Appui fonctionnement UPEP Kwilu (paiement AC)',6),
  ('Kwilu','2_7_Kwilu',         'Mise en œuvre activités','Briefing & installation CGP',7),
  ('Kwilu','2_8_Kwilu',         'Mise en œuvre activités','Collecte & production histoires à succès (3 territoires)',8),
  ('Kwilu','2_9_Kwilu',         'Mise en œuvre activités','Suivi fonctionnement CGP dans les territoires',9),
  ('Kwilu','2_1_Kwilu_FACT_9_8','Mise en œuvre activités','Missions FACT 9/8 – entretien véhicules INERA & pêche/élevage',10),
  ('Kwilu','2_11_Kwilu',        'Mise en œuvre activités','Installation parcs à bois + vente boutures fortifiées',11),
  ('Kwilu','2_12_Kwilu',        'Mise en œuvre activités','Distribution matériels aratoires aux OP',12),
  ('Kwilu','3_1_Kwilu',         'Atelier / Réunion supervision & revue','Participation supervision & revue mi-parcours PNDA (Kinshasa)',13),
  ('Kwilu','3_2_Kwilu',         'Atelier de restitution','Restitution APS BETRA/WEST (Kinshasa)',14),
  ('Kwilu','3_3_Kwilu',         'Formation','Remise à niveau en webmastering (communication)',15),
  ('Kwilu','3_4_Kwilu',         'Formation','Remise à niveau en passation de marchés (STEP & contrats)',16)
ON CONFLICT (province_name, ref_id) DO UPDATE SET nature = EXCLUDED.nature, objectif = EXCLUDED.objectif;

-- Province Kasaï
INSERT INTO missions_ref (province_name, ref_id, nature, objectif, ordre) VALUES
  ('Kasaï','1_1_Kasaï',                 'Accompagnement Assistant SG Agriculture à Mweka','Conduire le véhicule',1),
  ('Kasaï','1_2_Kasaï',                 'Distribution semences mucuna','Déposer les semences',2),
  ('Kasaï','1_3_Kasaï',                 'Appui UPEP Kwilu (gestion financière)','Appuyer UPEP',3),
  ('Kasaï','1_4_Kasaï',                 'Accompagnement comptable Kwilu','Conduire comptable + récupérer semences',4),
  ('Kasaï','1_5_Kasaï',                 'Enquête production agricole Tshikapa','Superviser enquête',5),
  ('Kasaï','1_6_Kasaï',                 'Appui campagne agricole A 2025','Appuyer campagne',6),
  ('Kasaï','1_7_Kasaï',                 'Récupération comptable Kwilu','Récupérer comptable',7),
  ('Kasaï','1_8_Kasaï',                 'Remise à niveau CPM Banque mondiale','Formation STEP + contrats',8),
  ('Kasaï','1_9_Kasaï',                 'Accompagnement CPM Kikwit','Conduire CPM',9),
  ('Kasaï','1_1_Kasaï_CGP_Mweka',       'Briefing + installation CGP Mweka','Installer CGP',10),
  ('Kasaï','1_11_Kasaï',                'Briefing + installation CGP Luebo','Installer CGP',11),
  ('Kasaï','1_12_Kasaï',                'Briefing + installation CGP Tshikapa','Installer CGP',12),
  ('Kasaï','1_13_Kasaï',                'Récupération semences FAO Kikwit','Récupérer semences',13),
  ('Kasaï','1_14_Kasaï',                'Mise à niveau communication (webmastering)','Formation communication',14),
  ('Kasaï','1_15_Kasaï',                'Appui UPEP Kwilu','Appuyer UPEP',15),
  ('Kasaï','1_16_Kasaï',                'Appui fonctionnement KWILU','Appuyer activités',16),
  ('Kasaï','1_17_Kasaï',                'Visite agrimultiplicateur','Déposer semences',17),
  ('Kasaï','1_18_Kasaï',                'Suivi réparation véhicules Kikwit','Suivre réparation',18),
  ('Kasaï','1_19_Kasaï',                'Supervision technique Tshikapa','Supervision',19),
  ('Kasaï','1_2_Kasaï_Revue_Mi_Parcours','Participation Revue Mi-Parcours','',20),
  ('Kasaï','1_21_Kasaï',                'Captage histoires à succès','Capturer histoires + réunions',21),
  ('Kasaï','1_22_Kasaï',                'Accompagnement comptable Kikwit','Accompagner comptable',22),
  ('Kasaï','1_23_Kasaï',                'Appui UPEP Kwilu','Appuyer UPEP',23),
  ('Kasaï','1_24_Kasaï',                'Atelier restitution APS BETRA/WEST','Participer atelier',24),
  ('Kasaï','1_25_Kasaï',                'Accompagnement ETGC Kikwit','Conduire véhicule',25),
  ('Kasaï','1_26_Kasaï',                'Suivi réparation véhicules Kikwit','Suivre réparation',26),
  ('Kasaï','1_27_Kasaï',                'Récupération comptable + véhicule Kwilu','Récupérer comptable + véhicule',27),
  ('Kasaï','1_28_Kasaï',                'Mission conjointe UNOPS/AGETIP','Reconnaissance axes routiers',28),
  ('Kasaï','3_1_Kasaï',                 'Atelier restitution EIES Kamonia','Participer atelier',29)
ON CONFLICT (province_name, ref_id) DO UPDATE SET nature = EXCLUDED.nature, objectif = EXCLUDED.objectif;

-- Province Kasaï Central
INSERT INTO missions_ref (province_name, ref_id, nature, objectif, ordre) VALUES
  ('Kasaï Central','2_1_Kasaï_Central',      'Supervision & accompagnement enquête PEA (Demba, Dibaya)','Enquêter bénéficiaires PNDA (B-2024, A-2024, B-2025)',1),
  ('Kasaï Central','2_2_Kasaï_Central',      'Déploiement semences mucuna (Demba, Luiza)','Appui adoption agriculture intelligente',2),
  ('Kasaï Central','2_3_Kasaï_Central',      'Briefing CGP sur MGP (Luiza, Dibaya, Demba)','Renforcement capacités CGP',3),
  ('Kasaï Central','2_4_Kasaï_Central',      'Évaluation boutures manioc FAO/INERA Ngandajika','Augmenter production manioc',4),
  ('Kasaï Central','2_5_Kasaï_Central',      'Supervision vente semences (Luiza, Dibaya, Demba)','Suivi achats/ventes semences',5),
  ('Kasaï Central','2_6_Kasaï_Central',      'Formation sécurité routière','Renforcer capacités utilisateurs engins roulants',6),
  ('Kasaï Central','2_7_Kasaï_Central',      'Suivi mise en œuvre PNDA (Hinterland Kananga)','Suivi activités PNDA',7),
  ('Kasaï Central','2_8_Kasaï_Central',      'Formation mise à niveau animateurs webmastering (UNCP)','Renforcement capacités communication',8),
  ('Kasaï Central','2_9_Kasaï_Central',      'Participation revue à mi-parcours (UNCP)','Participation revue PNDA',9),
  ('Kasaï Central','2_1_Kasaï_Central_AVEC', 'Livraison boutures manioc aux AVEC','Appui champs semenciers communautaires',10),
  ('Kasaï Central','2_11_Kasaï_Central',     'Production capsules histoires de succès','Élaboration histoires de succès PNDA',11),
  ('Kasaï Central','2_12_Kasaï_Central',     'Participation atelier validation PTBA/Kikwit','Validation PTBA',12)
ON CONFLICT (province_name, ref_id) DO UPDATE SET nature = EXCLUDED.nature, objectif = EXCLUDED.objectif;

-- ── 5. COMPTES UTILISATEURS ───────────────────────────────────────────────────
-- Les mots de passe sont identiques à ceux existants (encodage base64 original).
-- Pour changer un mot de passe : btoa("NouveauMotDePasse") dans la console du navigateur.

INSERT INTO users (email, password_b64, provinces, role) VALUES
  ('compte_kwilu@pnda.cd',  'S3dpbHUjMjAyNg==', ARRAY['Kwilu'],          'comptable'),
  ('compte_kasaic@pnda.cd', 'S2FzYWlDIzIwMjY=', ARRAY['Kasaï Central'], 'comptable'),
  ('compte_kasai@pnda.cd',  'S2FzYWkjMjAyNg==', ARRAY['Kasaï'],         'comptable'),
  ('compte_uncp@pnda.cd',   'VU5DUCMyMDI2',     ARRAY['UN'],             'comptable'),
  ('rnse@pnda.cd',          'Uk5TRSMyMDI2',     NULL,                    'super_admin'),
  ('raf@pnda.cd',           'UkFGIzIwMjY=',     NULL,                    'admin')
ON CONFLICT (email) DO NOTHING;

-- ── MIGRATION v2 — À exécuter si la v1 a déjà été appliquée ───────────────────
-- (idempotent grâce aux IF NOT EXISTS / ON CONFLICT)

ALTER TABLE missions_ref ADD COLUMN IF NOT EXISTS category TEXT NOT NULL DEFAULT 'implementation';
ALTER TABLE reports      ADD COLUMN IF NOT EXISTS seq_number INT;
ALTER TABLE users        ADD COLUMN IF NOT EXISTS display_name TEXT;

-- ── MIGRATION v3 — Territoires & Secteurs ────────────────────────────────────

CREATE TABLE IF NOT EXISTS territoires (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  province_name TEXT NOT NULL REFERENCES provinces(name) ON DELETE CASCADE,
  name          TEXT NOT NULL,
  label         TEXT NOT NULL,
  UNIQUE (province_name, name)
);

CREATE TABLE IF NOT EXISTS secteurs (
  id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  territoire_id   UUID NOT NULL REFERENCES territoires(id) ON DELETE CASCADE,
  name            TEXT NOT NULL,
  label           TEXT NOT NULL,
  UNIQUE (territoire_id, name)
);

ALTER TABLE territoires ENABLE ROW LEVEL SECURITY;
ALTER TABLE secteurs    ENABLE ROW LEVEL SECURITY;

-- Politiques RLS pour territoires et secteurs
DO $$ BEGIN
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='territoires' AND policyname='territoires_select') THEN
    CREATE POLICY "territoires_select" ON territoires FOR SELECT TO anon USING (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='territoires' AND policyname='territoires_insert') THEN
    CREATE POLICY "territoires_insert" ON territoires FOR INSERT TO anon WITH CHECK (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='territoires' AND policyname='territoires_update') THEN
    CREATE POLICY "territoires_update" ON territoires FOR UPDATE TO anon USING (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='territoires' AND policyname='territoires_delete') THEN
    CREATE POLICY "territoires_delete" ON territoires FOR DELETE TO anon USING (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='secteurs' AND policyname='secteurs_select') THEN
    CREATE POLICY "secteurs_select" ON secteurs FOR SELECT TO anon USING (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='secteurs' AND policyname='secteurs_insert') THEN
    CREATE POLICY "secteurs_insert" ON secteurs FOR INSERT TO anon WITH CHECK (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='secteurs' AND policyname='secteurs_update') THEN
    CREATE POLICY "secteurs_update" ON secteurs FOR UPDATE TO anon USING (true);
  END IF;
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='secteurs' AND policyname='secteurs_delete') THEN
    CREATE POLICY "secteurs_delete" ON secteurs FOR DELETE TO anon USING (true);
  END IF;
END $$;

-- ── SEED : Territoires et Secteurs ────────────────────────────────────────────

-- Kwilu
INSERT INTO territoires (province_name, name, label) VALUES
  ('Kwilu', 'BULUNGU', 'Bulungu'),
  ('Kwilu', 'IDIOFA',  'Idiofa'),
  ('Kwilu', 'GUNGU',   'Gungu')
ON CONFLICT (province_name, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Imbongo','Imbongo'),('Kwenge','Kwenge'),('Kipuka','Kipuka')
) AS s(name,label)
WHERE t.province_name='Kwilu' AND t.name='BULUNGU'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Kipuku','Kipuku'),('Kalanganda','Kalanganda'),('Idiofa_Musanga','Idiofa Musanga')
) AS s(name,label)
WHERE t.province_name='Kwilu' AND t.name='IDIOFA'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Mungindu','Mungindu'),('Kilamba','Kilamba'),('Mudikalunga','Mudikalunga'),('Gungu','Gungu'),('Lukamba','Lukamba')
) AS s(name,label)
WHERE t.province_name='Kwilu' AND t.name='GUNGU'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

-- Kasaï
INSERT INTO territoires (province_name, name, label) VALUES
  ('Kasaï', 'TSHIKAPA', 'Tshikapa'),
  ('Kasaï', 'LWEBO',    'Lwebo'),
  ('Kasaï', 'MWEKA',    'Mweka')
ON CONFLICT (province_name, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Bakwaniambi','Bakwaniambi'),('Bapende','Bapende'),('Tshikapa','Tshikapa')
) AS s(name,label)
WHERE t.province_name='Kasaï' AND t.name='TSHIKAPA'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Kabambayi','Kabambayi'),('Lwebo_Wedi','Lwebo-Wedi')
) AS s(name,label)
WHERE t.province_name='Kasaï' AND t.name='LWEBO'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Bakuba','Bakuba')
) AS s(name,label)
WHERE t.province_name='Kasaï' AND t.name='MWEKA'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

-- Kasaï Central
INSERT INTO territoires (province_name, name, label) VALUES
  ('Kasaï Central', 'DEMBA',  'Demba'),
  ('Kasaï Central', 'LUIZA',  'Luiza'),
  ('Kasaï Central', 'DIBAYA', 'Dibaya')
ON CONFLICT (province_name, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Tshibote','Tshibote'),('Diofwa','Diofwa'),('Tshibungu','Tshibungu'),
  ('Lusonge','Lusonge'),('Benamamba','Benamamba'),('Lombelo','Lombelo'),('Mwanzangoma','Mwanzangoma')
) AS s(name,label)
WHERE t.province_name='Kasaï Central' AND t.name='DEMBA'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Lusanza','Lusanza'),('Lueta','Lueta'),('Kalunga','Kalunga')
) AS s(name,label)
WHERE t.province_name='Kasaï Central' AND t.name='LUIZA'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

INSERT INTO secteurs (territoire_id, name, label)
SELECT t.id, s.name, s.label FROM territoires t
CROSS JOIN (VALUES
  ('Tshishilu','Tshishilu'),('Dibatayi','Dibatayi'),('Dibanda','Dibanda')
) AS s(name,label)
WHERE t.province_name='Kasaï Central' AND t.name='DIBAYA'
ON CONFLICT (territoire_id, name) DO UPDATE SET label = EXCLUDED.label;

-- Catégorisation automatique des missions existantes
UPDATE missions_ref SET category = 'atelier'
  WHERE nature ILIKE 'Atelier%'
     OR nature ILIKE 'Participation%ateliers%'
     OR nature ILIKE 'Participation%revue%'
     OR nature ILIKE 'Formation%'
     OR nature ILIKE 'Examen%validation%'
     OR nature ILIKE 'Briefing%'
     OR nature ILIKE 'Restitution%';

UPDATE missions_ref SET category = 'activite'
  WHERE nature ILIKE 'Mise en œuvre%';

-- Le reste garde 'implementation' (colonnes déjà à DEFAULT 'implementation')

-- Super-admin : promouvoir rnse
UPDATE users SET role = 'super_admin' WHERE email = 'rnse@pnda.cd';

-- Politiques supplémentaires (ignorées si déjà créées)
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
  IF NOT EXISTS (SELECT 1 FROM pg_policies WHERE tablename='reports' AND policyname='reports_update') THEN
    CREATE POLICY "reports_update" ON reports FOR UPDATE TO anon USING (true);
  END IF;
END $$;
