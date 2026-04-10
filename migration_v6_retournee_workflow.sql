-- ════════════════════════════════════════════════════════════════════════════
-- Migration V6 — Workflow de retour missionnaire / comptable
-- ════════════════════════════════════════════════════════════════════════════
-- À exécuter après migration_v5_statut_roles.sql dans l'éditeur SQL Supabase.
-- Cette migration formalise le statut "retournee" et ajoute la traçabilité
-- du retour du dossier par le comptable vers le missionnaire.
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1. Colonnes de traçabilité pour le retour au missionnaire ───────────────
ALTER TABLE reports ADD COLUMN IF NOT EXISTS returned_by TEXT DEFAULT NULL;
ALTER TABLE reports ADD COLUMN IF NOT EXISTS returned_at TIMESTAMPTZ DEFAULT NULL;

-- ── 2. Normalisation des statuts existants ──────────────────────────────────
-- Certains anciens jeux de données peuvent encore contenir un statut vide
-- ou hérité d'une ancienne version de la table.
UPDATE reports
SET statut = 'en_attente'
WHERE statut IS NULL
   OR btrim(statut) = ''
   OR statut = 'saisie';

-- ── 3. Contraindre les valeurs autorisées du workflow ───────────────────────
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM pg_constraint
    WHERE conname = 'reports_statut_workflow_check'
  ) THEN
    ALTER TABLE reports
      ADD CONSTRAINT reports_statut_workflow_check
      CHECK (statut IN ('en_attente', 'retournee', 'activee', 'realisee'));
  END IF;
END $$;

-- ── 4. Index dédié pour les missions retournées ─────────────────────────────
CREATE INDEX IF NOT EXISTS idx_reports_retournee
  ON reports (province, returned_at DESC)
  WHERE statut = 'retournee';

-- ── 5. Vérification rapide ───────────────────────────────────────────────────
-- SELECT statut, COUNT(*) FROM reports GROUP BY statut ORDER BY statut;
-- SELECT id, province, rapporteur, statut, returned_by, returned_at FROM reports ORDER BY saved_at DESC LIMIT 20;
