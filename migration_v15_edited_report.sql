-- ════════════════════════════════════════════════════════════════════════════
-- Migration V15 — Traçabilité de l'édition d'un rapport par le comptable
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte : la fonction saveEditedReport() (index.html) permet au comptable
-- de corriger les montants d'un rapport déjà soumis (nuitées, taux, avance...)
-- et enregistre qui a fait la modification et quand. Ces colonnes étaient
-- référencées dans le code mais absentes de toute migration — même classe de
-- bug que celui corrigé par la migration V10 (colonne inexistante → erreur
-- Postgrest lors de l'UPDATE).
-- ════════════════════════════════════════════════════════════════════════════

ALTER TABLE reports ADD COLUMN IF NOT EXISTS edited_by TEXT        DEFAULT NULL;
ALTER TABLE reports ADD COLUMN IF NOT EXISTS edited_at TIMESTAMPTZ DEFAULT NULL;
-- edited_by : email/identifiant du comptable qui a modifié le rapport
-- edited_at : horodatage de la dernière modification

CREATE INDEX IF NOT EXISTS idx_reports_edited
  ON reports (province, edited_at)
  WHERE edited_at IS NOT NULL;

-- ── Vérification ─────────────────────────────────────────────────────────────
-- SELECT id, edited_by, edited_at FROM reports WHERE edited_at IS NOT NULL ORDER BY edited_at DESC LIMIT 20;
