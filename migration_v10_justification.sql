-- ════════════════════════════════════════════════════════════════════════════
-- Migration V10 — Clôture des justificatifs et versement du solde (20%)
-- ════════════════════════════════════════════════════════════════════════════
-- Contexte métier :
--   • Le missionnaire perçoit 80% du total en avance au départ de la mission.
--   • À la fin (dans les 48h), il remet les justificatifs : rapport de mission
--     + ordre de mission avisé.
--   • Le comptable enregistre la réception des justificatifs (justification_at).
--   • Après vérification, le comptable confirme le versement du solde de 20%.
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1. Réception des justificatifs ──────────────────────────────────────────
ALTER TABLE reports ADD COLUMN IF NOT EXISTS justification_at  TIMESTAMPTZ DEFAULT NULL;
ALTER TABLE reports ADD COLUMN IF NOT EXISTS justification_by  TEXT        DEFAULT NULL;
-- justification_at : horodatage de la réception des justificatifs par le comptable
-- justification_by : email/identifiant du comptable qui a enregistré la réception

-- ── 2. Versement du solde (20%) ──────────────────────────────────────────────
ALTER TABLE reports ADD COLUMN IF NOT EXISTS solde_paye_at     TIMESTAMPTZ DEFAULT NULL;
ALTER TABLE reports ADD COLUMN IF NOT EXISTS solde_paye_by     TEXT        DEFAULT NULL;
-- solde_paye_at : horodatage du versement du solde 20% au missionnaire
-- solde_paye_by : email/identifiant du comptable qui a effectué le versement

-- ── 3. Index pour les requêtes de suivi ─────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_reports_justification
  ON reports (province, justification_at)
  WHERE justification_at IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_reports_solde_paye
  ON reports (province, solde_paye_at)
  WHERE solde_paye_at IS NOT NULL;

-- ── 4. Vue de vérification rapide ────────────────────────────────────────────
-- Rapports réalisés sans justificatifs (en retard ou en attente) :
-- SELECT id, province, rapporteur, statut, saved_at,
--        justification_at, solde_paye_at
-- FROM reports
-- WHERE statut = 'realisee'
-- ORDER BY saved_at DESC LIMIT 50;
