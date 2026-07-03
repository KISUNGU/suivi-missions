-- ════════════════════════════════════════════════════════════════════════════
-- Migration V11 — Numérotation FLD (Avis sur Justificatifs Provision)
-- ════════════════════════════════════════════════════════════════════════════
-- Format : FLD 043/06/2026  (séq / mois / année — séquence globale mensuelle)
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1. Colonne sur la table reports ─────────────────────────────────────────
ALTER TABLE reports ADD COLUMN IF NOT EXISTS fld_number TEXT DEFAULT NULL;

CREATE UNIQUE INDEX IF NOT EXISTS idx_reports_fld_number
  ON reports (fld_number)
  WHERE fld_number IS NOT NULL;

-- ── 2. Table de séquence FLD (globale, par mois/année) ─────────────────────
CREATE TABLE IF NOT EXISTS fld_sequences (
  year_month TEXT NOT NULL PRIMARY KEY,  -- ex : '2026-06'
  last_seq   INT  NOT NULL DEFAULT 0
);

ALTER TABLE fld_sequences ENABLE ROW LEVEL SECURITY;

DO $$ BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_policies
    WHERE tablename = 'fld_sequences' AND policyname = 'fld_sequences_all'
  ) THEN
    CREATE POLICY "fld_sequences_all" ON fld_sequences
      FOR ALL TO anon USING (true) WITH CHECK (true);
  END IF;
END $$;

-- ── 3. Fonction de séquence FLD ─────────────────────────────────────────────
CREATE OR REPLACE FUNCTION next_fld_seq(p_year INT, p_month INT)
RETURNS INT LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
DECLARE
  v_key TEXT;
  v_seq INT;
BEGIN
  v_key := p_year::TEXT || '-' || LPAD(p_month::TEXT, 2, '0');
  INSERT INTO fld_sequences(year_month, last_seq)
  VALUES (v_key, 1)
  ON CONFLICT (year_month) DO UPDATE
    SET last_seq = fld_sequences.last_seq + 1
  RETURNING last_seq INTO v_seq;
  RETURN v_seq;
END;
$$;

GRANT EXECUTE ON FUNCTION next_fld_seq(INT, INT) TO anon;
GRANT EXECUTE ON FUNCTION next_fld_seq(INT, INT) TO authenticated;

-- ── 4. Vérification ─────────────────────────────────────────────────────────
-- SELECT year_month, last_seq FROM fld_sequences ORDER BY year_month DESC;
-- SELECT id, fld_number, justification_at FROM reports WHERE fld_number IS NOT NULL ORDER BY justification_at DESC LIMIT 20;
