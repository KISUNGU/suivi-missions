-- ════════════════════════════════════════════════════════════════════════════
-- Migration V8 — Numérotation dédiée des bordereaux de liquidation
-- ════════════════════════════════════════════════════════════════════════════
-- À exécuter après les migrations précédentes dans l'éditeur SQL Supabase.
-- Garantit un numéro séquentiel strict par province et par année pour les
-- bordereaux de liquidation, indépendant du numéro de rapport.
-- ════════════════════════════════════════════════════════════════════════════

ALTER TABLE reports ADD COLUMN IF NOT EXISTS liquidation_seq_number INT;
ALTER TABLE reports ADD COLUMN IF NOT EXISTS liquidation_number TEXT;
ALTER TABLE reports ADD COLUMN IF NOT EXISTS liquidation_printed_at TIMESTAMPTZ;
ALTER TABLE reports ADD COLUMN IF NOT EXISTS liquidation_printed_by TEXT;

CREATE TABLE IF NOT EXISTS liquidation_sequences (
  province TEXT NOT NULL,
  year INT NOT NULL,
  last_seq INT NOT NULL DEFAULT 0,
  PRIMARY KEY (province, year)
);

ALTER TABLE liquidation_sequences ENABLE ROW LEVEL SECURITY;

DO $$ BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_policies
    WHERE tablename = 'liquidation_sequences' AND policyname = 'liquidation_sequences_all'
  ) THEN
    CREATE POLICY "liquidation_sequences_all" ON liquidation_sequences
      FOR ALL TO anon USING (true) WITH CHECK (true);
  END IF;
END $$;

CREATE OR REPLACE FUNCTION next_liquidation_seq(p_province TEXT, p_year INT)
RETURNS INT LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
DECLARE v_seq INT;
BEGIN
  INSERT INTO liquidation_sequences(province, year, last_seq)
  VALUES (p_province, p_year, 1)
  ON CONFLICT (province, year) DO UPDATE
    SET last_seq = liquidation_sequences.last_seq + 1
  RETURNING last_seq INTO v_seq;

  RETURN v_seq;
END;
$$;

GRANT EXECUTE ON FUNCTION next_liquidation_seq(TEXT, INT) TO anon;
GRANT EXECUTE ON FUNCTION next_liquidation_seq(TEXT, INT) TO authenticated;

CREATE UNIQUE INDEX IF NOT EXISTS idx_reports_liquidation_number
  ON reports (liquidation_number)
  WHERE liquidation_number IS NOT NULL;

-- Vérification
-- SELECT province, year, last_seq FROM liquidation_sequences ORDER BY year DESC, province;
-- SELECT id, liquidation_seq_number, liquidation_number, liquidation_printed_at FROM reports ORDER BY saved_at DESC LIMIT 20;