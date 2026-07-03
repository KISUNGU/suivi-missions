-- ════════════════════════════════════════════════════════════════════════════
-- Migration V14 — Correction SECURITY DEFINER sur les fonctions de séquence
-- ════════════════════════════════════════════════════════════════════════════
-- Problème : next_fld_seq et next_liquidation_seq n'avaient pas SECURITY
-- DEFINER, donc elles s'exécutaient sous le rôle "authenticated" qui n'a pas
-- de politique RLS sur les tables de séquences (uniquement "anon").
-- Résultat : erreur 403 lors de l'activation ou de la clôture des justifs.
--
-- Correction : SECURITY DEFINER (comme next_report_seq dans migration_v13)
-- fait tourner la fonction avec les droits de son propriétaire (postgres),
-- contournant ainsi le RLS des tables de séquences.
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1. next_fld_seq ──────────────────────────────────────────────────────────
CREATE OR REPLACE FUNCTION public.next_fld_seq(p_year INT, p_month INT)
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

REVOKE EXECUTE ON FUNCTION public.next_fld_seq(INT, INT) FROM PUBLIC;
GRANT  EXECUTE ON FUNCTION public.next_fld_seq(INT, INT) TO anon;
GRANT  EXECUTE ON FUNCTION public.next_fld_seq(INT, INT) TO authenticated;

-- ── 2. next_liquidation_seq ──────────────────────────────────────────────────
CREATE OR REPLACE FUNCTION public.next_liquidation_seq(p_province TEXT, p_year INT)
RETURNS INT LANGUAGE plpgsql SECURITY DEFINER SET search_path = public AS $$
DECLARE
  v_seq INT;
BEGIN
  INSERT INTO liquidation_sequences(province, year, last_seq)
  VALUES (p_province, p_year, 1)
  ON CONFLICT (province, year) DO UPDATE
    SET last_seq = liquidation_sequences.last_seq + 1
  RETURNING last_seq INTO v_seq;
  RETURN v_seq;
END;
$$;

REVOKE EXECUTE ON FUNCTION public.next_liquidation_seq(TEXT, INT) FROM PUBLIC;
GRANT  EXECUTE ON FUNCTION public.next_liquidation_seq(TEXT, INT) TO anon;
GRANT  EXECUTE ON FUNCTION public.next_liquidation_seq(TEXT, INT) TO authenticated;

-- ── Vérification ─────────────────────────────────────────────────────────────
-- SELECT proname, prosecdef FROM pg_proc
--   WHERE proname IN ('next_fld_seq','next_liquidation_seq','next_report_seq');
-- prosecdef = true  →  SECURITY DEFINER bien actif
