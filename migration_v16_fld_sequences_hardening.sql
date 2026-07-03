-- ════════════════════════════════════════════════════════════════════════════
-- Migration V16 — Durcissement RLS de fld_sequences (anon → authenticated)
-- ════════════════════════════════════════════════════════════════════════════
-- Problème : la migration V11 a créé fld_sequences avec une policy RLS et des
-- GRANTs ouverts au rôle "anon" (clé publique, sans authentification). C'est
-- une régression par rapport au verrouillage général effectué en V10 (toutes
-- les tables métier ont été restreintes à "authenticated"). Contrairement à
-- report_sequences et liquidation_sequences (déjà en "authenticated"),
-- fld_sequences restait accessible en lecture/écriture directe à n'importe
-- qui possédant la clé anonyme du projet, sans connexion.
--
-- Remarque : la fonction next_fld_seq() est SECURITY DEFINER (migration V14),
-- donc l'appel RPC utilisé par l'application n'est pas affecté par ce
-- changement — seul l'accès direct à la table est concerné.
-- ════════════════════════════════════════════════════════════════════════════

DROP POLICY IF EXISTS "fld_sequences_all" ON fld_sequences;

DO $$ BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_policies
    WHERE tablename = 'fld_sequences' AND policyname = 'fld_sequences_auth'
  ) THEN
    CREATE POLICY "fld_sequences_auth" ON fld_sequences
      FOR ALL TO authenticated USING (true) WITH CHECK (true);
  END IF;
END $$;

REVOKE ALL ON fld_sequences FROM anon;
REVOKE ALL ON fld_sequences FROM PUBLIC;
GRANT SELECT, INSERT, UPDATE ON fld_sequences TO authenticated;

-- ── 2. Les fonctions next_fld_seq / next_liquidation_seq n'ont besoin d'être
-- appelées que par des utilisateurs connectés (comptable/admin) ; l'accès
-- anon (hérité de la migration V11/V14) est retiré par cohérence avec
-- next_report_seq (V13) qui n'a jamais été ouvert à anon.
REVOKE EXECUTE ON FUNCTION public.next_fld_seq(INT, INT) FROM anon;
REVOKE EXECUTE ON FUNCTION public.next_liquidation_seq(TEXT, INT) FROM anon;

-- ── Vérification ─────────────────────────────────────────────────────────────
-- SELECT grantee, privilege_type FROM information_schema.role_table_grants WHERE table_name = 'fld_sequences';
-- SELECT policyname, roles::text FROM pg_policies WHERE tablename = 'fld_sequences';
