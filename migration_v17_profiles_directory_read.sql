-- ════════════════════════════════════════════════════════════════════════════
-- Migration V17 — Lecture annuaire profiles pour admin/comptable
-- ════════════════════════════════════════════════════════════════════════════
-- Problème : la policy SELECT sur profiles (migration V10) ne laissait lire
-- que sa propre ligne (auth.uid() = id) ou le super_admin. Résultat : les
-- comptes comptable/admin ne pouvaient pas résoudre le "nom d'enregistrement"
-- (display_name) des utilisateurs province à partir des rapports (le champ
-- reports.meta.rapporteur est un instantané figé au moment de la saisie, qui
-- contient parfois le compte/email brut au lieu du nom réel).
--
-- Cette migration élargit la lecture de profiles (email, display_name, role)
-- aux rôles admin et comptable, en plus de sa propre ligne et du super_admin,
-- pour permettre la résolution automatique du nom du missionnaire dans le
-- tableau de bord et les exports. Aucune donnée sensible supplémentaire
-- n'est exposée (pas de mots de passe, profiles ne contient que email/role/
-- display_name/provinces).
-- ════════════════════════════════════════════════════════════════════════════

DROP POLICY IF EXISTS "profiles_select_own_or_admin" ON profiles;

CREATE POLICY "profiles_select_own_or_staff" ON profiles
  FOR SELECT TO authenticated
  USING (
    auth.uid() = id
    OR is_super_admin()
    OR current_profile_role() IN ('admin', 'comptable')
  );

-- ── Vérification ─────────────────────────────────────────────────────────────
-- SELECT policyname, roles::text, cmd, qual FROM pg_policies WHERE tablename = 'profiles';
