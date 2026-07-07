-- Migration v18 : corriger les rapports dont la colonne "province" ne
-- correspond plus exactement (casse/espaces) a un nom de la table `provinces`.
--
-- Contexte : saisie.html a recemment transforme le champ "Province" en saisie
-- libre (texte). Si un missionnaire tape "kwilu" ou "Kwilu " au lieu de
-- "Kwilu", le rapport est enregistre avec une valeur qui ne correspond plus
-- exactement a `provinces.name`. Or le tableau de bord du comptable filtre
-- les rapports avec une comparaison exacte (`province IN (...)`), donc ces
-- rapports deviennent invisibles pour le comptable de la province concernee.
--
-- Cette migration ne fait rien "en aveugle" : elle commence par un SELECT de
-- diagnostic, puis propose un UPDATE de correction a executer seulement apres
-- avoir verifie le resultat du diagnostic.

-- ── 1. Diagnostic : rapports dont la province ne matche aucune province
--       connue de maniere exacte, mais dont on peut deviner la bonne valeur
--       par comparaison insensible a la casse/aux espaces ────────────────────
SELECT
  r.id,
  r.province        AS province_enregistree,
  r.province_label  AS province_label_enregistree,
  p.name            AS province_correcte,
  p.label           AS province_label_correct,
  p.code            AS province_code_correct,
  r.rapporteur,
  r.trimestre,
  r.annee,
  r.saved_at
FROM reports r
LEFT JOIN provinces p
  ON lower(trim(p.name)) = lower(trim(r.province))
WHERE r.province IS NOT NULL
  AND NOT EXISTS (SELECT 1 FROM provinces p2 WHERE p2.name = r.province)
ORDER BY r.saved_at DESC;

-- ── 2. Correction : une fois le SELECT ci-dessus verifie, executer ce bloc
--       pour realigner "province" / "province_label" / "province_code" sur
--       les valeurs canoniques de la table `provinces` (uniquement pour les
--       lignes ou une correspondance insensible a la casse a ete trouvee) ───
-- UPDATE reports r
-- SET province       = p.name,
--     province_label  = p.label,
--     province_code   = p.code
-- FROM provinces p
-- WHERE lower(trim(p.name)) = lower(trim(r.province))
--   AND NOT EXISTS (SELECT 1 FROM provinces p2 WHERE p2.name = r.province);

-- Verification apres correction :
-- SELECT id, province, province_label, province_code FROM reports ORDER BY saved_at DESC LIMIT 20;
