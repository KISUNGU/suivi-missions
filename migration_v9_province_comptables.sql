-- Migration v9 : aligner les comptes provinciaux sur le role comptable
-- Objectif : donner a Kwilu, Kasai et Kasai Central le meme tableau de bord
-- que le compte UNCP, avec workflow comptable + acces admin.

UPDATE users
SET role = 'comptable'
WHERE email IN (
  'compte_kwilu@pnda.cd',
  'compte_kasaic@pnda.cd',
  'compte_kasai@pnda.cd'
);

-- Verification conseillee
-- SELECT email, role, provinces
-- FROM users
-- WHERE email IN (
--   'compte_kwilu@pnda.cd',
--   'compte_kasaic@pnda.cd',
--   'compte_kasai@pnda.cd',
--   'compte_uncp@pnda.cd'
-- );