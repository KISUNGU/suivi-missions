# Analyse du projet PNDA-SE (Suivi des missions)

*État au 22 juillet 2026 — analysé à partir du code du dépôt, de l'historique Git et de la base Supabase en production.*

## 1. Vue d'ensemble

PNDA-SE est une application web de suivi des missions (frais de mission, justificatifs, liquidation) pour le Programme National de Développement Agricole. Elle est entièrement statique côté client (HTML/JS/Bootstrap, sans framework ni bundler), déployée sur GitHub Pages, avec Supabase comme backend (Postgres + Auth + Row Level Security). Un petit serveur Express (`server.js`, 19 lignes) sert uniquement au développement local.

Quatre interfaces distinctes, sans code partagé entre elles (chaque fichier réimplémente sa propre logique) :

- `saisie.html` (1 586 lignes) — chef de mission / province : saisie des missions, un formulaire complet par missionnaire, gestion de groupes, répartition multi-lieux.
- `index.html` (3 507 lignes) — comptable / super_admin (RNSE) : tableau de bord, validation, avis sur justificatifs, bordereaux, export Excel.
- `admin.html` (2 398 lignes) — comptable (RAF) : variante du tableau de bord, sans les fonctions avis/bordereau/justification.
- `superadmin.html` (2 243 lignes) — gestion des comptes, mots de passe, structure de référence (provinces/territoires).

Total : **~9 750 lignes** de HTML/JS applicatif, **0 test automatisé**, **28 fichiers de migration SQL** (numérotés v2 à v28, avec un doublon "v10" et un "v12" manquant).

## 2. Modèle de données et rôles

Table centrale unique : `reports` (64 lignes actuellement), avec un tableau JSONB `missions` contenant une entrée par missionnaire. Chaque mission porte ses propres dates, nuitées, taux, frais de voyage, détail des "autres frais à justifier", et depuis peu un `perdiem_breakdown` pour les missions à lieux multiples et taux différents.

Rôles (table `profiles`, 16 comptes) :

| Rôle | Effectif | Portée |
|---|---|---|
| `province` | 6 | chef de mission, scoped à une province |
| `comptable` | 4 | un par province (dont UNCP) |
| `admin` | 4 | RAF, portée nationale |
| `super_admin` | 2 | RNSE, portée nationale |

À ce stade, **seule la province UN (UNCP) a une activité réelle** : les 64 rapports en base sont tous rattachés à UN (51 en attente, 13 réalisées). Kwilu, Kasaï et Kasaï Central n'ont pas encore de données de mission — le circuit multi-provinces n'a donc pas encore été éprouvé en conditions réelles.

## 3. Sécurité (RLS) — état actuel

Le contrôle d'accès repose sur `has_province_access(province)` et des policies dédiées par opération (`reports_select_auth`, `_insert_auth`, `_update_auth`, `_delete_auth`). L'audit sécurité Supabase ne remonte que des avertissements attendus pour ce type d'app (exposition GraphQL des tables de référence, fonctions `SECURITY DEFINER` volontairement publiques pour les RPC de séquence) — rien de critique actuellement. Un point simple à corriger : la protection "mot de passe compromis" (HaveIBeenPwned) est désactivée dans Supabase Auth, à activer en un clic dans Authentication → Policies.

## 4. Incidents traités récemment

Le projet a connu une phase de correctifs intensive ces derniers jours (47 commits sur les 7 derniers jours, 184 au total, tous par le même auteur — signe d'itérations très rapprochées, y compris via plusieurs outils IA différents d'après les messages de commit : "claude", "Deepseek", etc.) :

1. **Fuite de données RAF → comptable Kwilu** : un rapport du RAF apparaissait chez le comptable Kwilu. Cause racine : `has_province_access()` acceptait toute province absente de la table `provinces` (contournement pensé pour les provinces inconnues, mais insensible à la casse/espaces). Corrigé (migrations v22-v23) + verrouillage de l'option "Autre" province côté saisie pour les comptes à périmètre restreint.
2. **Suppression impossible côté UNCP** : la policy `DELETE` ne tenait pas compte des rapports RAF/RNSE visibles par UNCP. Corrigé (migration v22).
3. **60 rapports saisis, 11 visibles** : 49 rapports UNCP avaient leur champ "province" rempli en texte libre avec le nom de la ville de destination (ex. "KANANGA", "Tshikapa & Kikwit") au lieu de "UN". Corrigé en base (province → UN, ville reportée dans le territoire) et au niveau du formulaire (option "Autre" retirée pour les comptes scoped) — cette même correction a ensuite été redocumentée côté dépôt sous la migration v28.
4. **Filtre de période "muet"** : le filtre comparait la date de *saisie* du rapport au lieu des dates réelles de la mission — déjà corrigé et vérifié contre les données réelles.

Ces quatre incidents partagent une cause commune : **la saisie libre de la province** a été, à plusieurs reprises, le point d'entrée de données mal formées qui cassent ensuite le filtrage par province utilisé partout dans le tableau de bord. C'est maintenant verrouillé, mais c'est le risque structurel le plus significatif du projet (voir §6).

## 5. Fonctionnalités construites récemment

- **Multi-lieux avec taux différents** : un missionnaire peut désormais avoir plusieurs étapes (lieu, dates, taux) dans une même mission ; nuitées calculées automatiquement par étape, total agrégé. Répercuté dans le bordereau et la "Justification de la mission".
- **Avis sur justificatifs** : calcul automatique favorable/sous réserve, détail des frais cochables/corrigibles, confirmation explicite requise si l'agent doit rembourser.
- **Justification de la mission** : document détaillé prévu/payé/réalisé/écart par ligne, conforme au modèle papier fourni.
- **Frais du jour de voyage** : nouvelle composante du frais de mission, indépendante du barème de nuitée.
- **Litige** : logique financière corrigée (n'affiche plus de litige une fois le solde payé).

## 6. Dette technique et risques

- **Duplication totale entre `index.html` et `admin.html`** : les mêmes fonctions (`getFilteredReports`, `reportMatchesPeriod`, `buildProvincePanel`, etc.) existent en double, avec un risque réel de divergence — déjà observé (le correctif du filtre de période aurait pu n'être appliqué qu'à un seul des deux fichiers). Aucune bibliothèque partagée.
- **Aucun test automatisé** : chaque correctif est vérifié manuellement (ou par un script Node ad hoc). Sur un projet à ce rythme de commits, le risque de régression silencieuse est réel.
- **28 migrations SQL non consolidées**, dont un doublon "v10" et un "v12" manquant — la lecture de l'historique RLS devient difficile ; un `supabase_db.sql` consolidé existe mais son niveau de synchronisation avec l'état réel de la base n'est pas garanti.
- **Rythme de commits très élevé avec des messages peu descriptifs** ("montant", "change", "tb compte", "commiter") — bon pour l'itération rapide, mais rend l'archéologie de bugs plus lente (comme vécu pour le filtre de période).
- **Édition concurrente du même dépôt par plusieurs outils/agents** (au moins Claude et un autre assistant nommé "Deepseek" dans les commits) : deux corrections que j'avais appliquées plus tôt dans cette session (restriction de l'option "Autre", `has_province_access`) ont été écrasées par des commits ultérieurs avant d'être réappliquées — signe qu'il n'y a pas de coordination entre les sessions de travail sur ce fichier.
- **Un seul environnement testé en réel** (province UN) : Kwilu/Kasaï/Kasaï Central n'ont pas encore fait tourner le workflow complet (saisie → activation → justificatifs → solde), donc des bugs spécifiques à ces comptes peuvent rester non détectés.

## 7. Recommandations, par priorité

1. **Activer la protection mot de passe compromis** dans Supabase Auth (2 minutes, aucun risque).
2. **Faire tester le circuit complet par un compte Kwilu ou Kasaï** avant de considérer le multi-provinces comme validé.
3. **Renommer/consolider les migrations** (fusionner le doublon v10, documenter le v12 manquant) pour que l'historique RLS reste lisible.
4. **Éviter les éditions concurrentes du même fichier par plusieurs outils IA en parallèle** sans se resynchroniser entre-temps — c'est la cause directe des régressions "Autre" et `has_province_access` observées cette semaine.
5. **Extraire la logique dupliquée** (`getFilteredReports`, `reportMatchesPeriod`, `buildProvincePanel`…) dans un fichier JS partagé chargé par `index.html` et `admin.html`, pour arrêter la divergence silencieuse entre les deux tableaux de bord.
6. **Ajouter quelques tests de non-régression légers** (au minimum sur les fonctions de calcul financier et de filtrage — pas besoin d'un framework lourd, un script Node comme ceux utilisés pour vérifier cette session suffirait, versionné dans le dépôt).

## 8. En résumé

Le projet est fonctionnellement riche et évolue vite, mais cette vitesse a un coût direct : la plupart des bugs traités cette semaine viennent de champs texte libre non contrôlés (province) ou de logique dupliquée non synchronisée entre fichiers — deux problèmes structurels plutôt que ponctuels. Le socle (RLS, rôles, calculs financiers) est aujourd'hui cohérent et vérifié sur les données réelles, mais le rythme d'édition concurrente par plusieurs outils IA sans coordination est le principal facteur de risque à court terme.
