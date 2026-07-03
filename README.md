# PNDA-SE — Suivi des missions

[![Version](https://img.shields.io/badge/version-1.0.0-1d4c34)](https://github.com/KISUNGU/suivi-missions)
[![Web App](https://img.shields.io/badge/plateforme-Web-0f6cbd)](https://kisungu.github.io/suivi-missions/)
[![Supabase](https://img.shields.io/badge/base_de_données-Supabase-3ecf8e)](https://supabase.com)
[![GitHub Pages](https://img.shields.io/badge/déploiement-GitHub_Pages-222)](https://kisungu.github.io/suivi-missions/)

Application web du PNDA-SE pour la saisie, le suivi et le reporting des missions de terrain.

---

## Aperçu

L'application permet de gérer le cycle de vie complet d'une mission :

- saisie du rapport de mission par le missionnaire (province) ;
- validation et complétion financière par le comptable ;
- tableau de bord consolidé pour les administrateurs avec export Excel ;
- gestion des utilisateurs, rôles et référentiels par le super-administrateur (RNSE).

**Accès en ligne (GitHub Pages) :** https://kisungu.github.io/suivi-missions/

---

## Architecture

L'application est une **web app statique** :

- Les pages HTML chargent Bootstrap, Supabase.js et SheetJS depuis des CDN.
- Toutes les données sont stockées dans **Supabase** (PostgreSQL cloud).
- L'export Excel est généré directement dans le navigateur (SheetJS).
- Un serveur Express minimal (`server.js`) est fourni pour le développement local.

```
Navigateur ──► index.html / admin.html / superadmin.html
                    │
                    ▼
              Supabase (cloud)
              PostgreSQL + Auth
```

---

## Pages

| URL | Fichier | Rôles autorisés | Contenu |
|---|---|---|---|
| `/saisie.html` | `saisie.html` | province, comptable, super_admin | Formulaire de saisie des missions + suivi de ses propres rapports |
| `/` | `index.html` | admin, comptable, super_admin | Tableau de bord consolidé (suivi, workflow, export Excel) |
| `/admin` | `admin.html` | admin, comptable, super_admin | Identique à `index.html` (variante) |
| `/superadmin` | `superadmin.html` | super_admin | Gestion utilisateurs, provinces, référentiels (RNSE) |

> `index.html` et `admin.html` restent deux tableaux de bord quasiment
> identiques (candidats à une fusion future). `saisie.html` est la page dédiée
> au rôle `province` : sélection d'une mission de référence (ou saisie libre),
> territoire/secteur, dates, participants/nuitées/taux per diem avec calcul
> automatique des montants, et suivi du statut de ses rapports (avec
> correction/renvoi pour un rapport « retourné » par le comptable).

---

## Rôles utilisateurs

| Rôle | Accès |
|---|---|
| `province` | Saisie de ses missions et suivi de ses propres rapports (`saisie.html`) |
| `comptable` | Complétion financière, activation, clôture des missions de sa province |
| `admin` | Tableau de bord consolidé toutes provinces, export Excel |
| `super_admin` | Gestion des utilisateurs, rôles, référentiels et provinces (RNSE) |

Le rôle et les provinces autorisées de chaque compte sont stockés dans la
table `profiles` (liée à Supabase Auth). Une province vide/`NULL` signifie
« accès à toutes les provinces » (cas des comptes `admin` et de certains
comptes `comptable` historiques rattachés à l'UNCP).

---

## Workflow métier

```
[province]   Crée la mission → statut : en_attente
[comptable]  Complète la partie financière → statut : activee
[comptable]  Retourne si correction nécessaire → statut : retournee
[province]   Corrige et renvoie → statut : en_attente
[comptable]  Clôture après réalisation → statut : realisee
```

---

## Prérequis

- [Node.js](https://nodejs.org) 18 ou plus récent
- Un compte [Supabase](https://supabase.com) (gratuit)

---

## Installation depuis zéro

### 1. Cloner le dépôt

```bash
git clone https://github.com/KISUNGU/suivi-missions.git
cd suivi-missions
```

### 2. Installer les dépendances

```bash
npm install
```

### 3. Configurer Supabase

Créer un projet sur https://supabase.com, puis aller dans **Project Settings → API** et copier le **Project URL** et la **clé anon publique**.

Ouvrir `supabase_config.js` et remplacer les valeurs :

```js
const SUPABASE_URL      = 'https://VOTRE_PROJET.supabase.co';
const SUPABASE_ANON_KEY = 'VOTRE_CLE_ANON';
```

### 4. Initialiser la base de données

Dans l'éditeur SQL de votre projet Supabase (**SQL Editor → New query**), exécuter les fichiers dans cet ordre :

```
supabase_db.sql                         ← schéma de base
supabase_migration_v2.sql               ← non utilisée en prod, conservée pour l'historique (voir note dans le fichier)
migration_v2_columns.sql                ← colonnes supplémentaires
migration_v3_territoires_secteurs.sql   ← territoires et secteurs
migration_v4_is_read.sql                ← marqueur de lecture
migration_v5_statut_roles.sql           ← statuts et rôle comptable
migration_v6_retournee_workflow.sql     ← workflow retour comptable
migration_v7_compte_uncp_comptable.sql  ← comptes UNCP
migration_v8_liquidation_sequences.sql  ← séquences de liquidation (non branchées côté UI)
migration_v9_province_comptables.sql    ← comptables par province
migration_v10_supabase_auth.sql         ← authentification Supabase Auth + verrouillage RLS (obligatoire)
migration_v13_report_sequences.sql      ← numérotation des rapports pour saisie.html (obligatoire)
```

> `migration_v10` est indispensable : sans elle, aucun utilisateur ne peut se
> connecter (aucun compte Supabase Auth n'existe) et les données restent
> accessibles publiquement via la clé anon.

### 5. Déployer la fonction Edge `admin-users`

La création, l'édition du mot de passe et la suppression de comptes depuis
l'espace RNSE (`superadmin.html`) nécessitent la clé `service_role`, qui ne
doit jamais être exposée au navigateur. Cette logique vit dans une fonction
Edge (`supabase/functions/admin-users`) :

```bash
supabase login
supabase link --project-ref VOTRE_PROJET
supabase functions deploy admin-users
```

### 6. Configurer les URL dans Supabase Auth

Dans **Authentication → URL Configuration**, ajouter :

- **Site URL** : `http://localhost:3000` (développement) ou `https://kisungu.github.io` (production)
- **Allowed Redirect URLs** : les mêmes

Optionnel mais recommandé : dans **Authentication → Policies**, activer
« Leaked password protection » (nécessite un plan Supabase Pro).

---

## Lancer en développement local

```bash
npm run dev
```

Ouvrir **http://localhost:3000** dans le navigateur.

Le serveur redémarre automatiquement si `server.js` est modifié. Les fichiers HTML sont rechargés avec F5.

```bash
npm start   # version sans rechargement automatique
```

---

## Déploiement sur GitHub Pages

L'application peut être hébergée gratuitement sur GitHub Pages (aucun serveur requis, tout s'exécute dans le navigateur).

1. Aller sur **https://github.com/KISUNGU/suivi-missions/settings/pages**
2. **Source** → `Deploy from a branch`
3. **Branch** → `main` / `/ (root)`
4. Cliquer **Save**

L'app sera disponible à **https://kisungu.github.io/suivi-missions/** après quelques minutes.

> Penser à mettre à jour le **Site URL** dans Supabase avec ce domaine.

---

## Structure des fichiers

```
suivi-missions/
├── index.html                  # Tableau de bord (admin / comptable / super_admin)
├── admin.html                  # Tableau de bord (variante quasi identique)
├── superadmin.html             # Gestion utilisateurs et référentiels (RNSE)
├── saisie.html                 # Formulaire de saisie des missions (province / comptable / super_admin)
├── auth-shared.js              # Écran de connexion + changement de mot de passe, partagés par les 4 pages
├── supabase_config.js          # URL et clé anon Supabase (à configurer)
├── server.js                   # Serveur Express pour le développement local
├── package.json
│
├── supabase/functions/admin-users/index.ts # Fonction Edge (gestion des comptes, clé service_role)
│
├── supabase_db.sql                         # Schéma initial
├── supabase_migration_v2.sql               # Non utilisée en prod (historique)
├── migration_v2_columns.sql
├── migration_v3_territoires_secteurs.sql
├── migration_v4_is_read.sql
├── migration_v5_statut_roles.sql
├── migration_v6_retournee_workflow.sql
├── migration_v7_compte_uncp_comptable.sql
├── migration_v8_liquidation_sequences.sql  # Non branchée côté UI
├── migration_v9_province_comptables.sql
├── migration_v10_supabase_auth.sql         # Authentification + verrouillage RLS
└── migration_v13_report_sequences.sql      # Numérotation des rapports (saisie.html)
```

---

## Le formulaire de saisie (`saisie.html`)

Un compte `province` (ou `comptable`/`super_admin`) y saisit un rapport
composé d'une ou plusieurs missions :

- **Mission de référence** : choix dans le référentiel `missions_ref` de sa
  province (préremplit nature/objectif) ou saisie libre.
- **Territoire / secteur** : listes déroulantes en cascade à partir des tables
  `territoires`/`secteurs`.
- **Calcul automatique** : `montant_usd = nuitées × taux per diem`,
  `total_usd = montant_usd + autres frais`, `solde_usd = total_usd − avance`.
  Ces trois champs sont recalculés à la volée mais l'utilisateur reste libre
  de changer nuitées/taux/frais/avance — ce ne sont pas des règles métier
  officielles, juste une aide à la saisie à ajuster si besoin.
- **Numérotation** : `seq_number` attribué via la fonction `next_report_seq`
  (migration v13), un compteur par province.
- **Suivi** : liste de ses propres rapports avec statut ; un rapport
  « retourné » par le comptable peut être rouvert, corrigé et renvoyé
  (repasse en `en_attente`).

---

## Limites connues

- **`index.html` et `admin.html`** sont deux pages quasi identiques (candidates
  à une fusion).
- **Liquidation (migration_v8)** : les colonnes et la fonction
  `next_liquidation_seq` existent en base mais ne sont appelées par aucune
  page — fonctionnalité inachevée.
- **`saisie.html` est un premier jet fonctionnel** : le calcul du montant per
  diem (`nuitées × taux`) est une hypothèse simple à valider avec le métier ;
  ajustez la formule dans `recalcRow()`/`collectMissions()` si les règles
  réelles diffèrent (ex. prise en compte du nombre de participants).

---

## Dépôt GitHub

https://github.com/KISUNGU/suivi-missions.git

## Licence

Projet `UNLICENSED` — usage interne PNDA-SE.