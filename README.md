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

| URL | Fichier | Rôles autorisés |
|---|---|---|
| `/` | `index.html` | province, comptable |
| `/admin` | `admin.html` | admin, comptable |
| `/superadmin` | `superadmin.html` | super_admin |

---

## Rôles utilisateurs

| Rôle | Accès |
|---|---|
| `province` | Saisie et suivi de ses propres missions |
| `comptable` | Complétion financière, activation, clôture des missions de sa province |
| `admin` | Tableau de bord consolidé toutes provinces, export Excel |
| `super_admin` | Gestion des utilisateurs, rôles, référentiels et provinces (RNSE) |

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
supabase_migration_v2.sql               ← table missions normalisée
migration_v2_columns.sql                ← colonnes supplémentaires
migration_v3_territoires_secteurs.sql   ← territoires et secteurs
migration_v4_is_read.sql                ← marqueur de lecture
migration_v5_statut_roles.sql           ← statuts et rôle comptable
migration_v6_retournee_workflow.sql     ← workflow retour comptable
migration_v7_compte_uncp_comptable.sql  ← comptes UNCP
migration_v8_liquidation_sequences.sql  ← séquences de liquidation
migration_v9_province_comptables.sql    ← comptables par province
```

### 5. Autoriser le domaine dans Supabase

Dans **Authentication → URL Configuration**, ajouter :

- **Site URL** : `http://localhost:3000` (développement) ou `https://kisungu.github.io` (production)
- **Allowed Redirect URLs** : les mêmes

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
├── index.html                  # Formulaire de saisie (province / comptable)
├── admin.html                  # Tableau de bord admin
├── superadmin.html             # Gestion utilisateurs et référentiels (RNSE)
├── supabase_config.js          # URL et clé anon Supabase (à configurer)
├── server.js                   # Serveur Express pour le développement local
├── package.json
│
├── supabase_db.sql                         # Schéma initial
├── supabase_migration_v2.sql               # Missions normalisées
├── migration_v2_columns.sql
├── migration_v3_territoires_secteurs.sql
├── migration_v4_is_read.sql
├── migration_v5_statut_roles.sql
├── migration_v6_retournee_workflow.sql
├── migration_v7_compte_uncp_comptable.sql
├── migration_v8_liquidation_sequences.sql
└── migration_v9_province_comptables.sql
```

---

## Dépôt GitHub

https://github.com/KISUNGU/suivi-missions.git

## Licence

Projet `UNLICENSED` — usage interne PNDA-SE.