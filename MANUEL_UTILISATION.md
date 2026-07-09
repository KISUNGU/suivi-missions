# Manuel d'utilisation simplifié

Ce petit manuel explique comment utiliser l'application PNDA-SE pour trois profils :

1. Utilisateur ordinaire, c'est-à-dire le compte de saisie de mission
2. Comptable
3. Super-administrateur, c'est-à-dire le responsable technique / RNSE

L'objectif est de permettre à une personne non experte de savoir **où se connecter**, **avec quels identifiants**, et **quoi faire ensuite**.

---

## 1. Comment se connecter

Tous les utilisateurs commencent de la même façon :

- Ouvrir la bonne page de l'application selon le rôle.
- Saisir l'**adresse e-mail**.
- Saisir le **mot de passe**.
- Valider la connexion.

### Identifiants à utiliser

Dans cette application, les identifiants sont généralement :

- **Adresse e-mail** de l'utilisateur
- **Mot de passe personnel**

Ces identifiants sont fournis ou créés par l'administrateur du système. Si vous ne les avez pas, il faut les demander au responsable qui gère les comptes.

### Pages de connexion

- **Utilisateur ordinaire** : [saisie.html](saisie.html)
- **Comptable** : [index.html](index.html) ou [admin.html](admin.html)
- **Super-administrateur** : [superadmin.html](superadmin.html)

---

## 2. Utilisateur ordinaire

Ce rôle sert à **remplir une réquisition de frais de mission** et à suivre ses propres rapports.

### Identifiants

- Votre **adresse e-mail de connexion**
- Votre **mot de passe personnel**

### Ce que vous pouvez faire

- Créer un nouveau rapport de mission
- Ajouter un ou plusieurs missionnaires
- Renseigner les dates, les montants et les frais à justifier
- Envoyer le rapport
- Voir vos rapports déjà envoyés
- Corriger un rapport retourné par le comptable

### Procédure pas à pas

1. Ouvrez [saisie.html](saisie.html).
2. Connectez-vous avec votre e-mail et votre mot de passe.
3. Vérifiez les informations générales :
   - Province
   - Territoire, si nécessaire
   - Objet de la mission
   - Trimestre
   - Année
4. Dans la partie **Missionnaires**, cliquez sur **Ajouter un missionnaire**.
5. Remplissez pour chaque missionnaire :
   - Nom complet
   - Fonction / titre
   - Service / entité
   - Date de départ
   - Date de retour
   - Taux per diem
   - Avance versée
   - Frais à justifier, si besoin
6. Répétez l'opération pour tous les missionnaires concernés.
7. Cliquez sur **Soumettre le rapport**.
8. Après l'envoi, consultez la table **Mes rapports** pour voir le statut.

### Si le comptable retourne le rapport

Si un rapport est marqué comme **retourné**, cela veut dire qu'il manque quelque chose ou qu'une correction est demandée.

Dans ce cas :

1. Ouvrez le rapport retourné dans **Mes rapports**.
2. Cliquez sur **Corriger**.
3. Modifiez les informations demandées.
4. Renvoyez le rapport.

### Conseil simple

Si vous n'êtes pas sûr d'un montant ou d'une date, vérifiez d'abord le document de mission papier avant de soumettre. Un petit oubli peut bloquer tout le traitement.

---

## 3. Comptable

Ce rôle sert à **contrôler**, **compléter** et **valider la partie financière** des missions.

### Identifiants

- Votre **adresse e-mail de connexion**
- Votre **mot de passe personnel**

### Ce que vous pouvez faire

- Consulter les rapports reçus
- Vérifier les dates, les nuitées et les montants
- Compléter ou contrôler les informations financières
- Retourner un rapport pour correction si quelque chose manque
- Suivre l'état des missions jusqu'à leur clôture

### Procédure pas à pas

1. Ouvrez [index.html](index.html) ou [admin.html](admin.html).
2. Connectez-vous avec votre e-mail et votre mot de passe.
3. Ouvrez la liste des rapports disponibles.
4. Examinez chaque rapport avec attention :
   - Province
   - Objet de la mission
   - Missionnaires
   - Dates
   - Montants
5. Si les données sont correctes, poursuivez le traitement selon le workflow du tableau de bord.
6. Si une correction est nécessaire, utilisez l'action de **retour** pour demander la correction au missionnaire.
7. Une fois la mission terminée et les justificatifs vérifiés, passez le dossier au statut final prévu par le système.

### Ce qu'il faut vérifier en priorité

- Le nom du missionnaire est correct
- Les dates de départ et de retour sont cohérentes
- Le nombre de nuitées est plausible
- Les montants correspondent aux règles internes
- La province du rapport est bien la bonne

### Conseil simple

Si un rapport semble incomplet, ne le validez pas trop vite. Mieux vaut le retourner une fois que corriger un dossier déjà avancé.

### Si la liste des rapports est vide pour "SE_KWILU"

Si vous voyez un écran vide alors que des rapports existent en base, le problème vient le plus souvent du **rattachement de la province**, pas d'une base déconnectée.

- Le nom de province utilisé par l'application est le nom canonique enregistré en base, par exemple **Kwilu**.
- Un libellé de compte comme **SE_KWILU** est un identifiant d'organisation ou de service, mais ce n'est pas forcément le nom de province utilisé par les filtres.
- Si le compte est rattaché à une valeur différente de **Kwilu**, les rapports peuvent ne pas apparaître dans les suivis provinciaux.

Dans ce cas, il faut demander au super-administrateur de vérifier :

1. le rôle du compte ;
2. la ou les provinces autorisées ;
3. le nom exact de province enregistré dans la base.

---

## 4. Super-administrateur

Ce rôle sert à **gérer l'application**, **les comptes utilisateurs**, **les provinces** et **les référentiels**.

### Identifiants

- Votre **adresse e-mail de connexion**
- Votre **mot de passe personnel**

### Ce que vous pouvez faire

- Créer un compte utilisateur
- Modifier un compte
- Réinitialiser ou changer un mot de passe
- Attribuer un rôle
- Associer un utilisateur à une ou plusieurs provinces
- Gérer les listes de référence

### Procédure pas à pas

1. Ouvrez [superadmin.html](superadmin.html).
2. Connectez-vous avec votre e-mail et votre mot de passe.
3. Accédez à la gestion des utilisateurs.
4. Pour créer un compte :
   - saisissez l'adresse e-mail
   - saisissez le nom affiché
   - choisissez le rôle
   - affectez la province si nécessaire
   - définissez un mot de passe initial
5. Vérifiez ensuite que l'utilisateur peut ouvrir la bonne page selon son rôle.
6. Si un compte est mal configuré, corrigez le rôle ou les provinces autorisées.
7. Utilisez aussi cet espace pour maintenir les données de référence et le paramétrage général.

### Règle importante

Un compte mal configuré donne souvent l'impression que l'application ne fonctionne pas. En réalité, c'est parfois simplement le **rôle** ou la **province autorisée** qui ne correspond pas.

---

## 5. Résumé très simple

- **Utilisateur ordinaire** : saisit une mission et envoie le rapport.
- **Comptable** : contrôle le rapport, corrige ou valide la partie financière.
- **Super-administrateur** : crée les comptes et gère les accès.

---

## 6. Aide en cas de problème

Si vous n'arrivez pas à vous connecter :

- vérifiez que vous utilisez la bonne adresse e-mail
- vérifiez le mot de passe
- vérifiez que vous ouvrez la bonne page selon votre rôle
- demandez au super-administrateur de confirmer votre compte et votre rôle

Si le problème persiste, il faut demander de l'aide au responsable informatique ou au gestionnaire des comptes.
