# Génération du fichier EXE

## Prérequis

- Node.js 24 ou plus récent

## Installation

```powershell
npm install
```

## Lancement en mode application desktop

```powershell
npm start
```

## Génération de l'exécutable Windows

```powershell
npm run dist
```

Le fichier `.exe` sera généré dans le dossier `dist/`.

## Notes

- Le format généré est `portable`, donc un seul `.exe` sans installation obligatoire.
- Les données locales restent stockées sur le poste utilisateur via le moteur Chromium/Electron.
- Le menu de l'application permet d'ouvrir directement le formulaire ou le tableau de bord admin.