# Courses & Menus — application familiale

Application web mobile et PC basée sur Google Apps Script et un Google Sheet privé.

## Installation

1. Importer `Courses & menus - Pour Codex.xlsx` dans Google Drive, puis l’ouvrir avec Google Sheets et l’enregistrer au format Google Sheets.
2. Dans ce nouveau classeur, ouvrir **Extensions → Apps Script**.
3. Remplacer le contenu de `Code.gs` par celui du fichier `Code.gs` fourni.
4. Ajouter un fichier HTML nommé `Index` et y coller le contenu de `Index.html`.
5. Dans les paramètres du projet, afficher le fichier manifeste et remplacer son contenu par `appsscript.json`.
6. Exécuter une première fois `setupApp` depuis l’éditeur et accepter les autorisations.
7. Exécuter une fois `importLegacyData`. Cette fonction importe directement les onglets historiques du classeur courant.
8. Choisir **Déployer → Nouveau déploiement → Application Web** :
   - Exécuter en tant que : **Moi**
   - Qui a accès : **Tous les utilisateurs** ou **Toute personne**
9. Copier l’URL `/exec` et la partager avec la famille.

Sur iPhone, ouvrir l’URL dans Safari puis **Partager → Sur l’écran d’accueil**.  
Sur Android, ouvrir l’URL dans Chrome puis **Ajouter à l’écran d’accueil**.

## Données

Les onglets préfixés `APP_` constituent la base de données. Ils peuvent être masqués, mais ne doivent pas être renommés ni réordonner leurs colonnes.

La liste de référence est permanente. La liste de courses active ne conserve pas l’historique détaillé des achats ; lorsqu’un article est coché, seule sa date de dernier achat est reportée dans le catalogue.

## Synchronisation

Les appareils vérifient les changements toutes les 3,5 secondes. Les écritures concurrentes sont protégées par un verrou Apps Script.
