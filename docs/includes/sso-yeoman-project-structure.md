### <a name="configuration"></a>Configuration

Les fichiers suivants spécifient les paramètres de configuration du module.

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.

- Le **./. Le fichier ENV** dans le répertoire racine du projet définit les constantes utilisées par le projet de add-in.

### <a name="task-pane"></a>Volet de tâches 

Les fichiers suivants définissent l’interface utilisateur et les fonctionnalités du volet Des tâches du module.

- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.

- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.

- Dans un projet JavaScript, le fichier **./src/taskpane/taskpane.js** contient le code d’initialisation du add-in. Dans un projet TypeScript, le fichier **./src/taskpane/taskpane.ts** contient du code pour initialiser le add-in, ainsi que du code qui utilise la bibliothèque d’API JavaScript Office pour ajouter les données de Microsoft Graph au document Office.

### <a name="authentication"></a>Authentification

Les fichiers suivants facilitent le processus DSO et écrivent des données dans Office document.

- Dans un projet JavaScript, le fichier **./src/helpers/documentHelper.js** contient du code qui utilise la bibliothèque d’API JavaScript Office pour ajouter les données de Microsoft Graph au document Office. Il n’existe aucun fichier de ce type dans un projet TypeScript ; Le code qui utilise la bibliothèque d’API JavaScript Office pour ajouter les données de Microsoft Graph au document Office existe dans **./src/taskpane/taskpane.ts** à la place.

- Le **fichier ./src/helpers/fallbackauthdialog.html** est la page sans interface utilisateur qui charge javaScript pour la stratégie d’authentification de secours.

- Le **fichier ./src/helpers/fallbackauthdialog.js** contient le JavaScript pour la stratégie d’authentification de secours qui se signe à l’utilisateur avec msal.js.

- Le fichier **./src/helpers/fallbackauthhelper.js** contient le javaScript du volet Des tâches qui appelle la stratégie d’authentification de secours dans les scénarios où l’authentification sso n’est pas prise en charge.

- Le fichier **./src/helpers/ssoauthhelper.js** contient l’appel JavaScript à l’API de l’authentification unique, `getAccessToken`, reçoit le jeton d’amorçage, initialise le remplacement du jeton d’amorçage pour un jeton d’accès à Microsoft Graph et appelle Microsoft Graph pour les données.