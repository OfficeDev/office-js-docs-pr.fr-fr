### <a name="configuration"></a>Configuration

Les fichiers suivants spécifient les paramètres de configuration du complément.

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.

- Le fichier **./.ENV** dans le répertoire racine du projet définit les constantes utilisées par le projet de complément.

### <a name="task-pane"></a>Volet de tâches

Les fichiers suivants définissent l’interface utilisateur et les fonctionnalités du volet des tâches du complément.

- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.

- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.

- Dans un projet JavaScript, le fichier **./src/taskpane/taskpane.js** contient le code d’initialisation du complément. Dans un projet TypeScript, le fichier **./src/taskpane/taskpane.ts** contient le code d’initialisation du complément, ainsi que le code qui utilise la bibliothèque d’API JavaScript Office pour ajouter les données de Microsoft Graph au document Office.

### <a name="authentication"></a>Authentification

Les fichiers suivants facilitent le processus SSO et écrivent des données dans le document Office.

- Dans un projet JavaScript, le fichier **./src/helpers/documentHelper.js** contient le code qui utilise la bibliothèque d’API JavaScript Office pour ajouter les données de Microsoft Graph au document Office. Il n’existe aucun fichier de ce type dans un projet TypeScript ; le code qui utilise la bibliothèque d’API JavaScript Office pour ajouter les données de Microsoft Graph au document Office existe dans **./src/taskpane/taskpane.ts** à la place.

- Le fichier **./src/helpers/fallbackauthdialog.html** est la page sans interface utilisateur qui charge le JavaScript pour la stratégie d’authentification de secours.

- Le fichier **./src/helpers/fallbackauthdialog.html** contient le JavaScript de la méthode d’authentification de secours qui connecte l'utilisateur avec msal.js.

- Le fichier **./src/helpers/fallbackauthhelper.js** contient le javaScript du volet des tâches qui appelle la stratégie d’authentification de secours dans les scénarios où l’authentification SSO n’est pas prise en charge.

- Le fichier **./src/helpers/ssoauthhelper.js** contient l’appel JavaScript à l’API de SSO, `getAccessToken`, reçoit le jeton d’accès, initialise le remplacement du jeton d’accès pour un nouveau jeton d’accès avec des autorisations à Microsoft Graph et appelle Microsoft Graph pour les données.