### <a name="configuration"></a>Configuration

Les fichiers suivants spécifient les paramètres de configuration du complément.

- Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.

- Le **./. ENV** dans le répertoire racine du projet définit les constantes utilisées par le projet de complément.

### <a name="task-pane"></a>Volet de tâches 

Les fichiers suivants définissent l’interface utilisateur et les fonctionnalités du volet Office du complément.

- Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.

- Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.

- Dans un projet JavaScript, le fichier **./SRC/TaskPane/TaskPane.js** contient le code permettant d’initialiser le complément. Dans un projet de machine à écrire, le fichier **./SRC/TaskPane/TaskPane.TS** contient le code d’initialisation du complément, ainsi que le code qui utilise la bibliothèque de l’API JavaScript pour Office pour ajouter les données de Microsoft Graph au document Office.

### <a name="authentication"></a>Authentification

Les fichiers suivants facilitent le processus SSO et écrivent des données dans le document Office.

- Dans un projet JavaScript, le fichier **./SRC/helpers/documentHelper.js** contient du code qui utilise la bibliothèque de l’API JavaScript pour Office pour ajouter les données de Microsoft Graph au document Office. Il n’existe pas de fichier de ce type dans un projet de type dactylographié ; le code qui utilise la bibliothèque de l’API JavaScript pour Office pour ajouter les données de Microsoft Graph au document Office existe à la place dans **./SRC/TaskPane/TaskPane.TS** .

- Le fichier **./SRC/helpers/fallbackauthdialog.html** est la page sans interface utilisateur qui charge le code JavaScript pour la stratégie d’authentification de secours.

- Le fichier **./SRC/helpers/fallbackauthdialog.js** contient le code JavaScript pour la stratégie d’authentification de secours qui se connecte à l’utilisateur avec MSAL. js.

- Le fichier **./SRC/helpers/fallbackauthhelper.js** contient le code JavaScript du volet Office qui appelle la stratégie d’authentification de secours dans les scénarios lorsque l’authentification SSO n’est pas prise en charge.

- Le fichier **./src/helpers/ssoauthhelper.js** contient l’appel JavaScript à l’API de l’authentification unique, `getAccessToken`, reçoit le jeton d’amorçage, initialise le remplacement du jeton d’amorçage pour un jeton d’accès à Microsoft Graph et appelle Microsoft Graph pour les données.