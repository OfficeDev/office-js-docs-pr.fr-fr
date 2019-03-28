
> [!NOTE]
> Cette procédure est uniquement nécessaire quand vous développez le complément. Lorsque votre complément de production est déployé dans AppSource ou dans un catalogue de compléments, les utilisateurs l’approuvent individuellement ou un administrateur l’approuvera pour l’organisation au moment de l’installation.

Exécutez cette procédure *après* avoir [enregistré le complément](../develop/register-sso-add-in-aad-v2.md).

1. Dans la chaîne suivante, remplacez l’espace réservé « {application_ID} » par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément : `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Collez l’URL résultante dans la barre d’adresse d’un navigateur pour y accéder.

1. Lorsque vous y êtes invité, connectez-vous avec les informations d’identification d’administrateur à votre client Office 365.

1. Vous êtes ensuite invité à accorder des autorisations pour votre complément pour accéder à vos données Microsoft Graph. Cliquez sur **Accepter**.

1. La fenêtre/l'onglet du navigateur est ensuite redirigé vers l' **URL** de redirection que vous avez spécifiée lorsque vous avez enregistré le complément. Si l'application Web du complément est en cours d'exécution, la page d'accueil du complément s'ouvre dans le navigateur; dans le cas contraire, vous obtiendrez une erreur 404. Mais le fait que le navigateur a tenté d'ouvrir la page d'accueil signifie que le consentement a bien été accordé.

>[!NOTE]
>Nous vous recommandons d'utiliser cette procédure en tant que meilleure pratique si vous utilisez un client O365 de développeur. Toutefois, si vous préférez, vous pouvez chargement un complément d'authentification unique en cours de développement et inviter l'utilisateur à fournir un formulaire de consentement. Pour plus d'informations, reportez-vous à [chargement sur Windows](/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) et [chargement sur Office Online](/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).
