

1. Aller vers [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).

1. Connectez-vous avec les informations d’identification d’administrateur à votre client Office 365. Par exemple, MyName@contoso.onmicrosoft.com

1. Cliquez sur **Ajouter une application**.

1. Lorsque vous y êtes invité, entrez **$ADD-IN-NAME$** comme nom de l'application, puis appuyez sur **Créez l'application**.

1. Quand la page de configuration de l’application s’ouvre, copiez **l'ID de l’application** et enregistrez-le. Vous l’utiliserez dans une procédure ultérieure.

    > [!NOTE]
    > Cet ID est la valeur « audience » lorsque d’autres applications, telles que l’application hôte Office (par exemple, PowerPoint, Word, Excel) recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Dans la section **Secrets de l’application**, appuyez sur **Générer un nouveau mot de passe**. Une boîte de dialogue contextuelle s’ouvre avec un nouveau mot de passe (également appelé « secret de l’application »). *Copiez le mot de passe immédiatement et enregistrez-le avec l’ID de l’application.* Vous en aurez besoin dans une procédure ultérieure. Ensuite, fermez la boîte de dialogue.

1. Dans la section **Plateformes**, cliquez sur **Ajouter une plateforme**.

1. Dans la boîte de dialogue qui s’ouvre, sélectionnez **API Web**.

1. L'URl **de l'ID de l'application** est génée à partir du for mulaire “api://$App ID GUID$”. Insérez le **$FQDN-SANS-PROTOCOLE$** (avec une barre oblique de division (/) ajoutée à la fin) entre deux barres obliques de division et le GUID. L'identifiant complet doit avoir la forme `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$` ; par exemple `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Si vous obtenez une erreur indiquant que le domaine est déjà possédé, mais que vous le possédez, suivez la procédure [Quickstart: Ajouter un nom de domaine personnalisé à Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) pour l'enregistrer, puis répétez cette étape.

    > [!NOTE]
    > La partie domaine du nom de **l'Étendue** situé juste en dessous de **l'URI de l'ID de l'application** changera automatiquement pour s'adapter, avec `/access_as_user` ajouté à la fin ; par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Dans la section **Applications pré-autorisées,** identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé. Chaque fois que vous en entrez un, une nouvelle zone de texte vide s’affiche. (Entrez uniquement le GUID.)
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. Ouvrez le menu déroulant **Scope** à côté de chaque **ID d’application** et activez la case à cocher pour `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

1. En haut de la section **Plateformes**, cliquez sur **Ajouter une plateforme** à nouveau, puis sélectionnez **Web**.

1. Dans la nouvelle section **Web** sous **Plateformes**, entrez les informations suivantes en guise d’**URL de redirection** : `https://$FQDN-WITHOUT-PROTOCOL$`.

1. Faites défiler jusqu’à la section **Autorisations pour Microsoft Graph** et à la sous-section **Autorisations déléguées**. Utilisez le bouton **Ajouter** pour ouvrir une boîte de dialogue **Sélectionner des autorisations**.

1. Dans la boîte de dialogue, cochez les cases pour `profile` et toutes les autres autorisations AAD et Microsoft Graph dont votre complément a besoin. Les éléments suivants en sont des exemples :

    * Files.Read.All
    * offline_access
    * openid
    * profil

    > [!NOTE]
    > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Il est conseillé de ne pas demander d'autorisation qui ne sont pas nécessaires. Ainsi, nous vous recommandons de décocher la case pour cette autorisation si votre complément n'en a pas réellement besoin.

1. Au bas de la boîte de dialogue, cliquez sur **OK**.

1. Au bas de la page d’inscription, cliquez sur **Enregistrer**.
