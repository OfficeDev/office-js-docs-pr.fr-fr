

1. Accédez à la page [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).

1. Connectez-vous à votre client Office 365 en utilisant les informations d’identification de l’***administrateur***. Par exemple, MonNom@contoso.onmicrosoft.com

1. Cliquez sur **Ajouter une application**.

1. Lorsque vous y êtes invité, entrez **$ADD-IN-NAME$** comme nom d’application, puis appuyez sur **Créer une application**.

1. Quand la page de configuration de l’application s’ouvre, copiez l’**ID de l’application** et enregistrez-le. Vous l’utiliserez dans une procédure ultérieure.

    > [!NOTE]
    > Cet ID est la valeur « audience » lorsque d’autres applications, telles que l’application hôte Office (par exemple, PowerPoint, Word, Excel) recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Dans la section **Secrets de l’application**, appuyez sur **Générer un nouveau mot de passe**. Une boîte de dialogue contextuelle s’ouvre avec un nouveau mot de passe (également appelé « secret de l’application »). *Copiez le mot de passe immédiatement et enregistrez-le avec l’ID de l’application.* Vous en aurez besoin dans une procédure ultérieure. Ensuite, fermez la boîte de dialogue.

1. Dans la section **Plateformes**, cliquez sur **Ajouter une plateforme**.

1. Dans la boîte de dialogue qui s’ouvre, sélectionnez **API Web**.

1. Un **URI d’ID d’application** sous la forme « api://$App ID GUID$ » a été généré. Insérez **$FQDN-WITHOUT-PROTOCOL$** (avec une barre oblique « / » à la fin) entre les doubles barres obliques et le GUID. La forme de l’ID entier doit être `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; par exemple`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Si un message d’erreur s’affiche indiquant que le domaine appartient déjà à quelqu’un et que vous en êtes le propriétaire, suivez la procédure décrite dans [Ajouter votre nom de domaine personnalisé à l’aide du Portail Azure Active Directory](/azure/active-directory/add-custom-domain) pour l’inscrire, puis répétez cette étape. (Cette erreur peut également se produire si vous ne vous êtes pas connecté au client Office 365 avec les informations d’identification d’un administrateur. Voir l’étape 2. Déconnectez-vous, puis reconnectez-vous avec les informations d’identification d’administrateur, puis répétez le processus décrit à l’étape 3.)

    > [!NOTE]
    > La partie domaine du nom de l’**étendue**, juste en dessous de l’**URI d’ID d’application**, change automatiquement en conséquence, avec l’ajout de `/access_as_user` à la fin. Par exemple : `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Dans la section **Applications pré-autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé. Chaque fois que vous en entrez un, une nouvelle zone de texte vide s’affiche. (Entrez uniquement le GUID.)
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. Ouvrez le menu déroulant **Scope** à côté de chaque **ID d’application** et activez la case à cocher pour `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

1. En haut de la section **Plateformes**, cliquez sur **Ajouter une plateforme** à nouveau, puis sélectionnez **Web**.

1. Dans la nouvelle section **Web** sous **Plateformes**, entrez les informations suivantes en guise d’**URL de redirection** : `https://$FQDN-WITHOUT-PROTOCOL$`.

1. Faites défiler jusqu’à la section **Autorisations pour Microsoft Graph** et à la sous-section **Autorisations déléguées**. Utilisez le bouton **Ajouter** pour ouvrir une boîte de dialogue **Sélectionner des autorisations**.

1. Dans la boîte de dialogue, activez les cases à cocher en regard de `profile`, et tout autre autorisation AAD ou Microsoft Graph dont votre complément a besoin. Voici quelques exemples :

    * Files.Read.All
    * offline_access
    * openid
    * profil

    > [!NOTE]
    > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Une bonne pratique consiste à demander uniquement les autorisations dont vous avez besoin. Ainsi, nous vous recommandons de désactiver la case à cocher de cette autorisation si votre complément n’en a pas réellement besoin.

1. Cliquez sur **OK** au bas de la boîte de dialogue.

1. Cliquez sur **Enregistrer** au bas de la page d’inscription.
