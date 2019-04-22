

1. Accédez à la page [portail Azure : enregistrement des applications](https://go.microsoft.com/fwlink/?linkid=2083908) pour enregistrer votre application.

1. Connectez-vous à votre client Office 365 en utilisant les informations d’identification d’***administrateur***. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez **Nouvelle inscription**. Sur la page **Inscrire une application**, définissez les valeurs comme suit.

    * Donner **$ADD-IN-NAME$** à **Name**.
    * Définissez les **Types de comptes pris en charge** à **Comptes dans un annuaire organisationnel et comptes personnels Microsoft (par ex. Skype, Xbox, Outlook.com)**.
    * Laissez **Redirect URI** vide.
    * Choisissez **Inscrire**.

1. Sur la page **$ADD-IN-NAME$**, copiez et enregistrez les valeurs pour l’**ID de l’application (client)** et l’**ID de répertoire (client)**. Vous utiliserez les deux plus tard.

    > [!NOTE]
    > Cet ID a la valeur « audience » lorsque d’autres applications, telles que l’application hôte Office (par exemple, PowerPoint, Word, Excel) demandent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Sélectionnez **Certificats et secrets** sous **Gérer**. Sélectionnez le bouton **Nouveau secret client**. Entrer une valeur pour **Description** puis sélectionnez une option appropriée pour **Expire le** puis **Ajouter**. *Copier la valeur secrète client immédiatement et enregistrez-la avec l’ID d’application* avant de continuer car vous en aurez besoin dans une procédure plus loin.

1. Sélectionnez **Exposer une API** sous **Gérer**. Sélectionnez le lien **Définir** pour générer l’URI de l’ID d’application sous la forme "api://$App ID GUID$". Insérez **$FQDN-WITHOUT-PROTOCOL$** (avec une barre oblique « / » à la fin) entre les doubles barres obliques et le GUID. La forme de l’ID entier doit être `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; par exemple`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Il se peut que vous obteniez une erreur inexacte à ce stade indiquant « l’URI de l’ID d’application doit être une URI valide commençant par HTTPS, API, URN, MS-APPX. Elle ne doit pas se terminer pas par une barre oblique. » Si l’ID respecte les conditions indiquées, ignorez l’erreur et enregistrez vos modifications.

    > [!NOTE]
    > Si un message d’erreur s’affiche indiquant que le domaine appartient déjà à quelqu’un et que c’est vous qui en êtes le propriétaire, suivez la procédure décrite dans Quickstart [ : Ajouter votre nom de domaine personnalisé à l’aide du Portail Azure Active Directory](/azure/active-directory/add-custom-domain) pour l’inscrire, puis répétez cette étape. (Cette erreur peut également se produire si vous ne vous êtes pas connecté au client Office 365 avec les informations d’identification d’un administrateur. Voir l’étape 2. Déconnectez-vous, puis reconnectez-vous avec les informations d’identification d’administrateur, puis répétez le processus décrit à l’étape 3.)

1. Sélectionnez le bouton **Ajouter une étendue**. Dans le volet qui s’ouvre, entrez `access_as_user` en tant que **nom de l’étendue**.

1. Donnez la valeur **Administrateurs et utilisateurs** à **Qui peut donner son consentement ?** .

1. Renseignez les champs pour configurer les invites de consentement des administrateurs et utilisateurs avec les valeurs appropriées pour l’étendue `access_as_user` qui permet à l’application Office hôte d’utiliser l’API web de votre complément avec les mêmes droits que l’utilisateur actuel. Suggestions :

    - **Titre consentement administrateur :** Office peut agir en tant qu’utilisateur.
    - **Description consentement administrateur :** activez Office pour qu’il appelle l’API de complément web avec les mêmes droits que l’utilisateur actuel.
    - **Titre consentement utilisateur :** Office peut agir en tant que vous.
    - **Description consentement administrateur :** activez Office pour qu’il appelle l’API de complément web avec les mêmes droits dont vous disposez.

1. Vérifiez que **State** est défini comme **Enabled**.

1. Sélectionnez **Ajouter une étendue**.

    > [!NOTE]
    > La partie domaine du **nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’**URI d’ID d’application** définie à l’étape précédente avec `/access_as_user` ajouté au bout (par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`).

1. Dans la section **Applications client autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé.
  
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

    Pour chaque ID, procédez comme suit :

      a. Sélectionnez le bouton **Ajouter une application cliente** puis, dans le volet qui s’ouvre, définissez l’**ID Client** pour le GUID respectif et cochez la case pour `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.

      b. Sélectionnez **Ajouter une application**.

1. Sous **Gérer**, sélectionnez **Authentification**. Dans la section **Rediriger les URI**, sélectionnez **Web** dans la liste déroulante **Type** puis définissez la valeur de **Rediriger l’URI** à `https://$FQDN-WITHOUT-PROTOCOL$`.

1. En haut du formulaire, sélectionnez **Enregistrer**.

1. Sélectionnez **Autorisations API** sous **Gestion** et sélectionnez **Ajouter une autorisation**. Dans le volet qui s’ouvre, sélectionnez **Microsoft Graph**, puis **Autorisations déléguées**.

1. Utilisez la zone de recherche **Sélectionnez les autorisations** pour rechercher les autorisations dont votre complément a besoin. Les éléments suivants sont des exemples.

    * Files.Read.All
    * offline_access
    * openid
    * profil

    > [!NOTE]
    > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Une bonne pratique consiste à demander uniquement les autorisations dont vous avez besoin. Ainsi, nous vous recommandons de désactiver la case à cocher de cette autorisation si votre complément n’en a pas réellement besoin.

1. Sélectionnez la case à cocher pour chacune des autorisations comme elle apparaît (notez que les autorisations ne restent pas visibles dans la liste lorsque vous sélectionnez chacune d’elles). Après avoir sélectionné les autorisations dont votre complément a besoin, sélectionnez le bouton **Ajouter des autorisations** situé en bas du panneau.
