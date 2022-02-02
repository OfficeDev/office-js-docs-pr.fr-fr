## <a name="create-an-app-registration"></a>Créer une inscription d’application

L’inscription de votre application (le Plateforme d'identités Microsoft) établit une relation d’Plateforme d'identités Microsoft. L’relation d’confiance est unidirectionnelle : votre Plateforme d'identités Microsoft et non l’inverse.

1. Connectez-vous au [portail Azure avec](https://portal.azure.com/) les informations d’identification ***admin** _ à Microsoft 365 client. Par exemple, _*MyName@contoso.onmicrosoft.com**.
1. Sous **Gérer**, sélectionnez **Inscriptions de l’applicationNouveau** >  **inscription**. Sur la page **Inscrire une application**, définissez les valeurs comme suit.

    * Définissez le **Nom** sur `<add-in-name>`.
    * Définissez les **types** de comptes pris en charge sur les comptes dans n’importe quel répertoire d’organisation (n’importe quel répertoire **Azure AD - multi-client) et les comptes Microsoft personnels (par exemple, Skype, Xbox).**
    * Laissez **Redirect URI** vide.
    * Choisissez **Inscrire**.

1. Copiez et enregistrez les valeurs de **l’ID d’application (client)** et de **l’ID d’annuaire (client**). Vous utiliserez les deux plus tard.

    > [!NOTE]
    > Cet ID est la valeur « audience » lorsque d’autres applications, telles que l’application cliente Office (par exemple, PowerPoint, Word, Excel), recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

## <a name="add-a-client-secret"></a>Ajouter une secret client

Parfois appelée mot de _passe d’application_, une question secrète client est une valeur de chaîne que votre application peut utiliser à la place d’un certificat pour s’identifier elle-même.

1. Dans le portail Azure, dans les **inscriptions d’applications**, sélectionnez votre application.
1. **Sélectionnez Certificats &** **secretsClientNouveau** >  >  **secret client**.
1. Ajoutez une description pour votre secret client.
1. Sélectionnez une date d’expiration pour la secret ou spécifiez une durée de vie personnalisée.
    * La durée de vie de la secret client est limitée à deux ans (24 mois) ou moins. Vous ne pouvez pas spécifier une durée de vie personnalisée de plus de 24 mois.
    * Microsoft recommande de définir une valeur d’expiration inférieure à 12 mois.
1. Sélectionnez **Ajouter**.
1. _Enregistrez la valeur de la secret_ à utiliser dans le code de votre application cliente. Cette valeur secrète ne _s’affiche plus jamais_ une fois que vous avez quitté cette page.

## <a name="expose-a-web-api"></a>Exposer une API web

1. Assurez-vous que vous affichez l’inscription d’application que vous avez créée.
1. Sous **Gérer**, **sélectionnez Exposer une API**, puis sélectionnez **le lien** Définir. Cela ouvre une zone **Définir l’URI de l’ID** d’application avec un URI d’ID d’application généré sous la forme `api://<application-id>`. Insérez votre nom de domaine complet avant .`<application-id>` L’ID entier doit avoir le formulaire `api://<fully-qualified-domain-name>/<application-id>`; par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    > [!NOTE]
    > Si un message d’erreur s’affiche indiquant que le domaine appartient déjà à quelqu’un et que c’est vous qui en êtes le propriétaire, suivez la procédure décrite dans Quickstart [ : Ajouter votre nom de domaine personnalisé à l’aide du Portail Azure Active Directory](/azure/active-directory/add-custom-domain) pour l’inscrire, puis répétez cette étape. (Cette erreur peut également se produire si vous n’êtes pas signé avec les informations d’identification d’un administrateur dans Microsoft 365 location. Voir l’étape 2. Déconnectez-vous, puis reconnectez-vous avec les informations d’identification d’administrateur, puis répétez le processus décrit à l’étape 3.)

## <a name="add-a-scope"></a>Ajouter une étendue

1. Sélectionnez le bouton **Ajouter une étendue**. Dans le volet qui s’ouvre, entrez `access_as_user` en tant que **nom de l’étendue**.

1. Donnez la valeur **Administrateurs et utilisateurs** à **Qui peut donner son consentement ?** .

1. Remplissez les champs pour configurer les invites de consentement de l’administrateur et de l’utilisateur avec des valeurs `access_as_user` appropriées pour l’étendue, ce qui permet à l’application cliente Office d’utiliser les API web de votre add-in avec les mêmes droits que l’utilisateur actuel. Suggestions :

    * **Nom complet du consentement de** l’administrateur : Office peut agir en tant qu’utilisateur.
    * **Description consentement administrateur :** activez Office pour qu’il appelle l’API de complément web avec les mêmes droits que l’utilisateur actuel.
    * **Nom complet du consentement de** l’utilisateur : Office peut agir en votre nom.
    * **Description du consentement d’utilisateur :** Activez Office pour qu’il appelle l’API du complément web avec les mêmes droits dont vous disposez.

1. Vérifiez que **State** est défini comme **Enabled**.

1. Sélectionnez **Ajouter une étendue**.

    > [!NOTE]
    > La partie domaine du **nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’**URI d’ID d’application** définie à l’étape précédente avec `/access_as_user` ajouté au bout (par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`).

1. Dans la section **Applications client autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé.
  
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    * `93d53678-613d-4013-afc1-62e9e444a0a5` (Office sur le web)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office sur le web)
    * `08e18876-6177-487e-b8b5-cf950c1e598c` (Office sur le web)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook sur le web)

    > [!NOTE]
    > L’ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` inclut tous les autres ID répertoriés et peut être utilisé de manière unique pour pré-autoriser tous les points de terminaison hôtes Office à utiliser avec votre service dans le flux d’utilisateur unique du Office.

    Pour chaque ID, prenez les mesures suivantes.

      a. Sélectionnez **Ajouter une application cliente**. Dans le panneau qui s’ouvre, définissez **l’ID client** sur le GUID respectif et cochez la case .`api://<fully-qualified-domain-name>/<application-id>/access_as_user`

      b. Sélectionnez **Ajouter une application**.

## <a name="add-microsoft-graph-permissions"></a>Ajouter des autorisations Graph Microsoft

1. Sous **Gérer**, sélectionnez **Authentification**, puis **ajoutez une plateforme**.

1. Dans le **volet Configurer les plateformes** , sélectionnez **Web**, puis définissez la valeur **d’URI** de redirection sur `https://<fully-qualified-domain-name>`.

1. Choisissez **Configurer**.

1. Sous **Gérer**, sélectionnez **les autorisations d’API**, puis **sélectionnez Ajouter une autorisation**. Dans le panneau qui s’ouvre, choisissez **Microsoft Graph**, puis les **autorisations déléguées**.

1. Utilisez la zone de recherche **Sélectionnez les autorisations** pour rechercher les autorisations dont votre complément a besoin. Les éléments suivants sont des exemples.

    * Files.Read.All
    * offline_access
    * openid
    * profil

    > [!NOTE]
    > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Il est recommandé de demander uniquement les autorisations nécessaires. Nous vous recommandons donc de décocher la case pour cette autorisation si votre add-in n’en a pas réellement besoin.

1. Sélectionnez la case à cocher pour chacune des autorisations comme elle apparaît (notez que les autorisations ne restent pas visibles dans la liste lorsque vous sélectionnez chacune d’elles). Après avoir sélectionné les autorisations dont votre add-in a besoin, sélectionnez le **bouton Ajouter des autorisations** .
