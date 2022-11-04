## <a name="register-the-add-in-with-microsoft-identity-platform"></a>Inscrire le complément auprès de Plateforme d'identités Microsoft

Vous devez créer une inscription d’application dans Azure qui représente votre serveur web. Cela permet la prise en charge de l’authentification afin que les jetons d’accès appropriés puissent être émis au code client dans JavaScript. Cette inscription prend en charge l’authentification unique dans le client et l’authentification de secours à l’aide de la bibliothèque d’authentification Microsoft (MSAL).


1. Connectez-vous au [Portail Azure](https://portal.azure.com/) avec les informations d’identification ***admin** _ de votre location Microsoft 365. Par exemple, _*MyName@contoso.onmicrosoft.com**.
1. Sélectionner les **inscriptions d’applications**. Si vous ne voyez pas l’icône, recherchez « Inscription de l’application » dans la barre de recherche.

    :::image type="content" source="../images/azure-portal-select-app-registration.png" alt-text="Page d’accueil Portail Azure.":::

    La page **Inscriptions d'applications** s’affiche.

1. Sélectionnez **Nouvelle inscription**.

    :::image type="content" source="../images/azure-portal-select-new-registration.png" alt-text="Nouvelle inscription dans le volet inscriptions d'applications.":::

    Le **volet Inscrire une application s’affiche** .

1. Sous **Gérer**, sélectionnez **inscriptions d'applications** >  **Nouvelle inscription**. Dans le volet **Inscrire une application, définissez** les valeurs comme suit.

    * Définissez le **Nom** sur `<add-in-name>`.
    * Définissez **Types de comptes pris en charge** **sur Comptes dans n’importe quel annuaire organisationnel (n’importe quel annuaire Azure AD - multilocataire) et comptes Microsoft personnels (par exemple, Skype, Xbox).**
    * Définissez **URI de redirection** pour utiliser la plateforme `<redirect-platform>` et l’URI sur `<redirect-uri>`.

    :::image type="content" source="../images/azure-portal-register-an-application.png" alt-text="Inscrire un volet d’application avec le nom et le compte pris en charge terminés.":::

1. Sélectionnez **Inscrire**. Un message s’affiche indiquant que l’inscription de l’application a été créée.

    :::image type="content" source="../images/azure-portal-application-created-message.png" alt-text="Message indiquant que l’inscription de l’application a été créée.":::

1. Copiez et enregistrez les valeurs de **l’ID d’application (client)** et de **l’ID d’annuaire (locataire).** Vous utiliserez les deux plus tard.

    :::image type="content" source="../images/azure-portal-copy-client-directory-ids.png" alt-text="Volet Inscription d’application pour Contoso affichant l’ID client et l’ID d’annuaire.":::

## <a name="add-a-client-secret"></a>Ajouter une clé secrète client

Parfois appelé _mot de passe d’application_, une clé secrète client est une valeur de chaîne que votre application peut utiliser à la place d’un certificat pour s’identifier elle-même.

1. Sélectionnez **Certificats & secrets**. Ensuite, sous l’onglet **Secrets client** , sélectionnez **Nouvelle clé secrète client**.

    :::image type="content" source="../images/azure-portal-create-new-client-secret.png" alt-text="Volet Certificats & secrets.":::

    Le volet **Ajouter une clé secrète client** s’affiche.

1. Ajoutez une description pour votre clé secrète client.
1. Sélectionnez une expiration pour le secret ou spécifiez une durée de vie personnalisée.
    * La durée de vie de la clé secrète client est limitée à deux ans (24 mois) ou moins. Vous ne pouvez pas spécifier une durée de vie personnalisée supérieure à 24 mois.
    * Microsoft vous recommande de définir une valeur d’expiration inférieure à 12 mois.

    :::image type="content" source="../images/azure-portal-client-secret-description.png" alt-text="Ajoutez un volet de clé secrète client avec la description et expire.":::

1. Sélectionnez **Ajouter**. Le nouveau secret est créé, la valeur est affichée temporairement.

> [!IMPORTANT]
> _Enregistrez la valeur du secret_ à utiliser dans le code de votre application cliente. Cette valeur de secret _n’est plus jamais affichée une fois_ que vous avez quitté ce volet.

## <a name="expose-a-web-api"></a>Exposer une API web

1. Sélectionnez **Exposer une API**.

    Le volet **Exposer une API s’affiche** .

    :::image type="content" source="../images/azure-portal-expose-an-api.png" alt-text="Volet Exposer une API d’une inscription d’application.":::

1. Sélectionnez **Définir** pour générer un URI d’ID d’application.

    :::image type="content" source="../images/azure-portal-set-api-uri.png" alt-text="Bouton Définir dans le volet Exposer une API de l’inscription de l’application.":::

    La section permettant de définir l’URI de l’ID d’application s’affiche avec un URI d’ID d’application généré au format `api://<app-id>`.

1. Mettez à jour l’URI de l’ID d’application sur `api://localhost:44355/<app-id>`.

    :::image type="content" source="../images/azure-portal-app-id-uri-details.png" alt-text="Modifiez le volet URI d’ID d’application avec le port localhost défini sur 44355.":::

    * **L’URI de l’ID d’application** est prérempli avec l’ID d’application (GUID) au format `api://<app-id>`.
    * Le format d’URI de l’ID d’application doit être : `api://<fully-qualified-domain-name>/<app-id>`
    * Insérez entre `fully-qualified-domain-name` `api://` et `<app-id>` (qui est un GUID). Par exemple : `api://contoso.com/<app-id>`.
    * Si vous utilisez localhost, le format doit être `api://localhost:<port>/<app-id>`. Par exemple : `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

    Pour plus d’informations sur l’URI de l’ID d’application, consultez [Attribut identificateur de manifeste d’applicationUris](/azure/active-directory/develop/reference-app-manifest#identifieruris-attribute).

    > [!NOTE]
    > Si un message d’erreur s’affiche indiquant que le domaine appartient déjà à quelqu’un et que c’est vous qui en êtes le propriétaire, suivez la procédure décrite dans Quickstart [ : Ajouter votre nom de domaine personnalisé à l’aide du Portail Azure Active Directory](/azure/active-directory/add-custom-domain) pour l’inscrire, puis répétez cette étape. (Cette erreur peut également se produire si vous n’êtes pas connecté avec les informations d’identification d’un administrateur dans la location Microsoft 365. Voir l’étape 2. Déconnectez-vous, puis reconnectez-vous avec les informations d’identification d’administrateur, puis répétez le processus décrit à l’étape 3.)

## <a name="add-a-scope"></a>Ajouter une étendue

1. Sélectionnez **Ajouter une étendue**.

    :::image type="content" source="../images/azure-portal-add-a-scope.png" alt-text="Sélectionnez le bouton Ajouter une étendue.":::

    Le volet **Ajouter une étendue** s’ouvre.

1. Dans le volet **Ajouter une étendue** , spécifiez les attributs de l’étendue .

    :::image type="content" source="../images/azure-portal-add-a-scope-details.png" alt-text="Ajoutez un volet d’étendue avec des exemples de valeurs.":::

    | Champ | Description | Valeurs |
    |-------|-------------|---------|
    | **Nom de l'étendue** | Nom de votre étendue. Une convention de nommage d’étendue courante est `resource.operation.constraint`. | Pour l’authentification unique, cette valeur doit être définie sur `access_as_user`. |
    | **Qui peut consentir** |  Détermine si le consentement de l’administrateur est requis ou si les utilisateurs peuvent donner leur consentement sans approbation de l’administrateur. | Pour découvrir l’authentification unique et les exemples, nous vous recommandons de définir cette option sur **Administrateurs et utilisateurs**. <br><br>Sélectionnez **Administrateurs uniquement pour obtenir des** autorisations à privilèges plus élevés.|
    | **Administration nom d’affichage du consentement** | Brève description de l’objectif de l’étendue visible uniquement par les administrateurs. | `Read-only access to user files and profiles.` |
    | **Administration description du consentement** | Description plus détaillée de l’autorisation accordée par l’étendue que seuls les administrateurs voient. | `Allow Office to have read-only access to all user files and profiles. Office can call the app's web APIs as the current user.` |
    | **Nom d’affichage du consentement de l’utilisateur** | Brève description de l’objectif de l’étendue. Affiché aux utilisateurs uniquement si vous définissez **Qui peut donner son consentement** aux **administrateurs et aux utilisateurs**. | `Read-only access to your files and profile.` |
    | **Description du consentement de l’utilisateur** | Description plus détaillée de l’autorisation accordée par l’étendue. Affiché aux utilisateurs uniquement si vous définissez **Qui peut donner son consentement** aux **administrateurs et aux utilisateurs**. | `Allow Office to have read-only access to your files and user profile.` |

1. Définissez **l’état** **sur Activé**, puis sélectionnez **Ajouter une étendue**.

    :::image type="content" source="../images/azure-portal-enable-state-add-scope-button.png" alt-text="Définissez l’état sur activé et sélectionnez le bouton Ajouter une étendue.":::

    La nouvelle étendue que vous avez définie s’affiche dans le volet.

    :::image type="content" source="../images/azure-portal-scope-added-successfully.png" alt-text="Nouvelle étendue affichée dans le volet Exposer une API.":::

    > [!NOTE]
    > La partie domaine du **nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’**URI d’ID d’application** définie à l’étape précédente avec `/access_as_user` ajouté au bout (par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`).

1. Sélectionnez **Ajouter une application cliente.**

    :::image type="content" source="../images/azure-portal-add-a-client-application.png" alt-text="Sélectionnez Ajouter une application cliente.":::

    Le volet **Ajouter une application cliente s’affiche** .

1. Dans **l’ID client** , entrez `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`. Cette valeur pré-autorise tous les points de terminaison d’application Microsoft Office.

    > [!NOTE]
    > L’ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` pré-autorise Office sur toutes les plateformes suivantes. Vous pouvez également entrer un sous-ensemble approprié des ID suivants si, pour une raison quelconque, vous souhaitez refuser l’autorisation à Office sur certaines plateformes. Laissez simplement de côté les ID des plateformes à partir desquelles vous souhaitez refuser l’autorisation. Les utilisateurs de votre complément sur ces plateformes ne pourront pas appeler vos API web, mais d’autres fonctionnalités de votre complément fonctionneront toujours.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office sur le web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook sur le web)

1. Dans **Étendues autorisées**, cochez la `api://localhost:44355/<app-id>/access_as_user` case.

1. Sélectionnez **Ajouter une application**.

    :::image type="content" source="../images/azure-portal-add-application.png" alt-text="Volet Ajouter une application cliente.":::

## <a name="add-microsoft-graph-permissions"></a>Ajouter des autorisations Microsoft Graph

1. Sélectionnez **Autorisations API**.

    :::image type="content" source="../images/azure-portal-api-permissions.png" alt-text="Volet Autorisations de l’API.":::

    Le volet **Autorisations de l’API** s’ouvre.

1. Sélectionnez **Ajouter une autorisation**.

    :::image type="content" source="../images/azure-portal-add-a-permission.png" alt-text="Ajout d’une autorisation dans le volet Autorisations de l’API.":::

    Le volet **Demander des autorisations d’API** s’ouvre.

1. Sélectionnez **Microsoft Graph**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-graph.png" alt-text="Bouton Demander des autorisations d’API avec Microsoft Graph.":::

1. Sélectionnez **Autorisations déléguées**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-delegated.png" alt-text="Bouton Demander des autorisations d’API avec autorisations déléguées.":::

1. Dans la zone de recherche **Sélectionner des autorisations** , recherchez les autorisations dont votre complément a besoin. Voici les valeurs classiques utilisées dans les exemples.

    * Files.Read
    * openid
    * profil

    > [!NOTE]
    > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Comme il est recommandé de demander uniquement les autorisations nécessaires, nous vous recommandons de décocher la case pour cette autorisation si votre complément n’en a pas réellement besoin.

1. Cochez la case pour chaque autorisation telle qu’elle apparaît. Notez que les autorisations ne restent pas visibles dans la liste lorsque vous sélectionnez chacune d’elles. Après avoir sélectionné les autorisations dont votre complément a besoin, sélectionnez **Ajouter des autorisations**.

    :::image type="content" source="../images/azure-portal-request-api-permissions-add-permissions.png" alt-text="Volet Demander des autorisations d’API avec certaines autorisations sélectionnées.":::

## <a name="configure-access-token-version"></a>Configurer la version du jeton d’accès

Vous devez définir la version du jeton d’accès acceptable pour votre application. Cette configuration est effectuée dans le manifeste de l’application Azure Active Directory.

### <a name="define-the-access-token-version"></a>Définir la version du jeton d’accès

La version du jeton d’accès peut changer si vous avez choisi un type de compte autre que **Comptes dans un annuaire organisationnel (n’importe quel annuaire Azure AD - Multilocataire) et des comptes Microsoft personnels (par exemple, Skype, Xbox).** Procédez comme suit pour vous assurer que la version du jeton d’accès est correcte pour l’utilisation de l’authentification unique Office.

1. Sélectionnez **Gérer** > **Manifeste** dans le volet gauche.

    :::image type="content" source="../images/azure-portal-manifest.png" alt-text="Sélectionnez Manifeste Azure.":::

    Le manifeste de l’application Azure Active Directory s’affiche.

1. Entrez **2** comme valeur pour la propriété `accessTokenAcceptedVersion`.

    :::image type="content" source="../images/azure-portal-manifest-token-version.png" alt-text="Valeur de la version du jeton d’accès acceptée.":::

1. Sélectionnez **Enregistrer**.

    Un message s’affiche sur le navigateur indiquant que le manifeste a été mis à jour avec succès.

    :::image type="content" source="../images/azure-portal-manifest-updated-message.png" alt-text="Message de manifeste mis à jour.":::

Félicitations ! Vous avez terminé l’inscription de l’application pour activer l’authentification unique pour votre complément Office.
