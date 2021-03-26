---
title: Créer un complément Office ASP.NET qui utilise l’authentification unique
description: Guide pas à pas sur la création (ou la conversion) d’un add-in Office avec un système principal ASP.NET pour utiliser l' sign-on unique (SSO).
ms.date: 03/11/2021
localization_priority: Normal
ms.openlocfilehash: e92bac3be81254a4c15f5e071602edbe788692ac
ms.sourcegitcommit: 5ad32261f80e7ab371aba032d9024ad1275c23f9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/26/2021
ms.locfileid: "51221373"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>Créer un complément Office ASP.NET qui utilise l’authentification unique

Lorsque les utilisateurs sont connectés à Office, votre complément peut utiliser les mêmes informations d’identification pour permettre aux utilisateurs d’accéder à plusieurs applications sans avoir à se connecter une deuxième fois. Pour obtenir une vue d’ensemble, consultez la rubrique [Activer l’authentification unique dans un complément Office](sso-in-office-add-ins.md).
Cet article vous explique tout au long du processus d’activation de l' sign-on unique (SSO) dans un add-in créé avec ASP.NET.

> [!NOTE]
> Pour un article similaire concernant un complément basé sur Node.js, consultez [Création d’un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Conditions préalables

* Visual Studio 2019 ou version ultérieure.

* [Outils de développement Office](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* Au moins quelques fichiers et dossiers stockés sur OneDrive Entreprise dans votre abonnement Microsoft 365.

* Un abonnement Microsoft Azure. Ce complément requiert Azure Active Directory (AD). Azure AD fournit des services d’identité que les applications utilisent à des fins d’authentification et d’autorisation. Un abonnement d’évaluation peut être obtenu sur le site de [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

Clonez ou téléchargez le référentiel sur [Complément Office ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).

> [!NOTE]
> Il existe deux versions de l’échantillon :
>
> * Le dossier **Before** est un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. Les sections suivantes de cet article vous guident tout au long de la procédure d’exécution de cette dernière.
> * La version **Complète** de l’échantillon s’apparente au complément obtenu si vous aviez terminé les procédures de cet article, sauf que le projet final comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, suivez simplement les instructions de cet article, mais remplacez « Avant » par « Finale » et ignorez les sections **Code côté client** et **Code côté serveur**.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Enregistrez le complément avec le point de terminaison Azure AD v2.0

1. Accédez à la page [portail Azure : enregistrement des applications](https://go.microsoft.com/fwlink/?linkid=2083908) pour enregistrer votre application.

1. Connectez-vous ***avec les informations d’identification*** d’administrateur à votre location Microsoft 365. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez **Nouvelle inscription**. Sur la page **Inscrire une application**, définissez les valeurs comme suit.

    * Définissez le **Nom** sur `Office-Add-in-ASPNET-SSO`.
    * Définissez les **Types de comptes pris en charge** à **Comptes dans un annuaire organisationnel (comptes Azure AD Directory multi-locataires) et les comptes personnels Microsoft (par ex. Skype, Xbox)**. (Si vous voulez que le complément soit utilisable uniquement par les utilisateurs de l’organisation où vous l’enregistrez, vous pouvez choisir **Comptes dans cet annuaire d’organisation uniquement...** à la place, mais vous devrez suivre quelques étapes de configuration supplémentaires. Si vous souhaitez en savoir plus, veuilles consulter **Configuration de pour un seul locataire**.)
    * Dans la section **redirection d’URI**, assurez-vous que **Web** est sélectionnée dans la liste déroulante, puis définissez l’URI sur` https://localhost:44355/AzureADAuth/Authorize`.
    * Choisissez **Inscrire**.

1. Dans la page **Office-Add-in-ASPNET-SSO,** copiez et enregistrez les valeurs de **l’ID d’application (client)** et de l’ID d’annuaire **(client).** Vous utiliserez les deux plus tard.

    > [!NOTE]
    > Cet ID d’application **(client)** est la valeur « audience » lorsque d’autres applications, telles que l’application cliente Office (par exemple, PowerPoint, Word, Excel), recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Sous **Gérer** sélectionnez **Certificats et clés secrètes**. Sélectionnez le bouton **Nouveau secret client**. Entrer une valeur pour **Description**, puis sélectionnez une option appropriée pour **Expire le** puis **Ajouter**. *Copier la valeur secrète client immédiatement et enregistrez-la avec l’ID d’application* avant de continuer car vous en aurez besoin dans une procédure plus loin.

1. Sélectionnez **Exposer une API** sous **Gérer**. Sélectionnez le lien **Définir** pour générer l’URI de l’ID d’application sous la forme « api://$App ID GUID$ », où $App ID GUID$ est l’**ID de l’application (client)**. Insérez `localhost:44355/` (Notez la barre oblique « / » ajoutée à la fin) après la `//` et avant le GUID. La forme de l’ID entier doit être `api://localhost:44355/$App ID GUID$`; par exemple`api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Sélectionnez **Enregistrer** dans la boîte de dialogue.

1. Sélectionnez le bouton **Ajouter une étendue**. Dans le volet qui s’ouvre, entrez `access_as_user` en tant que **nom de l’étendue**.

1. Donnez la valeur **Administrateurs et utilisateurs** à **Qui peut donner son consentement ?** .

1. Remplissez les champs pour configurer les invites de consentement de l’administrateur et de l’utilisateur avec des valeurs appropriées pour l’étendue qui permet à l’application cliente Office d’utiliser les API web de votre application avec les mêmes droits que `access_as_user` l’utilisateur actuel. Suggestions :

    * **Nom complet du consentement de l’administrateur**: Office peut agir en tant qu’utilisateur.
    * **Description consentement administrateur** : activez Office pour qu’il appelle les API de complément web avec les mêmes droits que l’utilisateur actuel.
    * **Nom d’affichage du consentement de l’utilisateur**: Office peut agir en votre nom.
    * **Description du consentement de** l’utilisateur : permettre à Office d’appeler les API web du add-in avec les mêmes droits que vous.

1. Vérifiez que **State** est défini comme **Activé**.

1. Sélectionnez **Ajouter une étendue**.

    > [!NOTE]
    > La partie domaine du **Nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’URI d’ID d’application définie à l’étape précédente avec `/access_as_user`ajouté à la fin, par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Dans la section **Applications client autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé.

    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office sur le web)
    * `08e18876-6177-487e-b8b5-cf950c1e598c` (Office sur le web)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook sur le web)

    Pour chaque ID, procédez comme suit :

    a. Sélectionnez le bouton **Ajouter une application client** puis, dans le volet qui s’ouvre, définissez l’ID Client pour le GUID respectif et cochez la case pour `api://localhost:44355/$App ID GUID$/access_as_user`.

    b. Sélectionnez **Ajouter une application**.

1. Sélectionnez **Autorisations API** sous **Gestion**, puis sélectionnez **Ajouter une autorisation**. Dans le volet qui s’ouvre, sélectionnez **Microsoft Graph**, puis **Autorisations déléguées**.

1. Utilisez la zone de recherche **Sélectionnez les autorisations** pour rechercher les autorisations dont votre complément a besoin. Sélectionnez les éléments suivants. Seule la première est réellement requise par votre module lui-même ; mais `profile` l’autorisation est requise pour que l’application Office obtienne un jeton pour votre application web de add-in. (Seuls Files.Read.All et profil sont réellement nécessaires au complément. Vous devez demander les deux autres, car la bibliothèque MSAL.NET en a besoin.)

    * Files.Read.All
    * offline_access
    * openid
    * profil

    > [!NOTE]
    > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Une bonne pratique consiste à demander uniquement les autorisations dont vous avez besoin. Ainsi, nous vous recommandons de désactiver la case à cocher de cette autorisation si votre complément n’en a pas réellement besoin.

1. Activez la case à cocher pour chaque autorisation telle qu’elle apparaît. Après avoir sélectionné les autorisations dont votre complément a besoin, sélectionnez le bouton **Ajouter des autorisations** situé en bas du panneau.

1. Sur la même page, sélectionnez le bouton **Accorder l’autorisation d’administrateur pour [nom du client]**, puis **Accepter** pour la confirmation qui s’affiche.

    > [!NOTE]
    > Une fois que vous avez choisi **Accorder le consentement d’administrateur pour [nom du locataire]**, vous pouvez voir un message de bannière vous invitant à réessayer dans quelques minutes afin de pouvoir construire l’invite d’autorisation. Si c’est le cas, vous pouvez commencer à travailler sur la section suivante, mais n’oubliez pas de revenir au portail et **_d’appuyer sur ce bouton_**!

## <a name="configure-the-solution"></a>Configurer la solution

1. À la racine du dossier **Before**, ouvrez le fichier (.sln) solution dans **Visual Studio**. Cliquez avec le bouton droit sur le nœud supérieur de l’**Explorateur de solutions** (le nœud solution, et non l’un des nœuds de projet), puis sélectionnez **Définir les projets de démarrage**.

1. Sous **Propriétés communes**, sélectionnez **Projet de démarrage**, puis **Plusieurs projets de démarrage**. Assurez-vous que l’**Action** pour les deux projets est définie sur **Démarrer**, et que le projet qui se termine par « ...WebAPI » apparaît en premier dans la liste. Fermez la boîte de dialogue.

1. De retour **dans l’Explorateur** de solutions, sélectionnez (ne cliquez pas avec le bouton droit) le projet **Office-Add-in-ASPNET-SSO-WebAPI.** Le volet **Propriétés** s’ouvre. Assurez-vous que **SSL activé** est **Vrai**. Vérifiez que l’**URL SSL** est `http://localhost:44355/`.

1. Dans « web.config », utilisez les valeurs que vous avez copiées dans le version précédente. Configurez les **Ida:ClientID** et **Ida:Audience** à votre **ID d’application (client)**, puis configurez **Ida:Password** sur votre code secret client. Définissez également **ida:Domain** `http://localhost:44355` sur (aucune barre oblique « / » à la fin). 

    > [!NOTE]
    > L’ID d’application **(client)** est la valeur « audience » lorsque d’autres applications, telles que l’application cliente Office (par exemple, PowerPoint, Word, Excel), recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Si vous n’avez pas choisi « Comptes dans ce répertoire d’organisation uniquement » pour **TYPES DE COMPTES PRIS EN CHARGE** lorsque vous avez enregistré le complément, enregistrez et fermez le fichier web.config. Dans le cas contraire, enregistrez-le et laissez-le ouvert.

1. Toujours dans l’Explorateur de **solutions,** choisissez le projet **Office-Add-in-ASPNET-SSO,** ouvrez le fichier manifeste de la solution « Office-Add-in-ASPNET-SSO.xml », puis faites défiler vers le bas du fichier. Juste au-dessus de la balise de fin `</VersionOverrides>`, vous trouverez le balisage suivant :

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Remplacez l’espace réservé « $application_GUID here$ » *aux deux endroits* du balisage par l’ID d’application que vous avez copiée lorsque vous avez inscrit votre complément. Les signes « $ » ne faisant pas partie de l’ID, vous ne devez pas les inclure. C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.

  > [!NOTE]
  > La valeur de la **ressource** est l’**URI de l’ID d’application** que vous avez défini lors de l’inscription du complément. La section **Étendues** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.

1. Enregistrez et fermez le fichier.

### <a name="setup-for-single-tenant"></a>Configuration d’un seul locataire

Si vous avez choisi « Comptes dans ce répertoire d’organisation uniquement » pour **TYPES DE COMPTES PRIS EN CHARGE** lorsque vous avez enregistré le complément, vous devez suivre ces étapes de configuration supplémentaires :

1. Revenez au portail Azure et ouvrez le volet **vue d’ensemble** de l’inscription du complément. Copiez l’**ID de répertoire (client)**.

1. Dans le fichier Web. config, remplacez le « Common » par la valeur de **Ida:Authority** avec le GUID que vous avez copié à l’étape précédente. Lorsque vous avez terminé, la valeur doit ressembler à ceci : `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.

1. Enregistrez et fermez le fichier web.config.

## <a name="code-the-client-side"></a>Code côté client

1. Ouvrez le fichier HomeES6.js dans le dossier **Scripts**. Il contient déjà du code :

    * Un polyfill qui affecte l’objet Office. promesse à l’objet fenêtre globale pour que le complément puisse s’exécuter lorsque Office utilise Internet Explorer pour l’interface utilisateur. (Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).)
    * Une affectation à la méthode `Office.initialize` qui affecte elle-même un gestionnaire à l’événement ClickButton `getGraphAccessTokenButton`.
    * Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.
    * Une méthode `logErrors` qui consigne dans la console les erreurs qui ne sont pas destinées à l’utilisateur final.
    * Code qui implémente le système d’autorisation de repli que le complément utilisera dans les scénarios où l’authentification unique n’est pas prise en charge ou a provoqué une erreur.

1. En dessous de l’affectation au `Office.initialize`, ajoutez le code ci-dessous. Tenez compte des informations suivantes à propos de ce code :

    * La gestion des erreurs dans le complément tente parfois automatiquement d’obtenir un jeton d’accès une deuxième fois, à l’aide d’un autre jeu d’options. La variable de compteur `retryGetAccessToken` permet de s’assurer que l’utilisateur ne tente pas de manière répétée d’obtenir un jeton sans y parvenir.
    * La fonction `getGraphData` est définie avec le mot clé ES6 `async`. L’utilisation de la syntaxe ES6 simplifie l’utilisation de l’API d’authentification unique dans les compléments Office. Il s’agit du seul fichier dans la solution qui utilise une syntaxe non prise en charge par Internet Explorer. Nous plaçons « ES6 » dans le nom du fichier comme rappel. La solution utilise le transpondeur tsc pour transpiler ce fichier en ES5, afin que le complément puisse être exécuté lorsque Office utilise Internet Explorer pour l’interface utilisateur. (Consultez le fichier tsconfig.json dans la racine du projet.)

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. Ajoutez la fonction suivante après la fonction `getGraphData`. Notez que vous créez la fonction `handleClientSideErrors` dans une étape ultérieure.

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graph and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```

1. Remplacez `TODO 1` par ce qui suit. Tenez compte du code suivant :

    * `getAccessToken` indique à Office d’obtenir un jeton de démarrage à partir d’Azure AD et de revenir au complément.
    * `allowSignInPrompt` indique à Office d’inviter l’utilisateur à se connecter si l’utilisateur n’est pas encore connecté à Office.
    * `allowConsentPrompt` indique à Office d’inviter l’utilisateur à donner son consentement pour permettre au add-in d’accéder au profil AAD de l’utilisateur, si le consentement n’a pas déjà été accordé. (L’invite qui en résulte *n’autorise pas* l’utilisateur à consentir à des étendues Microsoft Graph.)
    * `forMSGraphAccess` indique à Office que le complément envisage de permuter le jeton d'amorçage d’un jeton d’accès à Microsoft Graph (au lieu d’utiliser simplement le jeton d'amorçage comme jeton ID utilisateur). La configuration de cette option permet à Office d’annuler le processus d’acquisition d’un jeton d'amorçage (et de renvoyer le code d’erreur 13012) si l’administrateur du locataire de l’utilisateur n’a pas accordé le consentement du complément. Le code côté client du complément peut répondre au 13012 en branchant un système d’autorisation de secours. Si l’utilisateur n’est pas utilisé et que l’administrateur n’a pas donné son consentement, le jeton d’a bootstrap est renvoyé, mais la tentative de l’échanger avec le flux « de la part de » entraînerait une `forMSGraphAccess` erreur. Par conséquent, l’option `forMSGraphAccess` permet au complément de brancher rapidement vers le système de secours.
    * Vous créez la fonction `getData` dans une étape ultérieure.
    * Le paramètre `/api/values` est l’URL d’un contrôleur côté serveur qui transforme l’échange de jeton et utilise le jeton d’accès qu’il renvoie pour appeler Microsoft Graph.

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. Ajoutez la fonction suivante après la fonction `getGraphData`. Tenez compte du code suivant :

    * Il est utilisé par les systèmes d’authentification unique et de secours.
    * Le paramètre `relativeUrl` est un contrôleur côté serveur.
    * Le paramètre `accessToken` peut être un jeton d’amorçage ou un jeton d’accès complet.
    * Le `writeFileNamesToOfficeDocument` fait déjà partie du projet.
    * Vous créez la fonction `handleServerSideErrors` dans une étape ultérieure.

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a>Gérer les erreurs côté client

1. Sous la fonction `getData`, ajoutez la fonction suivante. Veuillez noter que `error.code` est un nombre, généralement compris dans la plage 13xxx.

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. Remplacez `TODO 2` par le code suivant. Pour plus d’informations sur ces erreurs, reportez-vous à [Résoudre les problèmes liés à SSO dans les compléments Office](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the web.
        showResult(["Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. Remplacez `TODO 3` par le code suivant. Pour toutes les autres erreurs, le complément se branche au système d’autorisation de secours. Pour plus d’informations sur ces erreurs, voir [Résoudre les problèmes d' ssO dans les add-ins Office.](troubleshoot-sso-in-office-add-ins.md) Dans ce module, le système de base ouvre une boîte de dialogue qui exige que l’utilisateur se connecte, même si l’utilisateur l’est déjà.

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>Gérer les erreurs côté serveur

1. Sous la fonction `handleClientSideErrors`, ajoutez la fonction suivante.

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. Remplacez `TODO 4` par ce qui suit. À propos de ce code, Notez que des classes d’erreur ASP.NET ont été créées avant d’être telles que l’authentification multi-facteur. Dans le cadre de la façon dont la logique côté serveur gère les demandes pour un deuxième facteur d’authentification, l’erreur côté serveur envoyée au client a une propriété de **Message**, mais aucune propriété **ExceptionMessage** n’est disponible. Cependant, toutes les autres erreurs auront une propriété **ExceptionMessage**, pour que le code côté client doit analyser la réponse pour les deux. L’une ou l’autre variable est non définie.

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. Remplacez `TODO 5` par ce qui suit. Lorsque Microsoft Graph exige un formulaire d’authentification supplémentaire, il envoie l’erreur AADSTS50076. Celle-ci inclut des informations sur la configuration requise supplémentaire dans la propriété **message les déclarations**. Pour gérer ce problème, le code effectue une deuxième tentative d’obtention du jeton d’amorçage, mais cette fois, il inclut la demande d’un facteur supplémentaire comme valeur de l’option `authChallenge`, ce qui indique à Azure AD d’inviter l’utilisateur à fournir toutes les formes requises d’authentification.

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. Remplacez `TODO 6` par ce qui suit.

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. Remplacez `TODO 7` par ce qui suit. Notez que, dans de rares cas, le jeton de démarrage n’a pas expiré lorsqu’il est validé par Office, mais arrive à expiration au moment où il est envoyé Azure AD pour l’échange. Azure AD enverra une réponse incluant l’erreur AADSTS500133. Dans ce cas, le code rappelle l’API de l’authentification unique (sauf une fois). Cette fois-ci, Office renvoie un nouveau jeton d’amorçage non expiré.

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Remplacez `TODO 8` par ce qui suit.

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. Enregistrez le fichier.

## <a name="code-the-server-side"></a>Code côté serveur

### <a name="configure-the-owin-middleware"></a>Configurer les intergiciels OWIN

1. Ouvrez le fichier Startup.cs à la racine du projet **Office-Add-in-ASPNET-SSO-WebAPI** et ajoutez la méthode suivante à la classe de **démarrage**. Notez que vous créez la méthode `ConfigureAuth` dans une étape ultérieure.

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. Enregistrez et fermez le fichier.

1. Cliquez avec le bouton droit de la souris sur le dossier **App_Start**, puis sélectionnez **Ajouter > Classe**.

1. Dans la boîte de dialogue **Ajouter un nouvel élément** nommez le fichier **Startup.Auth.cs**, puis cliquez sur **Ajouter**.

1. Raccourcissez le nom de l’espace de noms dans le nouveau fichier `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Vérifiez que toutes les instructions `using` suivantes se trouvent en haut du fichier.

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Ajoutez le mot clé `partial` à la déclaration de la classe `Startup`, si ce n’est pas déjà fait. Elle doit ressembler à ceci :

    `public partial class Startup`

1. Ajoutez la méthode suivante à la classe `Startup`. Cette méthode spécifie comment l’intergiciel OWIN valide les jetons d’accès qui lui sont transmis à partir de la méthode `getData` dans le fichier Home.js côté client. Le processus d’autorisation est déclenché chaque fois qu’un point de terminaison Web API décoré avec l’attribut `[Authorize]` est appelé.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. Remplacez le `TODO 1` par ce qui suit. Tenez compte des informations suivantes :

    * Le code demande à OWIN de s’assurer que l’audience spécifiée dans le jeton d’a bootstrap provenant de l’application Office doit correspondre à la valeur spécifiée dans le web.config.
    * Les comptes Microsoft ont un GUID d’émetteur différent de n’importe quel GUID de client d’organisation. Ainsi, pour prendre en charge les deux types de comptes, nous ne validons pas l’émetteur.
    * Si `SaveSigninToken` vous `true` paramètrez ce paramètre, OWIN enregistre le jeton d’approvisionnement brut à partir de l’application Office. Le complément en a besoin pour obtenir un jeton d’accès à Microsoft Graph avec le flux « de la part de ».
    * Les étendues ne sont pas validées par l’intergiciel OWIN. Les étendues du jeton d’amorçage, qui doivent inclure `access_as_user`, sont validées dans le contrôleur.

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. Remplacez `TODO 2` par ce qui suit. Tenez compte des informations suivantes :

    * La méthode `UseOAuthBearerAuthentication` est appelée au lieu de la méthode `UseWindowsAzureActiveDirectoryBearerAuthentication` plus courante, car cette dernière n’est pas compatible avec le point de terminaison Azure AD V2.
    * L’URL transmise à la méthode est l’endroit où l’intermédiaire OWIN obtient des instructions pour obtenir la clé dont il a besoin pour vérifier la signature sur le jeton d’a bootstrap reçu de l’application Office. Le segment d’autorité de l’URL provient du fichier web.config. Il s’agit soit de la chaîne « commun », soit d’un GUID pour un complément à un seul locataire.

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. Enregistrez et fermez le fichier.

### <a name="create-the-apivalues-controller"></a>Créer le contrôleur /api/values

1. Ouvrez le fichier **Controllers\ValueController.cs**. Ce contrôleur est utilisé lorsque le système d’authentification unique a correctement obtenu un jeton d’amorçage. Il n’est pas utilisé dans le cadre du système d’autorisation de secours. Ce système utilise l'AzureADAuthController, qui a été créé pour vous.

1. Vérifiez que les instructions `using` suivantes se trouvent en haut du fichier.

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. Juste au-dessus de la ligne qui déclare `ValuesController`, ajoutez l’attribut `[Authorize]`. Cela permet de s’assurer que votre complément exécutera le processus d’autorisation que vous avez configuré dans la dernière procédure chaque fois qu’une méthode de contrôleur est appelée. Seuls les appelants avec un jeton d’accès valide à votre complément peuvent ainsi appeler les méthodes du contrôleur.

1. Ajoutez la méthode suivante à `ValuesController`. Vous remarquerez que la valeur renvoyée est `Task<HttpResponseMessage>` et non `Task<IEnumerable<string>>`, laquelle serait plus courante pour une méthode `GET api/values`. Il s’agit d’un effet secondaire de ce fait que la logique d’autorisation OAuth doit se trouver dans le contrôleur, plutôt que dans un filtre ASP.NET. Certaines conditions d’erreur dans cette logique nécessitent qu’un objet de réponse HTTP soit envoyé au client du complément.

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. Remplacez `TODO1` par le code suivant pour confirmer que les étendues spécifiées dans le jeton incluent `access_as_user`. Notez que le deuxième paramètre de la méthode `SendErrorToClient` est un objet d’**Exception**. Dans ce cas, le code transmet `null` car même l’objet **Exception** bloque l’inclusion de la propriété **Message** dans la réponse HTTP qui est générée.


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. Remplacez `TODO 2` par le code suivant pour assembler toutes les informations nécessaires pour obtenir un jeton pour Microsoft Graph à l’aide du flux « de la part de ». Tenez compte du code suivant :

    * Votre add-in ne joue plus le rôle d’une ressource (ou d’une audience) à laquelle l’application Office et l’utilisateur ont besoin d’accéder. Désormais, il est lui-même un client qui a besoin d’accéder à Microsoft Graph. `ConfidentialClientApplication` est l’objet de « contexte client » MSAL.
    * À partir de MSAL.NET 3. x. x, le `bootstrapContext` est simplement le jeton d’amorçage.
    * L’autorité provient du fichier web.config. Il s’agit soit de la chaîne « commun », soit d’un GUID pour un complément à un seul locataire.
    * MSAL requiert les étendues `openid` et `offline_access` pour fonctionner, mais il génère une erreur si votre code les demande de façon redondante. Il lève également une erreur si votre code demande, qui est vraiment utilisé uniquement lorsque l’application cliente Office obtient le jeton à `profile` l’application web de votre add-in. Seul `Files.Read.All` est demandé explicitement.

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri(ConfigurationManager.AppSettings["ida:Domain"])
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. Remplacez `TODO 3` par le code suivant. Tenez compte des informations suivantes :

    * La méthode `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` recherchera tout d’abord dans le cache MSAL, c’est-à-dire en mémoire, un jeton d’accès correspondant. Uniquement s’il n’existe pas, elle lance le flux « de la part de » avec le point de terminaison Azure AD V2.
    * Les exceptions qui ne sont pas de type `MsalServiceException` ne sont intentionnellement pas capturées afin d’être propagées au client sous la forme de messages `500 Server Error`.

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. Remplacez `TODO 3a` par le code suivant. Tenez compte du code suivant :

    * Si l’authentification multifacteur est requise par la ressource Microsoft Graph et que l’utilisateur ne l'a pas encore fournie, Azure AD renvoie « 400 : emande incorrecte » avec l’erreur `AADSTS50076` et une propriété **Claims**. MSAL génère une exception **MsalUiRequiredException** (qui hérite de **MsalServiceException**) avec ces informations.
    * La valeur de la propriété **Claims** doit être transmise au client qui doit la transmettre à l’application Office, qui l’inclut ensuite dans une demande de nouveau jeton d’a bootstrap. Azure AD demandera à l’utilisateur d’accepter tous les formulaires d’authentification requis.
    * Les API qui créent des réponses HTTP à partir d’exceptions ne connaissent pas la propriété **Claims**, donc ils ne l’incluent pas dans l’objet de la réponse. Nous devons créer manuellement un message qui l’inclut. Une propriété **Message** personnalisé, cependant, bloque la création d’une propriété **ExceptionMessage**, afin que la seule façon de communiquer l’ID d’erreur `AADSTS50076` au client est de l’ajouter à la propriété **Message** personnalisée. JavaScript dans le client devra découvrir si une réponse a une propriété **Message** ou **ExceptionMessage**, afin qu’il sache laquelle lire.
    * Le message personnalisé est au format JSON pour que le code JavaScript côté client puisse l’analyser avec des méthodes d’objet `JSON` JavaScript connues.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. Remplacez `TODO 3b` par le code suivant. Tenez compte du code suivant :

    * Si l’appel à Azure AD contenait au moins une étendue (autorisation) pour laquelle ni l’utilisateur, ni un administrateur client a consenti (ou pour laquelle le consentement a été révoqué), Azure AD renvoie « 400 demande incorrecte » avec une erreur `AADSTS65001` MSAL génère une exception **MsalUiRequiredException** avec ces informations.
    * Si l’appel à Azure AD contenait au moins une étendue non reconnue par Azure AD, AAD renvoie « 400 Demande incorrecte » avec l’erreur `AADSTS70011`. MSAL génère une exception **MsalUiRequiredException** avec ces informations.
    * La description entière est incluse, car l’erreur 70011 est renvoyée dans d’autres conditions et elle doit être gérée dans ce complément uniquement lorsqu’elle indique une étendue non valide.
    * L’objet **MsalUiRequiredException** est transmis à `SendErrorToClient`. Cela permet de garantir qu’une propriété **ExceptionMessage** qui contient les informations d’erreur est incluse dans la réponse HTTP.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. Remplacez `TODO 3c` par le code suivant pour gérer toutes les autres **MsalServiceException** s. Comme indiqué précédemment,

    ```csharp
    else
    {
        throw e;
    }
    ```

1. Remplacez `TODO 4` par le code suivant. La méthode `GraphApiHelper.GetOneDriveFileNames`, créée pour vous, effectue la demande de données à Microsoft Graph et inclut le jeton d’accès.

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. Enregistrez et fermez le fichier.

## <a name="run-the-solution"></a>Exécutez la solution

1. Ouvrez le fichier de solution Visual Studio.
1. Dans le menu **Générer**, sélectionnez **Nettoyer la solution**. Une fois l’opération terminée, ouvrez de nouveau le menu **Build**, puis sélectionnez **Générer la solution**.
1. Dans l’**Explorateur de solutions**, sélectionnez le nœud de projet **Office-Add-in-ASPNET-SSO** (et non le projet dont le nom se termine par « WebAPI »).
1. Dans le volet **Propriétés**, ouvrez la liste déroulante **Document de départ**, puis choisissez l’une des trois options (Excel, Word ou PowerPoint).

    ![Choisissez l’application cliente Office souhaitée : Excel, PowerPoint ou Word](../images/SelectHost.JPG)

1. Appuyez sur la touche F5.
1. Dans l’application Office, sur le ruban **Accueil**, sélectionnez **Afficher le complément** dans le groupe **ASP.NET SSO** pour ouvrir le complément du panneau des tâches.
1. Cliquez sur le bouton **Obtenir des noms de fichier OneDrive**. Si vous êtes connecté à Office avec un compte Microsoft 365 Éducation ou de travail, ou un compte Microsoft, et que l' sso fonctionne comme prévu, les 10 premiers noms de fichiers et de dossiers dans votre OneDrive Entreprise sont affichés dans le volet Office. Si vous n’êtes pas connecté ou si vous êtes dans un scénario qui ne prend pas en charge SSO ou si l’authentification unique ne fonctionne pas pour une raison quelconque, vous serez invité à vous connecter. Une fois connecté, les noms de fichier et de dossier s’affichent.

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>Mise à jour du add-in lors de la mise en transit et de la production

Comme tous les applications web Office, lorsque vous êtes prêt à passer à un serveur intermédiaire ou de production, vous devez mettre à jour le domaine dans le manifeste avec `localhost:44355` le nouveau domaine. De même, vous devez mettre à jour le domaine dans web.config fichier.

Étant donné que le domaine apparaît dans l’inscription AAD, vous devez mettre à jour cette inscription pour utiliser le nouveau domaine à la place de l’endroit `localhost:44355` où il apparaît.
