---
title: Création d’un complément Office Node.js qui utilise l’authentification unique
description: Apprenez à créer un complément basé sur Node.js utilisant l’authentification unique Office.
ms.date: 01/16/2020
localization_priority: Priority
ms.openlocfilehash: 562351011ef8aaf2ba936cceea862ebfec888b11
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265451"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>Créer un complément Office Node.js qui utilise l’authentification unique (aperçu)

Les utilisateurs peuvent se connecter à Office et votre complément Web Office peut tirer parti de cette procédure de connexion pour autoriser les utilisateurs à accéder à votre complément et à Microsoft Graph sans obliger les utilisateurs à se connecter une deuxième fois. Pour obtenir une vue d’ensemble, consultez [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md).

Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré à Node.js et Express. Pour voir un article similaire sur un complément basé sur ASP.NET, reportez-vous à [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md).

> [!NOTE]
> Au lieu de suivre les étapes décrites dans cet article, vous pouvez utiliser le générateur d'Yeoman pour créer un complément Office compatible avec l’authentification unique, Node.js. Le générateur d’Yeoman simplifie le processus de création d’un complément avec authentification unique en automatisant les étapes nécessaires pour configurer l’authentification unique dans Azure et la génération du code nécessaire pour qu’un complément utilise l’authentification unique. Pour plus d'informations, consultez [Démarrage rapide de l'authentification unique](../quickstarts/sso-quickstart.md).

## <a name="prerequisites"></a>Conditions préalables

* [Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))

* [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

* TypeScript version 3.6.2 ou ultérieure.

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* Éditeur de code. Nous vous recommandons Visual Studio Code.

* Au moins des fichiers et classeurs sont stockés sur OneDrive Entreprise dans votre abonnement Office 365.

* Un abonnement Microsoft Azure. Ce complément requiert Azure Active Directory (AD). Azure AD fournit des services d’identité que les applications utilisent à des fins d’authentification et d’autorisation. Un abonnement d’évaluation peut être obtenu sur le site de [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

1. Clonez ou téléchargez le référentiel sur [Complément Office NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).

    > [!NOTE]
    > Il existe trois versions de l’échantillon :  
    > * Le dossier **Before** est un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. Les sections suivantes de cet article vous guident tout au long de la procédure d’exécution de cette dernière.
    > * La version **Complète** de l’échantillon s’apparente au complément obtenu si vous aviez terminé les procédures de cet article, sauf que le projet final comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, suivez simplement les instructions de cet article, mais remplacez « Avant » par « Finale » et ignorez les sections **Code côté client** et **Code côté serveur**.
    > * La version **SSOAutoSetup** est un exemple complet qui permet d’automatiser la plupart des étapes d’inscription du complément avec Azure AD et sa configuration. Utilisez cette version si vous voulez rapidement afficher un complément opérationnel avec SSO. Suivez simplement les étapes décrites dans le fichier Lisez-moi du dossier. Nous vous recommandons, à un certain point, de suivre les étapes d’inscription et de configuration manuelles décrites dans cet article pour mieux comprendre la relation entre Azure AD et un complément. 


1. Ouvrez une invite de commandes dans le dossier **auparavant**.

1. Saisissez `npm install`dans la console pour installer toutes les dépendances détaillées dans le fichier package.json.

1. Exécutez la commande `npm run install-dev-certs`. Sélectionnez **Oui** lorsque vous êtes invité à installer le certificat.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Enregistrez le complément avec le point de terminaison Azure AD v2.0

1. Accédez à la page [portail Azure : enregistrement des applications](https://go.microsoft.com/fwlink/?linkid=2083908) pour enregistrer votre application.

1. Connectez-vous à votre client Office 365 en utilisant les informations d’identification d’***administrateur***. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez **Nouvelle inscription**. Sur la page **Inscrire une application**, définissez les valeurs comme suit.

    * Définissez le **Nom** sur `Office-Add-in-NodeJS-SSO`.
    * Définissez les **Types de comptes pris en charge** à **Comptes dans un annuaire organisationnel et les comptes personnels Microsoft (par ex. Skype, Xbox et Outlook.com)**.
    * Configurez **URI de redirection** vers` https://localhost:44355/dialog.html`.
    * Choisissez **Inscrire**.

1. Sur la page **Office-Add-in-NodeJS-SSO**, copiez et enregistrez les valeurs pour l’**ID de l’application (client)** et l’**ID de répertoire (client)**. Vous utiliserez les deux plus tard.

    > [!NOTE]
    > Cet ID a la valeur « audience » lorsque d’autres applications, telles que l’application hôte Office (par exemple, PowerPoint, Word, Excel) demandent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Sous **Gérer**, sélectionnez **Authentification**. Dans la section **Implict Grant**, activez les cases à cocher pour **Jeton d’accès** et **Jeton d’ID**. L’exemple dispose d’un système d’autorisation de secours qui est appelé lorsque l’authentification unique n’est pas disponible. Le système utilise le Flux implicite.

1. Sélectionnez **Enregistrer** en haut du formulaire.

1. Sélectionnez **Certificats et secrets** sous **Gérer**. Sélectionnez le bouton **Nouveau secret client**. Entrer une valeur pour **Description** puis sélectionnez une option appropriée pour **Expire le** puis **Ajouter**. *Copier la valeur secrète client immédiatement et enregistrez-la avec l’ID d’application* avant de continuer car vous en aurez besoin dans une procédure plus loin.

1. Sélectionnez **Exposer une API** sous **Gérer**. Sélectionnez le lien **Définir** pour générer l’URI de l’ID d’application sous la forme « api://$App ID GUID$ », où $App ID GUID$ est l’**ID de l’application (client)**. Insérez `localhost:44355/` (remarquez la barre oblique « / » ajoutée à la fin) entre les doubles barres obliques et le GUID. La forme de l’ID entier doit être `api://localhost:44355/$App ID GUID$`; par exemple`api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`. 

1. Sélectionnez le bouton **Ajouter une étendue**. Dans le volet qui s’ouvre, entrez `access_as_user` en tant que **nom de l’étendue**.

1. Donnez la valeur **Administrateurs et utilisateurs** à **Qui peut donner son consentement ?** .

1. Renseignez les champs pour configurer les invites de consentement des administrateurs et utilisateurs avec les valeurs appropriées pour l’étendue `access_as_user` qui permet à l’application Office hôte d’utiliser l’API web de votre complément avec les mêmes droits que l’utilisateur actuel. Suggestions :

    - **Titre consentement administrateur** : Office peut agir en tant qu’utilisateur.
    - **Description consentement administrateur** : activez Office pour qu’il appelle les API de complément web avec les mêmes droits que l’utilisateur actuel.
    - **Titre consentement utilisateur** : Office peut agir à votre place.
    - **Description consentement administrateur** : activez Office pour qu’il appelle les API de complément web avec les mêmes droits dont vous disposez.

1. Vérifiez que **State** est défini comme **Activé**.

1. Sélectionnez **Ajouter une étendue**.

    > [!NOTE]
    > La partie domaine du **Nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’URI d’ID d’application définie à l’étape précédente avec `/access_as_user`ajouté à la fin, par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Dans la section **Applications client autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office sur le web)
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office sur le web)

    Pour chaque ID, procédez comme suit :

    a. Sélectionnez le bouton **Ajouter une application client** puis, dans le volet qui s’ouvre, définissez l’ID Client pour le GUID respectif et cochez la case pour `api://localhost:44355/$App ID GUID$/access_as_user`.

    b. Sélectionnez **Ajouter une application**.

1. Sélectionnez **Autorisations API** sous **Gestion** et sélectionnez **Ajouter une autorisation**. Dans le volet qui s’ouvre, sélectionnez **Microsoft Graph**, puis **Autorisations déléguées**.

1. Utilisez la zone de recherche **Sélectionnez les autorisations** pour rechercher les autorisations dont votre complément a besoin. Sélectionnez les éléments suivants. Votre complément proprement dit ne requiert que la première. Mais l’autorisation `profile` est également requise pour que l’hôte Office puisse obtenir un jeton pour l’application web de votre complément.

    * Files.Read.All
    * profil

    > [!NOTE]
    > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Une bonne pratique consiste à demander uniquement les autorisations dont vous avez besoin. Ainsi, nous vous recommandons de désactiver la case à cocher de cette autorisation si votre complément n’en a pas réellement besoin.

1. Activez la case à cocher pour chaque autorisation telle qu’elle apparaît. Après avoir sélectionné les autorisations dont votre complément a besoin, sélectionnez le bouton **Ajouter des autorisations** situé en bas du panneau.

1. Sur la même page, sélectionnez le bouton **Accorder l’autorisation d’administrateur pour [nom du client]**, puis **Oui** pour la confirmation qui s’affiche.

## <a name="configure-the-add-in"></a>Configurer le complément

1. Ouvrez le dossier `\Begin` dans le projet cloné dans votre éditeur de code.

1. Ouvrez le fichier `.ENV` et utilisez les valeurs que vous avez précédemment copiées. Configurez la **CLIENT_ID** sur votre **ID d’application (client)** et attribuez la **CLIENT_SECRET** à votre clé secrète client. Les valeurs ne doivent **pas** se trouver entre des guillemets. Quand vous avez terminé, votre modèle doit ressembler à ce qui suit : 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. Ouvrez le fichier `\public\javascripts\fallbackAuthDialog.js`. Dans la `msalConfig`déclaration, remplacez l’espace réservé $application_GUID here$ par l’ID d’application que vous avez copié lorsque vous avez inscrit votre complément. Les valeurs ne doivent pas être entre guillemets.

1. Ouvrez le fichier manifeste de complément « manifest\manifest_local. xml », puis faites défiler la page jusqu’à la fin du fichier. Juste au-dessus de la `</VersionOverrides>`balise de fin, vous trouverez la marque suivante :

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Remplacez l’espace réservé « $application_GUID here$ » *aux deux endroits* du balisage par l’ID d’application que vous avez copiée lorsque vous avez inscrit votre complément. Les « $ » ne faisant pas partie de l’ID, vous ne devez pas les inclure. C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.

    > [!NOTE]
    > La valeur de la **ressource** est l’**URI de l’ID d’application** que vous avez défini lors de l’inscription du complément. La section **Étendues** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.

## <a name="code-the-client-side"></a>Code du côté client

### <a name="create-the-sso-logic"></a>Créer la logique SSO

1. Ouvrez le fichier `public\javascripts\ssoAuthES6.js` dans votre éditeur de code. Il possède déjà du code qui garantit que les promesses sont prises en charge, même dans Internet Explorer 11, et un appel `Office.onReady` pour attribuer un gestionnaire au bouton unique du complément.

    > [!NOTE]
    > Comme leur nom l’indique, ssoAuthES6.js utilise la syntaxe JavaScript ES6, car l’utilisation de `async` et de `await` illustre le mieux la simplicité de l’API SSO. Lorsque le serveur localhost est démarré, ce fichier est transpilé vers la syntaxe ES5 pour que l’exemple s’exécute dans Internet Explorer 11. 

1. Ajoutez le code suivant sous la méthode Office.onReady :

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exhange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. Remplacez `TODO 1` par le code suivant. Tenez compte du code suivant:

    - `OfficeRuntime.auth.getAccessToken` commande à Office d’obtenir un jeton de démarrage à partir d’Azure AD. Un jeton d’amorçage est semblable à un jeton d’ID, mais il possède une `scp` propriété (étendue) ayant la valeur `access-as-user`. Ce type de jeton peut être échangé par une application Web pour un jeton d’accès à Microsoft Graph.
    - Le paramétrage de l’option de `allowSignInPrompt`sur TRUE signifie que si aucun utilisateur n’est actuellement connecté à Office, Office ouvre une invite de connexion contextuelle.
    - Le paramétrage de l’option de `forMSGraphAccess` sur TRUE signale à Office que le complément envisage d’utiliser le jeton de démarrage pour obtenir un jeton d’accès à Microsoft Graph, plutôt que de l’utiliser simplement comme jeton d’ID. Si l’administrateur du client n’a pas accordé l’autorisation d’accès au complément dans Microsoft Graph, `OfficeRuntime.auth.getAccessToken` renvoie l’erreur **13012**. Le complément peut répondre en rétablissant un autre système d’autorisation, ce qui est nécessaire car Office peut uniquement inviter pour accepter le profil Azure AD de l’utilisateur, et non les étendues Microsoft Graph. Le système d’autorisation de secours oblige l’utilisateur à se reconnecter et l’utilisateur *peut* être invité à accepter les étendues de Microsoft Graph. Par conséquent, l’option `forMSGraphAccess` permet de s’assurer que le complément ne fera pas d’échange de jetons échouant en raison d’une absence d’autorisation. (ayant reçu votre consentement de la part de l’administrateur lors d’une étape précédente, ce scénario ne se produira pas pour ce complément. Mais l’option est tout de même incluse ici pour illustrer les pratiques recommandées.)

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true }); 
    ```

1. Remplacez `TODO 2` par le code suivant. Vous créerez la méthode `getGraphToken` lors d’une étape ultérieure.

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. Remplacez `TODO 3` par ce qui suit. Tenez compte du code suivant : 

    - Si le client Office 365 est configuré pour exiger l’authentification multifacteur, l' `exchangeResponse` inclut une propriété `claims` contenant des informations sur les facteurs supplémentaires requis. Dans ce cas, `OfficeRuntime.auth.getAccessToken` doit être rappelé avec l’option `authChallenge` configurée avec la valeur de la propriété revendications. Cela indique à AAD d’inviter l’utilisateur à accepter tous les formulaires d’authentification requis.

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. Remplacez `TODO 4` par ce qui suit. Tenez compte du code suivant : 

    - Vous créerez la méthode `handleAADErrors` lors d’une étape ultérieure. Les erreurs Azure AD sont renvoyées au client sous forme de réponses de code HTTP 200. Elles ne génèrent pas d’erreur et ne déclenchent donc pas le `catch`blocage de la`getGraphData` méthode.
    - Vous créerez la méthode `makeGraphApiCall` lors d’une étape ultérieure. Elle effectue un appel AJAX au point de terminaison MS Graph. Les erreurs sont interceptées dans le `.fail` rappel de cet appel, et non dans le bloc `catch` de la méthode `getGraphData`.

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. Remplacez `TODO 5` par le code suivant

    - Les erreurs de l’appel de `getAccessToken` auront une propriété `code` avec un numéro d’erreur généralement dans la plage 13xxx. Vous créerez la méthode `handleClientSideErrors` lors d’une étape ultérieure.
    - La méthode `showMessage` affiche le texte dans le volet Tâches.

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. En dessous de la méthode `getGraphData`, ajoutez la fonction suivante. Veuillez noter que `/auth` est une route Express côté serveur qui échange le jeton de démarrage avec Azure AD pour un jeton d’accès à Microsoft Graph.

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. En dessous de la méthode `getGraphToken`, ajoutez la fonction suivante. Veuillez noter que `error.code` est un nombre, généralement compris dans la plage 13xxx.

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```
1. Remplacez `TODO 6` par le code suivant. Pour plus d’informations sur ces erreurs, reportez-vous à [Résoudre les problèmes liés à SSO dans les compléments Office](troubleshoot-sso-in-office-add-ins.md). 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the Web.
        showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. Remplacez `TODO 7` par le code suivant. Pour plus d’informations sur ces erreurs, reportez-vous à [Résoudre les problèmes liés à SSO dans les compléments Office](troubleshoot-sso-in-office-add-ins.md). La fonction `dialogFallback` appelle le système d’autorisation de secours. Dans ce complément, le système de secours ouvre une boîte de dialogue demandant à l’utilisateur de se connecter, même si l’utilisateur l’est déjà, et utilise MSAL.js et le flux implicite pour obtenir un jeton d’accès à Microsoft Graph.

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. Sous la fonction `handleClientSideErrors`, ajoutez la fonction suivante. 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. Dans de rares cas, le jeton de démarrage qu’Office a mis en cache n’a pas expiré lorsqu’il est validé par Office, mais arrive à expiration au moment où il atteint Azure AD pour l’échange. Azure AD enverra une réponse incluant l’erreur **AADSTS500133**. Dans ce cas, le complément doit simplement appeler `getGraphData` de manière récursive. Le jeton de démarrage mis en cache étant arrivé à expiration, Office en reçoit un nouveau à partir d’Azure AD. Remplacez donc `TODO 8` par le code suivant. 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)       
    {
        getGraphData();
    }
    ```

1. Pour vous assurer que le complément n’entre pas dans une boucle infinie d’appels vers `getGraphData`, le complément doit effectuer un suivi du nombre de fois où `getGraphData` a été appelé et vérifier qu’il n’est pas appelé de façon récursive plusieurs fois. Par conséquent, créez une variable de compteur dans une étendue globale aux fonctions de `handleAADErrors` et `getGraphData`. Un bon emplacement pour les variables globales se trouve juste en dessous de l’appel de méthode `Office.onReady`.

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. Modifiez la structure `if` dans la méthode `handleAADErrors` de façon à ce qu’elle :

    - Incrémente le compteur juste avant qu’il n’appelle `getGraphData`.
    - Vérifie que `getGraphData` n’a pas déjà été appelé une deuxième fois. 

    La version finale de la structure `if` doit donc ressembler à ceci :

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Remplacez `TODO 9` par ce qui suit. 

    ```javascript
    else {                
        dialogFallback();
    }
    ```

1. Enregistrez et fermez le fichier.

### <a name="get-the-data-and-add-it-to-the-office-document"></a>Obtenir les données et les ajouter au document Office

1. Dans le dossier `public\javascripts`, créez un fichier appelé `data.js`.

1. Ajoutez la fonction suivante au fichier. Il s’agit de la fonction appelée par la fonction `getGraphData` lorsqu’elle a acquis un jeton d’accès pour Microsoft Graph. 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. Remplacez `TODO 10` par ce qui suit. Tenez compte du code suivant : 

    - Cet objet est le paramètre de la méthode `$.ajax`.
    - Le `/getuserdata` est une route Express sur le serveur du complément que vous créez au cours d’une étape ultérieure. Elle appellera un point de terminaison Microsoft Graph et inclura le jeton d’accès dans son appel. 

    ```javascript
    {
        type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. Remplacez `TODO11` par ce qui suit. Tenez compte du code suivant :

    - Le `writeFileNamesToOfficeDocument` insère les données de Graph dans le document Office. Il est défini dans le fichier `public\javascripts\document.js`. 
    - Si `writeFileNamesToOfficeDocument` renvoie une erreur, il commence par « Impossible d’ajouter des noms de fichiers au document ».

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () { 
        showMessage("Your data has been added to the document."); 
    })
    .catch(function (error) {        
        showMessage(error);
    });
    ```

1. Enregistrez et fermez le fichier.

## <a name="code-the-server-side"></a>Code du côté serveur

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a>Créer le routeur d’authentification et la logique d’échange de jetons

1. Ouvrez le fichier `routes\authRoute.js` et ajoutez la fonction d’itinéraire suivante juste en dessous des instructions `require` et au-dessus de l’instruction `module.exports`. Veuillez noter que le paramètre d’URL de `router.get` est'/'. Cet itinéraire étant défini dans un routeur qui gère toutes les requêtes HTTP pour l’URL « /auth », il gère toutes les demandes pour « /auth ». La fonction `getGraphToken` côté client créée précédemment appelle cet itinéraire.  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exhange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. Remplacez `TODO 12` par le code suivant.

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. Remplacez `TODO 13` par le code suivant. Tenez compte du code suivant: 

    - Il s’agit du début d’un long `else` bloc, mais la fermeture `}` n’est pas encore terminée car vous y ajouterez d’autres codes. 
    - La chaîne de `authorization` est « Porteur » suivi du jeton de démarrage, de sorte que la première ligne du bloc `else` attribue le jeton à la `jwt`. (« JWT » signifie « jeton Web JSON »).
    - Les deux valeurs `process.env.*` sont les constantes que vous avez attribuées lors de la configuration du complément. 
    - Le paramètre de formulaire `requested_token_use` est paramétré sur « On_behalt_of ». Cette option indique à Azure AD que le complément demande un jeton d’accès à Microsoft Graph à l’aide du flux On-Behalf-Of. Azure répond en validant que le jeton de démarrage, affecté au paramètre de formulaire `assertion`, a une propriété `scp` configurée sur `access-as-user`.
    - Le paramètre de formulaire `scope` est défini sur « Files.Read.All », qui est la seule étendue Microsoft Graph dont le complément a besoin.

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. Remplacez `TODO 14` par le code suivant, qui termine le bloc `else`. Tenez compte du code suivant :

    - Le `tenant` const est défini sur « commun », car vous avez configuré le complément en tant que multiclient lorsque vous l’avez inscrit avec Azure AD, en particulier lorsque vous configurez **types de compte pris en charge** pour **les comptes de n’importe quel annuaire d’organisation et les comptes Microsoft personnels (par exemple, Skype, Xbox, Outlook.com)**. Si vous avez en revanche choisi de prendre en charge uniquement les comptes figurant dans la même location Office 365 que le complément enregistré, `tenant` dans ce code serait défini sur le GUID du client. 
    - Si la requête POST ne génère pas d’erreur, la réponse d’Azure AD est convertie en JSON et envoyée au client. Cet objet JSON possède une propriété `access_token` à laquelle Azure AD a attribué un jeton d’accès à Microsoft Graph.

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();
            
            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. Enregistrez et fermez le fichier.

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a>Créer l’itinéraire qui permettra de récupérer les données à partir de Microsoft Graph

1. Ouvrez le fichier `app.js` dans la racine du projet. Juste en dessous de la route pour « /Dialog.html », ajoutez l’itinéraire suivant. Cet itinéraire est appelé par la fonction `makeGraphApiCall` que vous avez créée lors d’une étape précédente.

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. Remplacez `TODO 15` par ce qui suit. Tenez compte du code suivant :

    - L’appelant de cet itinéraire, `makeGraphApiCall`, a ajouté un jeton d’accès à Microsoft Graph à la demande HTTP en tant qu’en-tête nommé « access_token ».
    - La fonction de `getGraphData` est défini dans le fichier `msgraph-helper.js`. (il ne s’agit pas de la même fonction que la fonction `getGraphData` côté client que vous avez définie dans le fichier `ssoAuthES6.js`).
    - Le dernier paramètre pour `queryParamsSegment` est codé en dur. Si vous modifiez ce code dans un complément production et qu’une partie quelconque de `queryParamsSegment` provient d’une intervention de l’utilisateur, n’oubliez pas qu’il est purgé afin qu’il ne puisse pas être utilisé dans une attaque par injection d’en-tête de réponse.
    - Le code minimise les données qui doivent provenir de Microsoft Graph en spécifiant uniquement la propriété nécessaire (« nom ») et uniquement les 10 premiers noms de dossier ou de fichier.

    ```javascript
    const graphToken = req.get('access_token');    
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. Remplacez `TODO 16` par ce qui suit. Tenez compte du code suivant :

    - Si Microsoft Graph renvoie une erreur, un jeton non valide ou expiré par exemple, une propriété de code dans l’objet renvoyé est attribuée à un état HTTP (par exemple, 401). Le code relaie l’erreur vers le client. Elle sera interceptée dans le `.fail` rappel de `makeGraphApiCall`.
    - Les données Microsoft Graph incluent des métadonnées OData et des eTags dont le complément n’a pas besoin, de sorte que le code construit un nouveau groupe contenant uniquement le noms des fichiers à envoyer au client.

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. Enregistrez et fermez le fichier.

## <a name="run-the-project"></a>Exécutez le projet

1. Assurez-vous d’avoir des fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.

1. Ouvrez une invite de commandes dans la racine du dossier `\Complete`. 

1. Exécutez la commande `npm start`. 

1. Vous devez charger une version du complément dans une application Office (Excel, Word ou PowerPoint) pour le tester. Les instructions sont fonction de votre plateforme. Vous trouverez des liens vers des instructions sur [Charger une version du complément Office pour le tester](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. Dans l’application Office, sur le ruban **Accueil**, sélectionnez le bouton **Afficher le complément** dans le groupe **Node.js SSO** pour ouvrir le complément du panneau des tâches.

1. Cliquez sur le bouton **Obtenir des noms de fichier OneDrive**. Si vous êtes connecté à Office à l’aide d’un compte professionnel ou scolaire (Office 365) ou d’un compte Microsoft et que l’authentification unique fonctionne comme prévu, les 10 premiers noms de fichier et de dossiers dans votre espace OneDrive Entreprise sont insérés dans le document. (la première opération peut prendre jusqu’à 15 secondes). Si vous n’êtes pas connecté ou si vous êtes dans un scénario qui ne prend pas en charge SSO ou si l’authentification unique ne fonctionne pas pour une raison quelconque, vous serez invité à vous connecter. Une fois connecté, les noms de fichier et de dossier s’affichent.

> [!NOTE]
> Si vous étiez précédemment connecté à Office avec un ID différent et si certaines applications précédemment ouvertes Office le sont toujours, Office ne changera pas systématiquement votre identifiant même si cela semble être le cas. Dans ce cas, l’appel vers Microsoft Graph peut échouer ou des données de l’ID précédent peuvent être renvoyées. Afin d’éviter ce problème, veillez à *fermer toutes les autres applications Office* avant de cliquer sur **Obtenir des noms de fichiers OneDrive**.
