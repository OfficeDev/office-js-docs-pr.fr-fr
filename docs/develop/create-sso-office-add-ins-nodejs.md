---
title: Création d’un complément Office Node.js qui utilise l’authentification unique
description: ''
ms.date: 12/07/2018
localization_priority: Priority
ms.openlocfilehash: cf249e47709a325f22fc1fda49ee76b7a3357b4f
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701931"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>Créer un complément Office Node.js qui utilise l’authentification unique (aperçu)

Les utilisateurs peuvent se connecter à Office et votre complément Web Office peut tirer parti de cette procédure de connexion pour autoriser les utilisateurs à accéder à votre complément et à Microsoft Graph sans obliger les utilisateurs à se connecter une deuxième fois. Pour obtenir une vue d’ensemble, consultez [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md).

Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré à Node.js et Express. 

> [!NOTE]
> Pour voir un article similaire sur un complément basé sur ASP.NET, reportez-vous à [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Conditions préalables

* [Nœud et npm](https://nodejs.org/en/), version 6.9.4 ou ultérieure.

* [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

* TypeScript version 2.2.2 ou ultérieure.

* Office 365 (version par abonnement, également appelée « Démarrer en un clic »). Dernière version mensuelle et build du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1). Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

1. Clonez ou téléchargez le référentiel sur [Complément Office NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso). 

    > [!NOTE]
    > Il existe trois versions de l’échantillon :  
    > * Le dossier **Before** est un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. Les sections suivantes de cet article vous guident tout au long de la procédure d’exécution de cette dernière. 
    > * La version **Finale** de l’échantillon s’apparente au complément que vous auriez si vous terminiez les procédures de cet article, sauf que le projet terminé comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, suivez simplement les instructions de cet article, mais remplacez « Avant » par « Finale » et ignorez les sections **Code côté client** et **Code côté serveur**.
    > * La version **mutualisée finale** est un échantillon final qui prend en charge l’architecture mutualisée. Si vous avez l’intention de prendre en charge des comptes Microsoft de différents domaines avec l’authentification unique, explorez cet exemple.
    >
    > _Quelle que soit la version que vous utilisez, vous devrez approuver un certificat pour l’hôte local. Consultez la note « IMPORTANT » dans le fichier Lisez-moi du référentiel._

2. Ouvrez une console Git Bash dans le dossier **Before**.

3. Saisissez `npm install` dans la console pour installer toutes les dépendances détaillées dans le fichier package.json.

4. Saisissez `npm run build ` dans la console pour générer le projet. 

    > [!NOTE]
    > Il se peut que vous voyiez certaines erreurs de construction indiquant que certaines variables sont déclarées mais pas utilisées. Ignorez ces erreurs. Elles représentent un effet secondaire du fait qu’il manque du code dans la version « Avant » de l’échantillon. Ce code sera ajouté ultérieurement.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Enregistrez le complément avec le point de terminaison Azure AD v2.0

Les instructions suivantes présentant un manière générique, vous pouvez les utiliser dans plusieurs emplacements. En lien avec ce article, procédez comme suit :
- Remplacez l’espace réservé **$ADD-IN-NAME$** par `“Office-Add-in-NodeJS-SSO`.
- Remplacez l’espace réservé **$FQDN-WITHOUT-PROTOCOL$** par `localhost:3000`.
- Lorsque vous spécifiez des autorisations dans la boîte de dialogue **Sélectionner les autorisations**, cochez les cases correspondant aux autorisations suivantes. Votre complément proprement dit ne requiert que la première. Mais l’autorisation `profile` est également requise pour que l’hôte Office puisse obtenir un jeton pour l’application web de votre complément.
    * Files.Read.All
    * profil

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a>Octroi du consentement administrateur pour le complément

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Configurer le complément

1. Dans votre éditeur de code, ouvrez le fichier src\server.ts. Près de la partie supérieure se trouve un appel à un constructeur d’une classe `AuthModule`. Il existe certains paramètres de chaîne dans le constructeur auxquels vous devez affecter des valeurs.

2. Pour la propriété `client_id`, remplacez l’espace réservé `{client GUID}` par l’ID d’application que vous avez enregistré lorsque vous avez inscrit le complément. Lorsque vous avez terminé, vous obtenez simplement un GUID entre guillemets simples. Il ne doit pas y avoir de caractère «{}».

3. Pour la propriété `client_secret`, remplacez l’espace réservé `{client secret}` par le secret de l’application que vous avez enregistré lorsque vous avez inscrit le complément.

4. Pour la propriété `audience`, remplacez l’espace réservé `{audience GUID}` par l’ID d’application que vous avez enregistré lorsque vous avez inscrit le complément. (La même valeur que celle affectée à la propriété `client_id`.)
  
3. Dans la chaîne affectée à la propriété `issuer`, vous verrez l’espace réservé *{O365 tenant GUID}*. Remplacez-le par l’ID de client Office 365. Pour obtenir de celui-ci, utilisez l’une des méthodes décrites dans [Trouver votre ID de client Office 365](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id). Lorsque vous avez terminé, la valeur de la propriété `issuer` doit ressembler à ceci :

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. Conservez les autres paramètres du constructeur `AuthModule` inchangés. Enregistrez et fermez le fichier.

1. Dans la racine du projet, ouvrez le fichier manifeste du complément « Office-Add-in-NodeJS-SSO.xml ».

1. Faites défiler vers le bas du fichier.

1. Juste au-dessus de la balise de fin `</VersionOverrides>`, vous trouverez le balisage suivant :

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Remplacez l’espace réservé « {application_GUID here} » *aux deux endroits* du balisage par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément. (Les « {} » ne font pas partie de l’ID ; vous ne devez pas les inclure.) C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.

    > [!NOTE]
    > * La valeur **Resource** correspond à l’**URI d’ID d’application** défini lorsque vous avez ajouté la plateforme d’API web à l’enregistrement du complément.
    > * La section **Scopes** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.

1. Enregistrez et fermez le fichier.

## <a name="code-the-client-side"></a>Code côté client

1. Ouvrez le fichier program.js dans le dossier **public**. Il contient déjà du code :

    * Une affectation à la méthode `Office.initialize` qui affecte elle-même un gestionnaire à l’événement ClickButton `getGraphAccessTokenButton`.
    * Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.
    * Une méthode `logErrors` qui consigne dans la console les erreurs qui ne sont pas destinées à l’utilisateur final.

11. En dessous de l’affectation au `Office.initialize`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :

    * La gestion des erreurs dans le complément tente parfois automatiquement d’obtenir un jeton d’accès une deuxième fois, à l’aide d’un autre jeu d’options. La variable de compteur `timesGetOneDriveFilesHasRun` et la variable d’indicateur `triedWithoutForceConsent` et `timesMSGraphErrorReceived` permettent de s’assurer que l’utilisateur ne tente pas de manière répétée d’obtenir un jeton sans y parvenir. 
    * Vous allez créer la méthode `getDataWithToken` à l’étape suivante, mais rappelez-vous qu’elle définit une option appelée `forceConsent` sur `false`. Vous en saurez plus à la prochaine étape.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :

    * [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) est la nouvelle API d’Office.js qui permet à un complément de demander à l’application hôte Office (Excel, PowerPoint, Word, etc.) un jeton d’accès au complément (pour l’utilisateur connecté à Office). L’application hôte Office demande alors le jeton au point de terminaison Azure AD 2.0. Dans la mesure où vous avez préalablement autorisé l’hôte Office sur votre complément lors de son inscription, Azure AD enverra le jeton.
    * Si aucun utilisateur n’est connecté à Office, l’hôte Office invite l’utilisateur à se connecter.
    * Le paramètre d’options définit `forceConsent` sur `false`, donc l’utilisateur ne sera pas invité à accorder à l’hôte Office l’accès à votre complément chaque fois qu’il utilisera le complément. La première fois que l’utilisateur exécutera le complément, l’appel à `getAccessTokenAsync` échouera, mais la logique de gestion des erreurs que vous ajouterez dans une étape ultérieure effectuera automatiquement un autre appel avec le jeu d’options `forceConsent` défini sur `true`, et l’utilisateur sera invité à donner son consentement, mais uniquement la première fois.
    * Vous créerez la méthode `handleClientSideErrors` à une étape ultérieure.

    ```javascript
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. Remplacez TODO1 par les lignes suivantes. Vous créez la méthode `getData` et la route « /api/values » côté serveur dans les étapes suivantes. Une URL relative est utilisée pour le point de terminaison car il doit être hébergé sur le même domaine que votre complément.

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :

    * Cette méthode appelle un point de terminaison d’API Web spécifié et lui transmet le même jeton d’accès que l’application hôte Office a utilisé pour accéder à votre complément. Côté serveur, ce jeton d’accès est utilisé dans le flux « de la part de » pour obtenir un jeton d’accès à Microsoft Graph.
    * Vous créerez la méthode `handleServerSideErrors` à une étape ultérieure.

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        }); 
    }
    ```

### <a name="create-the-error-handling-methods"></a>Création des méthodes de gestion des erreurs

1. En dessous de la méthode `getData`, ajoutez la méthode suivante. Cette méthode gérera les erreurs dans le client du complément lorsque l’hôte Office ne parviendra pas à obtenir un jeton d’accès pour le service web du complément. Ces erreurs sont signalées avec un code d’erreur, donc la méthode utilise une instruction `switch` pour les distinguer.

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
            //        nor Microsoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user triggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. Remplacez `TODO2` par le code suivant. L’erreur 13001 se produit si l’utilisateur n’est pas connecté, ou s’il a annulé, sans y répondre, une invite lui demandant d’indiquer un deuxième facteur d’authentification. Dans les deux cas, le code réexécute la méthode `getDataWithToken` et définit une option pour forcer une invite de connexion.

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. Remplacez `TODO3` par le code suivant. L’erreur 13002 se produit lorsque la connexion ou l’octroi du consentement de l’utilisateur a été abandonné. Demandez à l’utilisateur de réessayer, mais seulement une fois.

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. Remplacez `TODO4` par le code suivant. L’erreur 13003 se produit si l’utilisateur est connecté avec un compte qui n’est ni un compte professionnel ni un compte scolaire, ni un compte Microsoft. Demandez à l’utilisateur de se déconnecter, puis de se reconnecter avec un type de compte pris en charge.

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > Les erreurs 13004 et 13005 ne sont pas gérées dans cette méthode, car elles ne devraient se produire qu’en développement. Elles ne peuvent pas être résolues par du code d’exécution et il ne serait d’aucune utilité de les signaler à un utilisateur final.

1. Remplacez `TODO5` par le code suivant. L’erreur 13006 se produit lorsqu’une erreur non spécifiée indiquant que l’hôte est dans un état instable est survenue dans l’hôte Office. Demandez à l’utilisateur de redémarrer Office.

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. Remplacez `TODO6` par le code suivant. L’erreur 13007 se produit lorsqu’un problème est survenu au niveau de l’interaction de l’hôte Office avec AAD de telle sorte que l’hôte ne peut pas obtenir de jeton d’accès pour accéder à l’application/au service Web des compléments. Il peut s’agir d’un problème temporaire de réseau. Demandez à l’utilisateur de réessayer plus tard.

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. Remplacez `TODO7` par le code suivant. L’erreur 13008 se produit lorsque l’utilisateur a déclenché une opération qui appelle `getAccessTokenAsync` avant que la fin de l’appel précédent.

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. Remplacez `TODO8` par le code suivant. L’erreur 13009 se produit lorsque le complément ne prend pas en charge l’obligation d’afficher une invite de consentement, mais que `getAccessTokenAsync` a été appelé avec l’option `forceConsent` définie sur `true`. Dans le cas habituel, lorsque cela se produit, le code doit automatiquement réexécuter `getAccessTokenAsync` avec l’option de consentement définie sur `false`. Toutefois, dans certains cas, l’appel de la méthode avec `forceConsent` défini sur `true` était lui-même une réponse automatique à une erreur dans un appel à la méthode avec l’option définie sur `false`. Dans ce cas, le code ne doit pas réessayer, mais il doit à la place conseiller à l’utilisateur de se déconnecter et de se reconnecter.

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. Remplacez `TODO9` par le code suivant.

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. En dessous de la méthode `handleClientSideErrors`, ajoutez la méthode suivante. Cette méthode gérera les erreurs du service web du complément en cas de problème d’exécution du flux « de la part de » ou de problème d’obtention de données à partir de Microsoft Graph.

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to its
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. Remplacez `TODO10` par le code suivant. Tenez compte des informations suivantes :

    * Il existe des configurations d’Azure Active Directory où l’on demande à l’utilisateur de fournir un ou plusieurs facteurs d’authentification supplémentaires pour accéder à certaines cibles Microsoft Graph (par exemple, OneDrive), même si l’utilisateur peut se connecter à Office par un simple mot de passe. Dans ce cas, AAD enverra, avec l’erreur 50076, une réponse comportant la propriété `Claims`. 
    * L’hôte Office dois obtenir un nouveau jeton avec la valeur **Claims** pour l’option `authChallenge`. Cela demande à AAD d’inviter l’utilisateur à accepter tous les formulaires d’authentification requis. 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. Remplacez `TODO11` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :

    * L’erreur 65001 signifie que l’utilisateur a refusé de donner l’accès à Microsoft Graph (ou que l’accès a été révoqué) pour une ou plusieurs autorisations. 
    * Le complément doit obtenir un nouveau jeton avec l’option `forceConsent` définie sur `true`.

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. Remplacez `TODO12` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :

    * L’erreur 70011 signifie qu’une portée (autorisation) non valide a été demandée. Le complément doit signaler l’erreur.
    * Le code consigne les autres erreurs avec un numéro d’erreur AAD.

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. Remplacez `TODO13` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :

    * Le code côté serveur que vous créerez à une étape ultérieure enverra le message qui se termine par `... expected access_as_user` si l’étendue (autorisation) `access_as_user` ne se trouve pas dans le jeton d’accès que le client du complément envoie à AAD, afin qu’il soit utilisé dans le flux « de la part de ».
    * Le complément doit signaler l’erreur.

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. Remplacez `TODO14` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :

    * Il est peu probable qu’un jeton expiré ou non valide soit envoyé à Microsoft Graph. Cependant, si cela se produit, le code côté serveur que vous créerez dans une étape ultérieure se terminera par la chaîne `Microsoft Graph error`.
    * Dans ce cas, le complément doit recommencer l’intégralité du processus d’authentification en réinitialisant les variables de compteur `timesGetOneDriveFilesHasRun` et d’indicateur `timesGetOneDriveFilesHasRun`, puis en appelant à nouveau la méthode de gestionnaire de boutons. Toutefois, il ne doit faire cela qu’une seule fois. Si l’erreur se produit à nouveau, il doit simplement la consigner.
    * Le code consigne l’erreur si elle se produit deux fois de suite.

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }        
    }
    ```

1. Remplacez `TODO15` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*.

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a>Code côté serveur

Il existe deux fichiers côté serveur qui doivent être modifiés. 
- Le fichier src\auth.js fournit des fonctions d’assistance pour l’autorisation. Il dispose déjà des membres génériques qui sont utilisés dans une variété de flux d’autorisation. Nous devons ajouter des fonctions qui implémentent le flux « de la part de ».
- Le fichier src\server.js possède les membres de base requis pour exécuter un serveur et les intergiciels express. Nous devons y ajouter des fonctions qui servent la page d’accueil et une API Web pour obtenir des données Microsoft Graph.

### <a name="create-a-method-to-exchange-tokens"></a>Créer une méthode pour échanger des jetons

1. Ouvrez le fichier \src\auth.ts. Ajoutez la méthode ci-après à la classe `AuthModule`. Tenez compte des informations suivantes :

    * Le paramètre `jwt` est le jeton d’accès à l’application. Dans le flux « de la part de », il est échangé avec AAD contre un jeton d’accès à la ressource.
    * Le paramètre scopes a une valeur par défaut, mais dans cet exemple, elle sera remplacée par le code appelant.
    * Le paramètre de ressource est facultatif. Il ne doit pas être utilisé lorsque le [service STS (Secure Token Service)](https://docs.microsoft.com/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) est le point de terminaison AAD V2.0. Le point de terminaison V2.0 déduit la ressource des étendues et renvoie une erreur si une ressource est envoyée dans la requête HTTP. 
    * La génération d’une exception dans le bloc `catch` ne provoquera *pas* l’envoi immédiat du message « 500 Erreur interne du serveur » au client. L’appel de code dans le fichier server.js interceptera cette exception et la convertira en un message d’erreur qui sera envoyé au client.

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the 
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                            + JSON.stringify(exception), 
                                            exception);
            }
        }
        ```

2. Remplacez `TODO3` par le code suivant. Tenez compte des informations suivantes :
    * Un STS qui prend en charge le flux « de la part de » attend certaines paires de propriété/valeur dans le corps de la requête HTTP. Ce code construit un objet qui devient le corps de la requête. 
    * Une propriété de ressource est ajoutée au corps si, et uniquement si, une ressource a été transmise à la méthode.

        ```typescript
        const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };
            let finalParams = {};
            if (resource) {
                // In JavaScript we could just add the resource property to the v2Params
                // object, but that won't compile in TypeScript.
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            } 
        ```

3. Remplacez `TODO4` par le code suivant, qui envoie la requête HTTP au point de terminaison de jeton du STS.

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. Remplacez `TODO5` par le code suivant. Vous remarquerez que la génération d’une exception ne provoquera *pas* l’envoi immédiat d’un message « 500 Erreur interne du serveur » au client. L’appel de code dans le fichier server.js interceptera cette exception et la convertira en un message d’erreur qui sera envoyé au client.

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. Remplacez `TODO6` par le code suivant. Vous remarquerez que le code prolonge le jeton d’accès à la ressource et son délai d’expiration, en plus de le renvoyer. Le code d’appel permet d’éviter les appels inutiles au STS en réutilisant un jeton d’accès non expiré à la ressource. Vous verrez comment procéder dans la section suivante.

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. Enregistrez le fichier, mais ne le fermez pas.

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>Créer une méthode pour accéder à la ressource à l’aide du flux « de la part de »

1. Toujours dans src/auth.ts, ajoutez la méthode ci-après à la classe `AuthModule`. Tenez compte des informations suivantes :

    * Les commentaires ci-dessus concernant les paramètres de la méthode `exchangeForToken` s’appliquent aussi aux paramètres de cette méthode.
    * La méthode recherche d’abord dans le stockage permanent un jeton d’accès à la ressource qui n’a pas expiré et qui ne va pas expirer dans la minute qui suit. Il appelle la méthode `exchangeForToken` que vous avez créée dans la dernière section uniquement si nécessaire.

    ```typescript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(await resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. Enregistrez et fermez le fichier.

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>Créer les points de terminaison que serviront la page d’accueil et les données du complément

1. Ouvrez le fichier src\server.ts. 

2. Ajoutez la méthode suivante au bas du fichier. Cette méthode servira la page d’accueil du complément. Le manifeste du complément spécifie l’URL de la page d’accueil.

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. Ajoutez la méthode suivante en bas du fichier. Cette méthode traite toutes les requêtes concernant l’API `values`.
    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. Remplacez `TODO7` par le code suivant, qui valide le jeton d’accès reçu à partir de l’application hôte Office. La méthode `verifyJWT` est définie dans le fichier src\auth.ts. Elle valide toujours l’audience et l’émetteur. Nous utilisons le paramètre facultatif pour spécifier que nous souhaitons également vérifier que l’étendue du jeton d’accès est `access_as_user`. C’est la seule autorisation d’accès au complément dont l’utilisateur et l’hôte Office ont besoin pour obtenir un jeton d’accès à Microsoft Graph au moyen du flux « de la part de ». 

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > Vous ne pouvez utiliser l’étendue `access_as_user` que pour autoriser l’API qui gère le flux « de la part de » pour les compléments Office. D’autres API dans votre service peuvent avoir leurs propres exigences d’étendue. Cela permet de limiter ce à quoi donnent accès les jetons acquis par Office.

5. Remplacez `TODO8` par le code suivant. Tenez compte des informations suivantes :

    * L’appel vers `acquireTokenOnBehalfOf` ne comprend pas de paramètre de ressource, étant donné que nous avons construit l’objet `AuthModule` (`auth`) avec le point de terminaison AAD V2.0 qui ne prend pas en charge une propriété de ressource.
    * Le deuxième paramètre de l’appel spécifie les autorisations dont le complément aura besoin pour obtenir une liste des fichiers et dossiers de l’utilisateur dans OneDrive. (L’autorisation `profile` n’est pas demandée, car elle n’est nécessaire qu’au moment où l’hôte Office obtient le jeton d’accès à votre complément, pas lorsque vous travaillez dans ce jeton pour un jeton d’accès à Microsoft Graph.)

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. Remplacez `TODO9` par la ligne suivante. Tenez compte des informations suivantes :

    * La classe MSGraphHelper est définie dans src\msgraph-helper.ts. 
    * Nous réduisons les données qui doivent être renvoyées en spécifiant que nous ne souhaitons que la propriété name et uniquement les 3 premiers éléments.

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

7. Remplacez `TODO10` par le code suivant. Notez que ce code gère les erreurs « 401 Non autorisé » de Microsoft Graph qui signalent un jeton expiré ou non valide. Il est très peu probable que cela se produise, car la logique persistante du jeton doit empêcher ces erreurs. (Reportez-vous à la section **Créer une méthode pour accéder à la ressource à l’aide du flux « de la part de »** ci-dessus.) Si cela se produit, ce code communiquera l’erreur au client avec, dans le nom de l’erreur, « Microsoft Graph error ». (Reportez-vous à la méthode `handleClientSideErrors` que vous avez créée dans le fichier program.js dans une étape précédente.) Le code que vous ajouterez au fichier ODataHelper.js à une étape ultérieure vous permet de traiter les erreurs provenant de Microsoft Graph.

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. Remplacez `TODO11` par le code suivant. Notez que Microsoft Graph renvoie des métadonnées OData et une propriété **eTag** pour chaque élément, même si `name` est la seule propriété demandée. Le code envoie uniquement les noms d’éléments au client.

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. Enregistrez et fermez le fichier.

### <a name="add-response-handling-to-the-odatahelper"></a>Ajouter une gestion des réponses à ODataHelper

1. Ouvrez le fichier src\odata-helper.ts. Le fichier est presque complet. Il manquant le corps du rappel au gestionnaire pour l’événement de « fin » de demande. Remplacez `TODO` par le code suivant. Tenez compte des informations suivantes sur ce code :

    * La réponse du point de terminaison OData peut-être une erreur, supposons une erreur 401 si le point de terminaison nécessite un jeton d’accès et que celui-ci n’est pas valide ou a expiré. Cependant, un message d’erreur reste un *message*, pas une erreur dans l’appel de `https.get`, donc la ligne `on('error', reject)` à la fin de `https.get` n’est pas déclenchée. Par conséquent, le code distingue les messages de réussite (200) des messages d’erreur, et envoie un objet JSON à l’appelant soit les informations d’erreur, soit avec les informations demandées.

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  Remplacez `TODO1` par le code suivant. Notez que le code suppose que les données sont renvoyées au format JSON.

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  Remplacez `TODO2` par le code suivant. Tenez compte des informations suivantes :

    * Une réponse d’erreur d’une source OData aura toujours un code d’état (statusCode) et généralement un message d’état (statusMessage). Certaines sources OData ajoutent également une propriété d’erreur au corps avec des informations supplémentaires, telles qu’un message et un code internes, ou plus spécifiques.
    * L’objet de promesse est résolu, pas rejeté. `https.get` s’exécute quand un service web appelle un point de terminaison OData de serveur à serveur. Cependant, cet appel s’inscrit dans le contexte d’un appel d’un client à une API web dans le service web. La demande « externe » du client au service web n’aboutit jamais si cette demande « interne » est rejetée. De plus, la résolution de la requête avec l’objet `Error` personnalisé est obligatoire si l’émetteur de l’appel `http.get` doit communiquer les erreurs du point de terminaison OData au client.

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;
    
    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. Enregistrez et fermez le fichier.

## <a name="deploy-the-add-in"></a>Déploiement du complément

Vous devez maintenant indiquer à Office où trouver le complément.

1. Créez un partage réseau, ou [partagez un dossier sur le réseau](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).

2. Placez une copie du fichier manifeste Office-Add-in-NodeJS-SSO.xml, depuis la racine du projet, dans le dossier partagé.

3. Lancez PowerPoint et ouvrez un document.

4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

5. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

6. Choisissez **Catalogues de compléments approuvés**.

7. Dans le champ **URL du catalogue**, saisissez le chemin réseau permettant d’accéder au partage de dossier qui contient le fichier Office-Add-in-NodeJS-SSO.xml, puis sélectionnez **Ajouter un catalogue**.

8. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.

9. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage de Microsoft Office. Fermez PowerPoint.

## <a name="build-and-run-the-project"></a>Création et exécution du projet

Il existe deux manières de créer et d’exécuter le projet selon que vous utilisez Visual Studio Code. Pour les deux façons, le projet est généré et reconstruit automatiquement, puis ré-exécuté lorsque vous apportez des modifications au code.

1. Si vous n’utilisez pas Visual Studio Code : 
 1. Ouvrez un terminal de nœud et accédez au dossier racine du projet.
 2. Dans le terminal, entrez **npm run build**. 
 3. Ouvrez un second terminal de nœud et accédez au dossier racine du projet.
 4. Dans le terminal, entrez **npm run start**.

2. Si vous utilisez VS Code :
 1. Ouvrez le projet dans VS Code.
 2. Appuyez sur CTRL-MAJ-B pour générer le projet.
 3. Appuyez sur **F5** pour exécuter le projet dans une session de débogage.


## <a name="add-the-add-in-to-an-office-document"></a>Ajouter le complément à un document Office

1. Redémarrez PowerPoint et ouvrez ou créez une présentation.

1. Si l’onglet **Développeur** n’est pas visible sur le ruban, activez-le en procédant comme suit :
 1. Accédez à **Fichier** | **Options** | **Personnaliser le ruban**.
 2. Cliquez sur la case à cocher pour activer **Développeur** dans l’arborescence des noms de contrôle dans la partie droite de la page **Personnaliser le ruban**.
 3. Appuyez sur **OK**.

2. Sous l’onglet **Développeur** de PowerPoint, choisissez **Mes compléments**.

3. Sélectionnez l’onglet **DOSSIER PARTAGÉ**.

4. Choisissez **Échantillon SSO NodeJS**, puis sélectionnez **OK**.

5. Dans le ruban **Accueil**, un nouveau groupe appelé **SSO NodeJS** apparaît avec un bouton intitulé **Afficher le complément** et une icône. 

## <a name="test-the-add-in"></a>Test du complément

1. Assurez-vous que vous disposez de fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.

2. Cliquez sur le bouton **Afficher le complément** pour ouvrir le complément.

2. Le complément s’ouvre avec une page d’accueil. Cliquez sur le bouton **Obtenir mes fichiers à partir de OneDrive**.

2. Si vous êtes connecté à Office, une liste de vos fichiers et dossiers sur OneDrive apparaîtront en dessous du bouton. La première fois, l’opération peut prendre plus de 15 secondes.

3. Si vous n’êtes pas connecté à Office, une fenêtre contextuelle s’ouvre et vous invite à vous connecter. Une fois que vous êtes connecté, la liste de vos fichiers et dossiers s’affiche après quelques secondes. *N’appuyez pas sur le bouton une deuxième fois.*

> [!NOTE]
> Si vous étiez précédemment connecté à Office avec un ID différent, et si certaines applications Office sont toujours ouvertes, Office ne changera pas systématiquement votre identifiant même s’il semble l’avoir fait dans PowerPoint. Dans ce cas, l’appel vers Microsoft Graph peut échouer, ou des données de l’ID précédent peuvent être renvoyées. Afin d’éviter ce problème, veillez à *fermer toutes les autres applications Office* avant de cliquer sur **Obtenir mes fichiers à partir de OneDrive**.
