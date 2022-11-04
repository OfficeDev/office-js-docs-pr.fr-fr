---
title: Création d’un complément Office Node.js qui utilise l’authentification unique
description: Découvrez comment créer un complément basé sur Node.js qui utilise l’authentification unique Office.
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 35128da43b3f27a58df5e188a5001bfa8aba4a4c
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68841728"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Création d’un complément Office Node.js qui utilise l’authentification unique

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément. L’exemple de complément que vous créez comporte deux parties : un volet Office qui se charge dans Microsoft Excel et un serveur de niveau intermédiaire qui gère les appels à Microsoft Graph pour le volet Office. Le serveur de niveau intermédiaire est créé avec Node.js et Express et expose une SEULE API REST, `/getuserfilenames`, qui retourne une liste des 10 premiers noms de fichiers dans le dossier OneDrive de l’utilisateur. Le volet Office utilise la `getAccessToken()` méthode pour obtenir un jeton d’accès pour l’utilisateur connecté au serveur de niveau intermédiaire. Le serveur de niveau intermédiaire utilise le flux On-Behalf-Of (OBO) pour échanger le jeton d’accès contre un nouveau jeton avec accès à Microsoft Graph. Vous pouvez étendre ce modèle pour accéder à toutes les données Microsoft Graph. Le volet Office appelle toujours une API REST de niveau intermédiaire (en passant le jeton d’accès) lorsqu’il a besoin de services Microsoft Graph. Le niveau intermédiaire utilise le jeton obtenu via OBO pour appeler les services Microsoft Graph et retourner les résultats au volet Office.

Cet article fonctionne avec un complément qui utilise Node.js et Express. Pour voir un article similaire sur un complément basé sur ASP.NET, reportez-vous à [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Conditions préalables

- [Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))

- [Git Bash](https://git-scm.com/downloads) (ou un autre client Git)

- Un éditeur de code - Nous vous recommandons Visual Studio Code

- Au moins quelques fichiers et dossiers stockés sur OneDrive Entreprise dans votre abonnement Microsoft 365

- Build de Microsoft 365 qui prend en charge [l’ensemble de conditions requises IdentityAPI 1.3](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Vous pouvez obtenir un [bac à sable de développeur gratuit](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) qui fournit un abonnement renouvelable de 90 jours Microsoft 365 E5 développeur. Le bac à sable développeur inclut un abonnement Microsoft Azure que vous pouvez utiliser pour les inscriptions d’applications dans les étapes ultérieures de cet article. Si vous préférez, vous pouvez utiliser un abonnement Microsoft Azure distinct pour les inscriptions d’applications. Obtenez un abonnement d’essai sur [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

1. Clonez ou téléchargez le référentiel sur [Complément Office NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

   > [!NOTE]
   > Il existe deux versions de l’échantillon :
   >
   > - Le dossier **Begin** est un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. Les sections suivantes de cet article vous guident tout au long de la procédure d’exécution de cette dernière.
   > - Le dossier **Complete** contient le même exemple avec toutes les étapes de codage de cet article terminées. Pour utiliser la version terminée, suivez simplement les instructions de cet article, mais remplacez « Begin » par « Complete » et ignorez les sections **Coder côté client** et **Code côté serveur de niveau intermédiaire** .

1. Ouvrez une invite de commandes dans le dossier **Begin** .

1. Saisissez `npm install`dans la console pour installer toutes les dépendances détaillées dans le fichier package.json.

1. Exécutez la commande `npm run install-dev-certs`. Sélectionnez **Oui** lorsque vous êtes invité à installer le certificat.

Utilisez les valeurs suivantes pour les espaces réservés pour les étapes d’inscription d’application suivantes.

| Espace réservé           | Valeur                                 |
|-----------------------|---------------------------------------|
| `<add-in-name>`       | **Office-Add-in-NodeJS-SSO**          |
| `<redirect-platform>` | **Application monopage (SPA)**     |
| `<redirect-uri>`      | `https://localhost:44355/dialog.html` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-add-in"></a>Configurer le complément

1. Ouvrez le dossier `\Begin` dans le projet cloné dans votre éditeur de code.

1. Ouvrez le `.ENV` fichier et utilisez les valeurs que vous avez copiées précédemment à partir de l’inscription de l’application **Office-Add-in-NodeJS-SSO** . Définissez les valeurs comme suit :

   | Nom              | Valeur                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | **ID d’application (client)** à partir de la page de vue d’ensemble de l’inscription d’application. |
   | **CLIENT_SECRET** | **Clé secrète client** enregistrée à partir de **la page Certificats & Secrets** .       |
   | **DIRECTORY_ID**  | **ID d’annuaire (locataire)** à partir de la page de vue d’ensemble de l’inscription de l’application.   |

   Les valeurs ne doivent **pas** se trouver entre des guillemets. Quand vous avez terminé, votre modèle doit ressembler à ce qui suit :

   ```javascript
   CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
   CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
   DIRECTORY_ID=478aa78e-20ba-4c0d-9ffe-c4f62e5de3d5
   NODE_ENV=development
   SERVER_SOURCE=https://localhost:44355

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:44355/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. Remplacez l’espace réservé « $app-id-guid$ » _aux deux emplacements_ du balisage par **l’ID d’application** que vous avez copié lors de la création de l’inscription de l’application **Office-Add-in-NodeJS-SSO** . Les symboles « $ » ne font pas partie de l’ID. Ne les incluez donc pas. Il s’agit du même ID que celui utilisé pour le CLIENT_ID dans . Fichier ENV.

   > [!NOTE]
   > La **\<Resource\>** valeur est **l’URI d’ID d’application** que vous définissez lors de l’inscription du complément. La **\<Scopes\>** section est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.

1. Ouvrez le fichier `\public\javascripts\fallback-msal\authConfig.js`. Remplacez l’espace réservé « $app-id-guid$ » par l’ID d’application que vous avez enregistré à partir de l’inscription de l’application **Office-Add-in-NodeJS-SSO** que vous avez créée précédemment.

1. Enregistrez les modifications du fichier.

## <a name="code-the-client-side"></a>Code du côté client

### <a name="create-client-request-and-response-handler"></a>Créer un gestionnaire de demandes et de réponses client

1. Ouvrez le fichier `public\javascripts\ssoAuthES6.js` dans votre éditeur de code. Il possède déjà du code qui garantit que les promesses sont prises en charge, même dans Internet Explorer 11, et un appel `Office.onReady` pour attribuer un gestionnaire au bouton unique du complément.

   > [!NOTE]
   > Comme leur nom l’indique, ssoAuthES6.js utilise la syntaxe JavaScript ES6, car l’utilisation de `async` et de `await` illustre le mieux la simplicité de l’API SSO. Lorsque le serveur localhost est démarré, ce fichier est transpilé en syntaxe ES5 afin que l’exemple prend en charge Internet Explorer 11.

    Une partie clé de l’exemple de code est la demande du client. La requête cliente est un objet qui effectue le suivi des informations sur la demande d’appel d’API REST sur le serveur de niveau intermédiaire. Cela est nécessaire, car l’état de la demande du client doit être suivi ou mis à jour dans le scénario suivant :

    - L’authentification unique échoue et l’authentification de secours est requise. Le jeton d’accès est acquis via MSAL dans une boîte de dialogue contextuelle. L’objectif est de ne pas échouer dans ce scénario et de revenir normalement à l’approche d’authentification alternative.

    L’objet de requête client effectue le suivi des données suivantes :

    - `authSSO` - true si vous utilisez l’authentification unique, sinon false.
    - `verb` - Verbe d’API REST tel que GET et POST.
    - `accessToken`- Jeton d’accès au serveur ASP.NET Core.
    - `url`- URL de l’API REST à appeler sur le serveur ASP.NET Core.
    - `callbackRESTApiHandler` - Fonction permettant de transmettre les résultats de l’appel d’API REST.
    - `callbackFunction` : fonction à laquelle transmettre la requête du client quand elle est prête.

1. Pour initialiser l’objet de demande client, dans la `createRequest` fonction , remplacez par `TODO 1` le code suivant.

    ```javascript
    const clientRequest = {
      authSSO: authSSO,
      verb: verb,
      accessToken: null,
      url: url,
      callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. Remplacez `TODO 2` par le code suivant. Tenez compte du code suivant :

    - Il vérifie si l’authentification unique est utilisée. La méthode permettant d’acquérir le jeton d’accès est différente pour l’authentification unique et pour l’authentification de secours.
    - Si l’authentification unique retourne le jeton d’accès, elle appelle la `callbackfunction` fonction . Pour l’authentification de secours, il appelle `dialogFallback`, qui appellera finalement la fonction de rappel une fois que l’utilisateur se connecte via MSAL.

    ```javascript
    // Get access token.

    if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      callbackFunction(clientRequest);
    } catch (error) {
      // handle the SSO error which will inform us if we need to switch to fallback auth.
      let fallbackRequired = handleSSOErrors(error);
      if (fallbackRequired) switchToFallbackAuth(clientRequest);
    }
   } else {
     // Use fallback auth to get access token.
     dialogFallback(clientRequest);
   }
    ```

1. Dans la fonction `getFileNameList`, remplacez `TODO 3` par le code suivant. Tenez compte du code suivant :

    - La fonction `getFileNameList` est appelée lorsque l’utilisateur choisit le bouton **Obtenir les noms de fichiers OneDrive** dans le volet Office.
    - Il crée une requête cliente pour suivre des informations sur l’appel, telles que l’URL de l’API REST.
    - Lorsque l’API REST retourne un résultat, il est passé à la `handleGetFileNameResponse` fonction . Ce rappel est passé en tant que paramètre à `createRequest` et fait l’objet d’un suivi dans `clientRequest.callbackRESTApiHandler`.
    - Le code appelle `callWebServer` avec la demande du client pour effectuer les étapes suivantes et appeler l’API REST.

    ```javascript
    createRequest(
      "GET",
      "/getuserfilenames",
      handleGetFileNameResponse,
      async (clientRequest) => {
        await callWebServer(clientRequest);
      }
    );
    ```

1. Dans la fonction `handleGetFileNameResponse`, remplacez `TODO 4` par le code suivant. Tenez compte du code suivant :

    - Le code transmet la réponse (qui contient une liste de noms de fichiers) à pour `writeFileNamesToOfficeDocument` écrire les noms de fichiers dans le document.
    - Le code recherche les erreurs. Il affiche un message de réussite si les noms de fichiers sont écrits; sinon, il affiche une erreur.

    ```javascript
    if (response !== null) {
      try {
        await writeFileNamesToOfficeDocument(response);
        showMessage("Your OneDrive filenames are added to the document.");
      } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin
        // "Unable to add filenames to document."
        showMessage(error);
      }
    } else
    showMessage("A null response was returned to handleGetFileNameResponse.");
    ```

1. Dans la fonction `handleSSOErrors`, remplacez `TODO 5` par le code suivant. Pour plus d’informations sur ces erreurs, reportez-vous à [Résoudre les problèmes liés à SSO dans les compléments Office](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    let fallbackRequired = false;

   switch (err.code) {
     case 13001:
       // No one is signed into Office. If the add-in cannot be effectively used when no one
       // is logged into Office, then the first call of getAccessToken should pass the
       // `allowSignInPrompt: true` option. Since this sample does that, you should not see
       // this error.
       showMessage(
         "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
       );
       break;
     case 13002:
       // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
       // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
       showMessage(
         "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
       );
       break;
     case 13006:
       // Only seen in Office on the web.
       showMessage(
         "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
       );
       break;
     case 13008:
       // Only seen in Office on the web.
       showMessage(
        "Office is still working on the last operation. When it completes, try this operation again."
       );
       break;
     case 13010:
       // Only seen in Office on the web.
       showMessage(
         "Follow the instructions to change your browser's zone configuration."
       );
       break;
    ```

1. Remplacez `TODO 6` par le code suivant. Pour plus d’informations sur ces erreurs, voir [Résoudre les problèmes d’authentification unique dans les compléments Office](troubleshoot-sso-in-office-add-ins.md). Pour toutes les erreurs qui ne peuvent pas être gérées, `true` est retourné à l’appelant. Cela indique que l’appelant doit passer à l’utilisation de MSAL comme authentification de secours.

    ```javascript
     default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      fallbackRequired = true;
      break;
    }
    return fallbackRequired;
    ```

### <a name="call-the-rest-api-on-the-middle-tier-server"></a>Appeler l’API REST sur le serveur de niveau intermédiaire

1. Dans la fonction `callWebServer`, remplacez `TODO 7` par le code suivant. Tenez compte du code suivant :

    - L’appel AJAX réel sera effectué par la `ajaxCallToRESTApi` fonction .
    - Cette fonction tente d’obtenir un nouveau jeton d’accès si le serveur de niveau intermédiaire retourne une erreur indiquant que le jeton actuel a expiré.
    - Si l’appel AJAX ne peut pas être effectué correctement, `switchToFallbackAuth` est appelé pour utiliser l’authentification MSAL au lieu de l’authentification unique Office.

    ```javascript
    try {
    const data = await $.ajax({
      type: clientRequest.verb,
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
    });
    clientRequest.callbackRESTApiHandler(data);

    } catch (error) {
     // TODO 8: Check for expired SSO token and refresh if needed.

    // TODO 9: Check for Microsoft Graph and other errors.

    }
    ```

1. Remplacez `TODO 8` par le code suivant. Tenez compte du code suivant :

    - Lorsque le serveur identifie un jeton expiré, il retourne une erreur de type « TokenExpiredError ».
    - L’essai... catch appelle Office.auth.getAccessToken pour obtenir un jeton actualisé avec une nouvelle expiration.
    - Le code tente à nouveau d’appeler l’API du serveur.

    ```javascript
    // Check for expired SSO token. Refresh and retry the call if it expired.
    if (
      error.responseJSON &&
      authSSO === true &&
      error.responseJSON.type === "TokenExpiredError"
    ) {
      try {
        const accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
        const data = await $.ajax({
          type: clientRequest.verb,
          url: clientRequest.url,
          headers: { Authorization: "Bearer " + accessToken },
          cache: false,
        });
        clientRequest.callbackRESTApiHandler(data);
      } catch (error) {
        showMessage(error.responseText);
        switchToFallbackAuth(clientRequest);
        return;
      }
    }
    ```

1. Remplacez `TODO 9` par le code suivant. Tenez compte du code suivant :

    - Pour les erreurs **Microsoft Graph** , affichez le message dans le volet Office.
    - Pour tous les autres messages, affichez le message dans le volet Office.

    ```javascript
    // Check for a Microsoft Graph API call error. which is returned as bad request (403)
    if (error.status === 403) {
      if (error.responseJSON && error.responseJSON.type === "Microsoft Graph") {
        showMessage(error.responseJSON.errorDetails);
      } else {
        showMessage(error);
      }
      return;
    }

    // For all other error scenarios, display the message and use fallback auth.
    showMessage("Unknown error from web server: " + JSON.stringify(error));
    if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

L’authentification de secours utilise la bibliothèque MSAL pour connecter l’utilisateur. Le complément lui-même est un spa et utilise une inscription d’application SPA pour accéder au serveur de niveau intermédiaire.

1. Dans la fonction `switchToFallbackAuth`, remplacez `TODO 10` par le code suivant. Tenez compte du code suivant :

    - Il définit le global `authSSO` sur false et crée une nouvelle requête client qui utilise MSAL pour l’authentification. La nouvelle requête a un jeton d’accès MSAL au serveur de niveau intermédiaire.
    - Une fois la demande créée, elle appelle `callWebServer` pour continuer à tenter d’appeler le serveur de niveau intermédiaire.

    ```javascript
    // Guard against accidental call to this function when fallback is already in use.

    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
      clientRequest.verb,
      clientRequest.url,
      clientRequest.callbackRESTApiHandler,
      async (fallbackRequest) => {
        // Hand off to call using fallback auth.
        await callWebServer(fallbackRequest);
      }
    );
    ```

## <a name="code-the-middle-tier-server"></a>Coder le serveur de niveau intermédiaire

Le serveur de niveau intermédiaire fournit des API REST que le client doit appeler. Par exemple, l’API `/getuserfilenames` REST obtient une liste de noms de fichiers à partir du dossier OneDrive de l’utilisateur. Chaque appel d’API REST nécessite un jeton d’accès par le client pour s’assurer que le client approprié accède à ses données. Le jeton d’accès est échangé contre un jeton Microsoft Graph via le flux On-Behalf-Of (OBO). Le nouveau jeton Microsoft Graph est mis en cache par la bibliothèque MSAL pour les appels d’API suivants. Il n’est jamais envoyé en dehors du serveur de niveau intermédiaire. Pour plus d’informations, consultez [Demande de jeton d’accès de niveau intermédiaire](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>Créer l’itinéraire et implémenter le flux On-Behalf-Of

1. Ouvrez le fichier `routes\getFilesRoute.js` et remplacez par `TODO 11` le code suivant. Tenez compte du code suivant :

    - Il appelle `authHelper.validateJwt`. Cela garantit que le jeton d’accès est valide et qu’il n’a pas été falsifié.
    - Pour plus d’informations, consultez [Validation des jetons](/azure/active-directory/develop/access-tokens#validating-tokens).

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 12: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. Remplacez `TODO 12` par le code suivant. Tenez compte du code suivant :

    - Il demande uniquement les étendues minimales dont il a besoin, telles que `files.read`.
    - Il utilise msal `authHelper` pour effectuer le flux OBO dans l’appel à `acquireTokenOnBehalfOf`.

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(" ")[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === "access_as_user"
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 13: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 14: Handle any errors.
    }
    ```

1. Remplacez `TODO 13` par le code suivant. Tenez compte du code suivant :

    - Il construit l’URL de l’appel API Graph Microsoft, puis effectue l’appel via la `getGraphData` fonction .
    - Il retourne des erreurs en envoyant une réponse HTTP 500 avec des détails.
    - En cas de réussite, il retourne le json avec la liste des noms de fichiers au client.

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

    const graphData = await getGraphData(
      response.accessToken,
      rootUrl,
      params
    );

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      res
        .status(403)
        .send({
          type: "Microsoft Graph",
          errorDetails:
            "An error occurred while calling the Microsoft Graph API.\n" +
            graphData,
        });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    ```

1. Remplacez `TODO 14` par le code suivant. Ce code vérifie spécifiquement si le jeton a expiré, car le client peut demander un nouveau jeton et appeler à nouveau.

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
   } else {
     res.status(403).send({ type: "Unknown", errorDetails: err });
   }
   ```

L’exemple doit gérer à la fois l’authentification de secours via MSAL et l’authentification unique via Office. L’exemple essaie d’abord l’authentification unique, et la `authSSO` valeur booléenne en haut du fichier suit si l’exemple utilise l’authentification unique ou a basculé vers l’authentification de secours.

## <a name="run-the-project"></a>Exécutez le projet

1. Assurez-vous d’avoir des fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.

1. Ouvrez une invite de commandes dans la racine du dossier `\Begin`.

1. Exécutez la commande `npm install` pour installer toutes les dépendances de package.

1. Exécutez la commande `npm start` pour démarrer le serveur de niveau intermédiaire.

1. Vous devez charger une version du complément dans une application Office (Excel, Word ou PowerPoint) pour le tester. Les instructions sont fonction de votre plateforme. Vous trouverez des liens vers des instructions sur [Charger une version du complément Office pour le tester](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. Dans l’application Office, sur le ruban **Accueil**, sélectionnez le bouton **Afficher le complément** dans le groupe **Node.js SSO** pour ouvrir le complément du panneau des tâches.

1. Cliquez sur le bouton **Obtenir des noms de fichier OneDrive**. Si vous êtes connecté à Office avec un compte Microsoft 365 Éducation ou professionnel, ou un compte Microsoft, et que l’authentification unique fonctionne comme prévu, les 10 premiers noms de fichier et de dossier de votre OneDrive Entreprise sont insérés dans le document. (La première fois, il peut s’agir de 15 secondes.) Si vous n’êtes pas connecté, ou si vous êtes dans un scénario qui ne prend pas en charge l’authentification unique, ou si l’authentification unique ne fonctionne pas pour une raison quelconque, vous êtes invité à vous connecter. Une fois connecté, les noms des fichiers et des dossiers s’affichent.

> [!NOTE]
> Si vous étiez précédemment connecté à Office avec un ID différent et si certaines applications précédemment ouvertes Office le sont toujours, Office ne changera pas systématiquement votre identifiant même si cela semble être le cas. Dans ce cas, l’appel vers Microsoft Graph peut échouer ou des données de l’ID précédent peuvent être renvoyées. Afin d’éviter ce problème, veillez à _fermer toutes les autres applications Office_ avant de cliquer sur **Obtenir des noms de fichiers OneDrive**.

## <a name="security-notes"></a>Notes de sécurité

- L’itinéraire `/getuserfilenames` dans `getFilesroute.js` utilise une chaîne littérale pour composer l’appel pour Microsoft Graph. Si vous modifiez l’appel de sorte qu’une partie de la chaîne provient d’une entrée utilisateur, nettoyez l’entrée afin qu’elle ne puisse pas être utilisée dans une attaque par injection d’en-tête de réponse.

- Dans `app.js` le contenu suivant, une stratégie de sécurité est en place pour les scripts. Vous pouvez spécifier des restrictions supplémentaires en fonction des besoins de sécurité de votre complément.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Suivez toujours les meilleures pratiques en matière de sécurité dans la [documentation Plateforme d'identités Microsoft](/azure/active-directory/develop/).
