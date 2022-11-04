---
title: Créer un complément Office ASP.NET qui utilise l’authentification unique
description: Guide pas à pas pour créer (ou convertir) un complément Office avec un back-end ASP.NET pour utiliser l’authentification unique (SSO).
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b0179429f9d81b893394278580b6ef8891dd0a87
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68842102"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>Créer un complément Office ASP.NET qui utilise l’authentification unique

Lorsque les utilisateurs sont connectés à Office, votre complément peut utiliser les mêmes informations d’identification pour permettre aux utilisateurs d’accéder à plusieurs applications sans avoir à se connecter une deuxième fois. Pour obtenir une vue d’ensemble, consultez la rubrique [Activer l’authentification unique dans un complément Office](sso-in-office-add-ins.md).
Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément créé avec ASP.NET.

## <a name="prerequisites"></a>Conditions préalables

- Visual Studio 2019 ou version ultérieure.

- Charge de travail **développement Office/SharePoint** lors de la configuration de Visual Studio.

- [Outils de développement Office](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

- Au moins quelques fichiers et dossiers stockés sur OneDrive Entreprise dans votre abonnement Microsoft 365.

- Un compte Azure avec un abonnement actif : [créez un compte gratuitement](https://azure.microsoft.com/free/?WT.mc_id=A261C142F).

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

Clonez ou téléchargez le référentiel sur [Complément Office ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

> [!NOTE]
> Il existe deux versions de l’exemple.
>
> - The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.
> - La version **Complète** de l’échantillon s’apparente au complément obtenu si vous aviez terminé les procédures de cet article, sauf que le projet final comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, suivez simplement les instructions de cet article, mais remplacez « Avant » par « Finale » et ignorez les sections **Code côté client** et **Code côté serveur**.

Utilisez les valeurs suivantes pour les espaces réservés pour les étapes d’inscription d’application suivantes.

| Espace réservé           | Valeur                                           |
|-----------------------|-------------------------------------------------|
| `<add-in-name>`       | **Office-Add-in-ASPNET-SSO**                    |
| `<redirect-platform>` | **Web**                                         |
| `<redirect-uri>`      | `https://localhost:44355/AzureADAuth/Authorize` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-solution"></a>Configurer la solution

1. À la racine du dossier **Before**, ouvrez le fichier (.sln) solution dans **Visual Studio**. Cliquez avec le bouton droit sur le nœud supérieur de l’**Explorateur de solutions** (le nœud solution, et non l’un des nœuds de projet), puis sélectionnez **Définir les projets de démarrage**.

1. Sous **Propriétés communes**, sélectionnez **Projet de démarrage**, puis **Plusieurs projets de démarrage**. Assurez-vous que l’**Action** pour les deux projets est définie sur **Démarrer**, et que le projet qui se termine par « ...WebAPI » apparaît en premier dans la liste. Fermez la boîte de dialogue.

1. De retour dans **Explorateur de solutions**, sélectionnez (ne cliquez pas avec le bouton droit) le projet **Office-Add-in-ASPNET-SSO-WebAPI**. Le volet **Propriétés** s’ouvre. Assurez-vous que **SSL activé** est **Vrai**. Vérifiez que l’**URL SSL** est `http://localhost:44355/`.

1. Dans « web.config », utilisez les valeurs que vous avez copiées dans le version précédente. Configurez les **Ida:ClientID** et **Ida:Audience** à votre **ID d’application (client)**, puis configurez **Ida:Password** sur votre code secret client. Définissez également **ida:Domain** sur `http://localhost:44355` (aucune barre oblique « / » à la fin).

    > [!NOTE]
    > **L’ID d’application (client)** est la valeur « audience » lorsque d’autres applications, telles que l’application cliente Office (par exemple, PowerPoint, Word, Excel), recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Si vous n’avez pas choisi « Comptes dans ce répertoire d’organisation uniquement » pour **TYPES DE COMPTES PRIS EN CHARGE** lorsque vous avez enregistré le complément, enregistrez et fermez le fichier web.config. Dans le cas contraire, enregistrez-le et laissez-le ouvert.

1. Toujours dans **Explorateur de solutions**, choisissez le projet **Office-Add-in-ASPNET-SSO** et ouvrez le fichier manifeste de complément « Office-Add-in-ASPNET-SSO.xml », puis faites défiler jusqu’au bas du fichier. Juste au-dessus de la balise de fin `</VersionOverrides>` , vous trouverez le balisage suivant.

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

1. Remplacez l’espace réservé « $application_GUID here$ » *aux deux endroits* du balisage par l’ID d’application que vous avez copiée lorsque vous avez inscrit votre complément. Les signes « $ » ne faisant pas partie de l’ID, vous ne devez pas les inclure. C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.

    > [!NOTE]
    > La **\<Resource\>** valeur est **l’URI d’ID d’application** que vous définissez lors de l’inscription du complément. La **\<Scopes\>** section est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.

1. Enregistrez et fermez le fichier.

### <a name="setup-for-single-tenant"></a>Configuration d’un seul locataire

Si vous avez choisi « Comptes dans cet annuaire organisationnel uniquement » pour **TYPES DE COMPTES PRIS EN CHARGE** lorsque vous avez inscrit le complément, vous devez effectuer ces étapes de configuration supplémentaires.

1. Revenez au portail Azure et ouvrez le volet **vue d’ensemble** de l’inscription du complément. Copiez l’**ID de répertoire (client)**.

1. Dans le fichier Web. config, remplacez le « Common » par la valeur de **Ida:Authority** avec le GUID que vous avez copié à l’étape précédente. Lorsque vous avez terminé, la valeur doit ressembler à ceci : `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.

1. Enregistrez et fermez le fichier web.config.

## <a name="code-the-client-side"></a>Code côté client

1. Ouvrez le fichier HomeES6.js dans le dossier **Scripts**. Il contient déjà du code.

    - Un polyfill qui affecte l’objet Office. promesse à l’objet fenêtre globale pour que le complément puisse s’exécuter lorsque Office utilise Internet Explorer pour l’interface utilisateur. (Pour plus d’informations, voir [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).)
    - Affectation à la `Office.initialize` fonction qui, à son tour, affecte un gestionnaire à l’événement click de `getGraphAccessTokenButton` bouton.
    - Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.
    - Une méthode `logErrors` qui consigne dans la console les erreurs qui ne sont pas destinées à l’utilisateur final.
    - Code qui implémente le système d’autorisation de secours que le complément utilisera dans les scénarios où l’authentification unique n’est pas prise en charge ou où une erreur s’est produite.

1. Après l’affectation à `Office.initialize`, ajoutez le code suivant. Tenez compte du code suivant :

    - La gestion des erreurs dans le complément tente parfois automatiquement d’obtenir un jeton d’accès une deuxième fois, à l’aide d’un autre jeu d’options. La variable de compteur `retryGetAccessToken` permet de s’assurer que l’utilisateur ne tente pas de manière répétée d’obtenir un jeton sans y parvenir.
    - La fonction `getGraphData` est définie avec le mot clé ES6 `async`. L’utilisation de la syntaxe ES6 simplifie l’utilisation de l’API d’authentification unique dans les compléments Office. Il s’agit du seul fichier dans la solution qui utilise une syntaxe non prise en charge par Internet Explorer. Nous plaçons « ES6 » dans le nom du fichier comme rappel. La solution utilise le transpondeur tsc pour transpiler ce fichier en ES5, afin que le complément puisse être exécuté lorsque Office utilise Internet Explorer pour l’interface utilisateur. (Consultez le fichier tsconfig.json dans la racine du projet.)

    ```javascript
    let retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. Après la `getGraphData` fonction , ajoutez la fonction suivante. Notez que vous créez la fonction `handleClientSideErrors` dans une étape ultérieure.

    > [!NOTE]
    > Pour faire la distinction entre les deux jetons d’accès que vous utilisez dans cet article, le jeton retourné par getAccessToken() est appelé jeton d’amorçage. Il est ensuite échangé via le flux On-Behalf-Of contre un nouveau jeton ayant accès à Microsoft Graph.

    ```javascript
    async function getDataWithToken(options) {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for a new access token to Microsoft Graph and then get the data
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


1. Remplacez par `TODO 1` le code suivant pour obtenir le jeton d’accès à partir de l’hôte Office. Le paramètre *options* contient les paramètres suivants passés à partir de la fonction précédente `getGraphData()` .

    - `allowSignInPrompt` a la valeur true. Cela indique à Office d’inviter l’utilisateur à se connecter si l’utilisateur n’est pas déjà connecté à Office.
    - `allowConsentPrompt` a la valeur true. Cela indique à Office d’inviter l’utilisateur à donner son consentement pour permettre au complément d’accéder au profil Microsoft Azure Active Directory de l’utilisateur, si le consentement n’a pas déjà été accordé. (L’invite qui en résulte n’autorise *pas* l’utilisateur à donner son consentement à des étendues Microsoft Graph.)
    - `forMSGraphAccess` a la valeur true. Cela indique à Office de retourner une erreur (code 13012) si l’utilisateur ou l’administrateur n’a pas accordé son consentement aux étendues Graph pour le complément. Pour accéder à Microsoft Graph, le complément doit échanger le jeton d’accès contre un nouveau jeton d’accès via le flux on-behalf-of. La définition `forMSGraphAccess` de la valeur true permet d’éviter le scénario dans lequel **getAccessToken()** réussit, mais le flux on-behalf-of échoue ultérieurement pour Microsoft Graph. Le code côté client du complément peut répondre au 13012 en branchant un système d’autorisation de secours.

    Notez également pour le code suivant :

    - Vous créez la fonction `getData` dans une étape ultérieure.
    - Le `/api/values` paramètre est l’URL d’un contrôleur côté serveur qui utilisera le flux on-behalf-of pour échanger le jeton contre un nouveau jeton d’accès pour appeler Microsoft Graph.

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken(options);

    getData("/api/values", bootstrapToken);
    ```

1. Après la `getGraphData` fonction , ajoutez ce qui suit. Tenez compte du code suivant :

    - Il est utilisé par les systèmes d’authentification unique et de secours.
    - Le paramètre `relativeUrl` est un contrôleur côté serveur.
    - Le paramètre `accessToken` peut être un jeton d’amorçage ou un jeton d’accès complet.
    - Le `writeFileNamesToOfficeDocument` fait déjà partie du projet.
    - Vous créez la fonction `handleServerSideErrors` dans une étape ultérieure.

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

1. Après la `getData` fonction , ajoutez la fonction suivante. Veuillez noter que `error.code` est un nombre, généralement compris dans la plage 13xxx.

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
        showResult(["No one is signed into Office. But you can use many of the add-in's functions anyway. If you want to sign in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-in's functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
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

1. Remplacez `TODO 3` par le code suivant. Pour toutes les autres erreurs, le complément se branche au système d’autorisation de secours. Pour plus d’informations sur ces erreurs, voir [Résoudre les problèmes d’authentification unique dans les compléments Office](troubleshoot-sso-in-office-add-ins.md). Dans ce complément, le système de secours ouvre une boîte de dialogue qui oblige l’utilisateur à se connecter, même si l’utilisateur l’est déjà.

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>Gérer les erreurs côté serveur

1. Après la `handleClientSideErrors` fonction , ajoutez la fonction suivante.

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
    const message = JSON.parse(result.responseText).Message;
    const exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. Remplacez `TODO 5` par ce qui suit. Lorsque Microsoft Graph exige un formulaire d’authentification supplémentaire, il envoie l’erreur AADSTS50076. Celle-ci inclut des informations sur la configuration requise supplémentaire dans la propriété **message les déclarations**. Pour gérer ce problème, le code effectue une deuxième tentative d’obtention du jeton d’amorçage, mais cette fois, il inclut la demande d’un facteur supplémentaire comme valeur de l’option `authChallenge`, ce qui indique à Azure AD d’inviter l’utilisateur à fournir toutes les formes requises d’authentification.

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            const claims = JSON.parse(message).Claims;
            const claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. Remplacez par `TODO 6` ce qui suit :

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

1. Remplacez par `TODO 8` ce qui suit :

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

1. Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:

    `public partial class Startup`

1. Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. Remplacez le `TODO 1` par ce qui suit. Tenez compte du code suivant :

    - Le code indique à OWIN de s’assurer que l’audience spécifiée dans le jeton d’amorçage provenant de l’application Office doit correspondre à la valeur spécifiée dans le web.config.
    - Les comptes Microsoft ont un GUID d’émetteur différent de n’importe quel GUID de locataire organisationnel. Par conséquent, pour prendre en charge les deux types de comptes, nous ne validons pas l’émetteur.
    - Si vous définissez `SaveSigninToken` sur `true` , OWIN enregistre le jeton d’amorçage brut à partir de l’application Office. Le complément en a besoin pour obtenir un jeton d’accès à Microsoft Graph avec le flux « de la part de ».
    - Les étendues ne sont pas validées par l’intergiciel OWIN. Les étendues du jeton d’amorçage, qui doivent inclure `access_as_user`, sont validées dans le contrôleur.

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. Remplacez `TODO 2` par ce qui suit. Tenez compte du code suivant :

    - La méthode `UseOAuthBearerAuthentication` est appelée au lieu de la méthode `UseWindowsAzureActiveDirectoryBearerAuthentication` plus courante, car cette dernière n’est pas compatible avec le point de terminaison Azure AD V2.
    - L’URL transmise à la méthode est l’endroit où le middleware OWIN obtient des instructions pour obtenir la clé dont il a besoin pour vérifier la signature sur le jeton d’amorçage reçu de l’application Office. Le segment d’autorité de l’URL provient du fichier web.config. Il s’agit soit de la chaîne « commun », soit d’un GUID pour un complément à un seul locataire.

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

1. Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.

1. Ajoutez la méthode suivante à `ValuesController`. Vous remarquerez que la valeur renvoyée est `Task<HttpResponseMessage>` et non `Task<IEnumerable<string>>`, laquelle serait plus courante pour une méthode `GET api/values`. Il s’agit d’un effet secondaire de ce fait que la logique d’autorisation OAuth doit se trouver dans le contrôleur, plutôt que dans un filtre ASP.NET. Certaines conditions d’erreur dans cette logique nécessitent qu’un objet de réponse HTTP soit envoyé au client du complément.

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //         token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get a new access token for Microsoft Graph.

        // TODO 4: Use the new access token to call Microsoft Graph.
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

    - Votre complément ne joue plus le rôle d’une ressource (ou d’une audience) à laquelle l’application Office et l’utilisateur ont besoin d’accéder. Désormais, il est lui-même un client qui a besoin d’accéder à Microsoft Graph. `ConfidentialClientApplication` est l’objet de « contexte client » MSAL.
    - À partir de MSAL.NET 3. x. x, le `bootstrapContext` est simplement le jeton d’amorçage.
    - L’autorité provient du fichier web.config. Il s’agit soit de la chaîne « commun », soit d’un GUID pour un complément à un seul locataire.
    - MSAL génère une erreur si votre code demande `profile`, qui est réellement utilisé uniquement lorsque l’application cliente Office obtient le jeton à l’application web de votre complément. Seul `Files.Read.All` est demandé explicitement.

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

1. Remplacez `TODO 3` par le code suivant. Tenez compte du code suivant :

    - La méthode `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` recherchera tout d’abord dans le cache MSAL, c’est-à-dire en mémoire, un jeton d’accès correspondant. Uniquement s’il n’existe pas, elle lance le flux « de la part de » avec le point de terminaison Azure AD V2.
    - Les exceptions qui ne sont pas de type `MsalServiceException` ne sont intentionnellement pas capturées afin d’être propagées au client sous la forme de messages `500 Server Error`.

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

    - Si l’authentification multifacteur est requise par la ressource Microsoft Graph et que l’utilisateur ne l'a pas encore fournie, Azure AD renvoie « 400 : emande incorrecte » avec l’erreur `AADSTS50076` et une propriété **Claims**. MSAL génère une exception **MsalUiRequiredException** (qui hérite de **MsalServiceException**) avec ces informations.
    - La valeur de la propriété **Claims** doit être transmise au client qui doit la transmettre à l’application Office, qui l’inclut ensuite dans une demande de nouveau jeton d’amorçage. Azure AD demandera à l’utilisateur d’accepter tous les formulaires d’authentification requis.
    - The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.
    - Le message personnalisé est au format JSON pour que le code JavaScript côté client puisse l’analyser avec des méthodes d’objet `JSON` JavaScript connues.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. Remplacez `TODO 3b` par le code suivant. Tenez compte du code suivant :

    - Si l’appel à Azure AD contenait au moins une étendue (autorisation) pour laquelle ni l’utilisateur, ni un administrateur client a consenti (ou pour laquelle le consentement a été révoqué), Azure AD renvoie « 400 demande incorrecte » avec une erreur `AADSTS65001` MSAL génère une exception **MsalUiRequiredException** avec ces informations.
    - Si l’appel à Azure AD contenait au moins une étendue non reconnue par Azure AD, AAD renvoie « 400 Demande incorrecte » avec l’erreur `AADSTS70011`. MSAL génère une exception **MsalUiRequiredException** avec ces informations.
    - La description entière est incluse, car l’erreur 70011 est renvoyée dans d’autres conditions et elle doit être gérée dans ce complément uniquement lorsqu’elle indique une étendue non valide.
    - The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. Remplacez `TODO 3c` par le code suivant pour gérer toutes les autres **MsalServiceException** s.

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

    ![Choisissez l’application cliente Office souhaitée : Excel, PowerPoint ou Word.](../images/SelectHost.JPG)

1. Appuyez sur la touche F5.
1. Dans l’application Office, sur le ruban **Accueil**, sélectionnez **Afficher le complément** dans le groupe **ASP.NET SSO** pour ouvrir le complément du panneau des tâches.
1. Cliquez sur le bouton **Obtenir des noms de fichier OneDrive**. Si vous êtes connecté à Office avec un compte Microsoft 365 Éducation ou professionnel, ou un compte Microsoft, et que l’authentification unique fonctionne comme prévu, les 10 premiers noms de fichiers et de dossiers de votre OneDrive Entreprise s’affichent dans le volet Office. Si vous n’êtes pas connecté, ou si vous êtes dans un scénario qui ne prend pas en charge l’authentification unique, ou si l’authentification unique ne fonctionne pas pour une raison quelconque, vous serez invité à vous connecter. Une fois connecté, les noms des fichiers et des dossiers s’affichent.

### <a name="testing-the-fallback-path"></a>Test du chemin d’accès de secours

Pour tester le chemin d’autorisation de secours, forcez le chemin d’authentification unique à échouer en procédant comme suit.

1. Ajoutez le code suivant tout en haut de la `getDataWithToken` méthode dans le fichier HomeES6.js.

    ```javascript
    function MockSSOError(code) {
        this.code = code;
    }
    ```

1. Ajoutez ensuite la ligne suivante en haut du `try` bloc dans cette même méthode, juste au-dessus de l’appel à `getAccessToken`.

    ```javascript
    throw new MockSSOError("13003");
    ```

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>Mise à jour du complément lorsque vous passez à la préproduction et à la production

Comme tous les compléments Web Office, lorsque vous êtes prêt à passer à un serveur intermédiaire ou de production, vous devez mettre à jour le `localhost:44355` domaine dans le manifeste avec le nouveau domaine. De même, vous devez mettre à jour le domaine dans le fichier web.config.

Étant donné que le domaine apparaît dans l’inscription AAD, vous devez mettre à jour cette inscription pour utiliser le nouveau domaine à la place de `localhost:44355` où qu’il apparaisse.
