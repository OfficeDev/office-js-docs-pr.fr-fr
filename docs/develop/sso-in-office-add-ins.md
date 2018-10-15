---
title: Activer l’authentification unique pour des compléments Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: fb4eacee9419339116e15ef3fccc03b291faf3ec
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506027"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Activer l’authentification unique pour des compléments Office (préversion)

Les utilisateurs d'Office se connectent à leur service à l’aide de leur compte Microsoft personnel, professionnel ou scolaire. Vous pouvez profiter de cette procédure et utiliser notre système d'authentification unique (SSO) pour permettre à l’utilisateur de votre complément de se connecter sans avoir a s'identifier une nouvelle fois.

![Image illustrant le processus de connexion pour un complément](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>Mode de prévisualisation

L’API d'authentification unique (SSO) est prise en charge seulement en préversion. Elle est disponible pour les développeurs pour expérimenter ; mais elle ne doit pas être utilisé dans un complément en production. En outre, les compléments qui utilisent l’authentification unique ne sont pas acceptés dans [AppSource](https://appsource.microsoft.com).

Certaines applications sur Office ne prennent pas le mode prévisualisation SSO. Ce mode est disponible dans Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’API SSO et les applications qui la supporte, veuillez consulter [Exigences pour IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).

### <a name="requirements-and-best-practices"></a>Conditions requises et meilleures pratiques

Pour utiliser l’authentification unique, vous devez charger la version bêta de la bibliothèque de JavaScript Office `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` dans la page de démarrage HTML du complément.

Pour utiliser l’authentification unique avec un complément **Outlook** , vous devez activer l’authentification moderne pour la location Office 365. Pour plus d’informations sur la procédure à suivre, consultez [Exchange Online : activation de l’authentification moderne pour votre client](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Ne dépendez *pas*  seulement sur le système SSO comme méthode unique d’authentification pour votre complément. Vous devez aussi implémenter un autre  système d’authentification au cas ou une erreur surviendrait. Par exemple, vous pouvez utiliser un système d’authentification et de tables utilisateur, ou vous pouvez exploiter l'un des fournisseurs de connexion de mise en réseau. Pour plus d’informations sur la procédure a suivre, consultez [Autorisation de services externes dans votre complément pour Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). Pour *Outlook*, il existe un système de secours recommandé. Pour plus d’informations sur ce système, consultez [Scénario : implémentation de l’authentification unique à votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

### <a name="how-sso-works-at-runtime"></a>Fonctionnement de l’authentification unique au moment de l’exécution

Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique.

![Diagramme illustrant le processus d’authentification unique](../images/sso-overview-diagram.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessTokenAsync](#sso-api-reference). Cela indique à l'application hôte Office d'obtenir un jeton d'accès au complément. Voir [Exemple de jeton d'accès](#example-access-token).
2. Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.
3. Si c’est la première fois que l’utilisateur utilise votre complément, il est invité à donner son consentement.
4. L’application hôte Office demande le **jeton de complément** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.
5. Azure AD envoie le jeton de complément à l’application hôte Office.
6. L’application hôte Office envoie le **jeton de complément** au complément dans le cadre de l’objet de résultat renvoyé par l’appel `getAccessTokenAsync`.
7. Dans le complément, JavaScript peut analyser le jeton et extraire les informations dont il a besoin, telles que l'adresse e-mail de l'utilisateur. 
8. Si vous le souhaitez, le complément peut envoyer une demande HTTP à son côté serveur pour plus de données relatives à l’utilisateur ; comme les préférences de l’utilisateur. Le jeton d’accès pourrait également être envoyé au côté serveur pour être analysé et validé. 

## <a name="develop-an-sso-add-in"></a>Développer un complément d’authentification unique

Cette section décrit les tâches impliquées dans la création d’un complément pour Office qui utilise l’authentification unique. Ces tâches seront décrites indépendamment  de la langue ou de la structure que vous utilisez. Pour obtenir des exemples de procédures détaillées, consultez :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Créer l’application de service

Enregistrez le complément sur le portail d’inscription pour le point de terminaison Azure v2.0 :https://apps.dev.microsoft.com. Il s’agit d’un processus de 5 à 10 minutes qui inclut les tâches suivantes :

* Obtenir un identificateur de client et une clé secrète pour le complément.
* Spécifier les autorisations nécessaires à votre complément pour le point de terminaison AADv. 2.0 (et facultativement pour Microsoft Graph). Ce « profil » d'autorisation est toujours nécessaire.
* Attribuez à votre complément la possibilité de faire confiance à l'application hôte d'Office.
* Pré-autorisez l’application hôte Office pour le complément avec l’autorisation par défaut *access_as_user*.

Pour plus de détails sur ce processus, voir [Enregistrer un complément Office qui utilise l'authentification unique auprès du point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Configurez le complément

Ajoutez un nouveau balisage au manifeste du complément :

* **WebApplicationInfo** : le parent des éléments suivants.
* **Id** - l’identificateur du client du complément. C'est un identificateur d’application que vous obtenez dans le cadre de l'enregistrement du complément. Voir [Enregistrer un complément Office qui utilise l’authentification unique avec le point de terminaison v2.0 Azure AD](register-sso-add-in-aad-v2.md).
* **Resource** : l'URL du complément.
* **Scopes** : parent d’un ou plusieurs éléments **Scope**.
* **Scope** - spécifie une autorisation dont le complément a besoin pour AAD. L' `profile` autorisation est toujours nécessaire et peut être la seule autorisation nécessaire, si votre complément n’a pas accès à Microsoft Graph. Si c’est le cas, vous devez également des éléments **Scope** pour les autorisations de Microsoft Graph requises ; par exemple, `User.Read`, `Mail.Read`. Les bibliothèques à utiliser dans votre code pour accéder à Microsoft Graph peuvent avoir besoin d'autorisations supplémentaires. Par exemple, la bibliothèque de l’authentification de Microsoft (MSAL) pour .NET nécessite `offline_access` autorisation. Pour plus d’informations, consulter [Autoriser Microsoft Graph à partir d’un complément Office](authorize-to-microsoft-graph.md).

Pour les hôtes Office autres qu’Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

Voici un exemple de balise :

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a>Ajouter du code côté client

Ajoutez un code JavaScript pour le complément à :

* Appelez [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).

* Analyser le jeton ou le transmettre au code côté serveur du complément. 

Voici un exemple simple d'un appel à `getAccessTokenAsync`. 

> [!NOTE]
> Cet exemple gère explicitement un seul type d’erreur. Pour obtenir des exemples de gestion d'erreurs plus complexes, consultez [Home.js dans Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) et [program.js dans Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). Vous pouvez aussi consulter [Résoudre les messages d’erreur pour l’authentification unique (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

Voici un exemple simple de transmission de jeton du complément au côté serveur. Le jeton est inclus en tant qu’un `Authorization` en-tête lors de l’envoi d’une demande au côté serveur. Cet exemple envisage l’envoi des données JSON, donc il utilise la méthode  `POST` , mais `GET` est suffisant pour envoyer le jeton d’accès lorsque vous n'écrivez pas sur le serveur.

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a>Quand appeler la méthode

Si votre complément ne peut pas être utilisé quand aucun utilisateur n’est connecté à Office, vous devez appeler `getAccessTokenAsync` *au lancement du complément*.

Si le complément a certaines fonctionnalités qui ne nécessitent pas une connexion, alors vous devez appeler `getAccessTokenAsync` *lorsque l’utilisateur exécute une action qui nécessite d’être connecté*. Il n’existe aucune dégradation significative de performances avec des appels superflus `getAccessTokenAsync` , car Office met en cache le jeton d’accès pour le réutiliser, jusqu'à ce qu’il expire, sans effectuer un autre appel au point de terminaison de AAD v. 2.0 chaque fois que `getAccessTokenAsync` est appelée. Ainsi, vous pouvez ajouter des appels de `getAccessTokenAsync` à toutes les fonctions et les gestionnaires qui déclenchent une action dans laquelle le jeton est nécessaire.

### <a name="add-server-side-code"></a>Ajouter du code côté serveur

Dans la plupart des scénarios, il n'y a guère de raison d'obtenir le jeton d’accès, si votre complément ne le transmet pas du côté serveur pour l'utiliser. Voici certaines tâches côté serveur que votre complément peut faire :

* Créer une ou plusieurs méthodes pour l’API Web qui utilisent des informations sur l’utilisateur extrait du jeton ; par exemple, une méthode qui recherche des préférences de l’utilisateur dans votre base de données. (Voir **Utiliser le jeton d’authentification unique comme une identité** ci-dessous). En fonction de votre langue et de votre infrastructure, certaines bibliothèques peuvent être disponibles qui simplifieront le codage à effectuer.
* Obtenir des données de Microsoft Graph. Votre code côté serveur devrait procéder comme suit :

    * Valider le jeton (voir **Valider le jeton d’accès** ci-dessous).
    * Démarrer le flux « de la part de » par un appel au point de terminaison Azure AD v2.0 qui inclut le jeton d’accès, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (son identificateur et sa clé secrète). Dans ce contexte, le jeton d’accès est appelé le jeton d’amorçage.
    * Mettre en cache le nouveau jeton renvoyé le flux « de la part de ».
    * Obtenir des données à partir de Microsoft Graph en utilisant le nouveau jeton.

 Pour plus de détails sur l'obtention d'un accès autorisé aux données Microsoft Graph de l'utilisateur, voir [Autoriser Microsoft Graph dans votre complément Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Valider le jeton d’accès

Une fois que l’API Web reçoit le jeton d’accès, elle doit le valider avant de l'utiliser. Le jeton est un JSON Web Token (JWT), ce qui signifie que la validation fonctionne comme la validation des jetons dans la plupart des flux OAuth. Il existe un certain nombre de bibliothèques qui peuvent gérer la validation de JWT, mais les concepts de base sont :

- Vérifier que le jeton est bien formé ;
- vérifier que le jeton a été émis par l’autorité souhaitée ;
- vérifier que le jeton est destiné à l’API web.

Suivez les recommandations suivantes quand vous validez le jeton :

- Les jetons d’authentification unique valides seront délivrés par l’autorité Azure, `https://login.microsoftonline.com`. La revendication `iss` dans le jeton doit commencer par cette valeur.
- Le paramètre `aud` du jeton devra correspondre à l’identificateur d’application de l’enregistrement du complément.
- Le paramètre `scp` du jeton devra correspondre à `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>L'utilisation du jeton SSO comme identité

Si votre complément doit vérifier l’identité de l’utilisateur, le jeton d’authentification unique contient des informations qui peuvent être utilisées pour établir l’identité. Ces revendications dans le jeton sont associées à l’identité:

- `name` - le nom de l’utilisateur ;
- `preferred_username` -  l'adresse e-mail de l'utilisateur;
- `oid` - un GUID représentant l'identificateur de l’utilisateur dans Azure Active Directory;
- `tid` - un GUID représentant l'ID de l'organisation de l'utilisateur dans l'Azure Active Directory.

Étant donné que les valeurs `name` et `preferred_username` peuvent être amenées à changer, nous vous recommandons d’utiliser les valeurs `oid` et `tid` pour corréler l’identité de l’utilisateur avec le service d’autorisation de votre API.

Par exemple, votre service peut mettre rassembler ces valeurs comme `{oid-value}@{tid-value}`, puis les stocker en tant que valeur dans l’entrée de l’utilisateur dans votre base de données d'utilisateurs internes. Ainsi pour les demandes ultérieures, l’utilisateur pourra être récupéré à l’aide de la même valeur. De même, l'accès à des ressources spécifiques pourra être déterminé en fonction de vos mécanismes de contrôle d’accès existant.

### <a name="example-access-token"></a>Exemple de jeton d’accès

Vous trouverez ci-dessous une charge décodée typique d’un jeton d’accès. Pour plus d’informations sur les propriétés, consulter [référence de jetons Azure Active Directory v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",         
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",         
    iat: 1521143967,         
    nbf: 1521143967,         
    exp: 1521147867,         
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",         
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",         
    azpacr: "0",         
    e_exp: 262800,         
    name: "Mila Nikolova",         
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",         
    preferred_username: "milan@contoso.com",         
    scp: "access_as_user",         
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",         
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",         
    uti: "MICAQyhrH02ov54bCtIDAA",         
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a>L'utilisation de l'authentification unique (SSO) avec un complément Outlook

Il y a quelques petites, mais importantes différences entre l'utilisation du SSO dans un complément Outlook et son utilisation dans un complément Excel, PowerPoint ou Word. Veillez lire [Authentifier un utilisateur avec un jeton d’authentification unique dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) et [Scénario : Implémenter la connexion unique à votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

## <a name="sso-api-reference"></a>Référence de l’API de l’authentification unique

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

L’espace de noms Office Auth, `Office.context.auth`, fournit une méthode, `getAccessTokenAsync` qui permet à l’hôte Office d'obtenir un jeton d’accès a l'application web de votre complément. Indirectement, cela permet également le complément d'accéder aux données de Microsoft Graph de l’utilisateur connecté sans que l’utilisateur aie se connecter une deuxième fois.

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

La méthode contacte le point de terminaison Azure Active Directory V 2.0 pour obtenir un jeton d’accès à l’application web de votre complément. Cela permet aux compléments d’identifier les utilisateurs. Le code côté serveur peut utiliser ce jeton pour accéder à Microsoft Graph pour l’application web du complément à l’aide du [flux d’OAuth « de la part de »](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> Dans Outlook, cette API n'est pas prise en charge si le complément est chargé dans une boîte aux lettres Outlook.com ou Gmail.

<table><tr><td>Hôtes</td><td>Excel, OneNote, Outlook, PowerPoint, Word</td></tr>

 <tr><td>Ensembles de conditions requises</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a>Paramètres

`options` - Facultatif. Accepte un objet `AuthOptions` (voir ci-dessous) pour définir les comportements de connexion.

`callback` -Facultatif. Accepte une méthode de rappel qui peut analyser le jeton pour l’identificateur d’utilisateur ou utiliser le jeton dans le flux « de la part de » pour accéder à Microsoft Graph. Si  [ AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status`  a « réussi », alors `AsyncResult.value`est le jeton d’accès au format AAD v. 2.0 brut.

L'interface  `AuthOptions` fournit des options pour l'expérience utilisateur, lorsque Office obtient un jeton d'accès au complément d'AAD v. 2.0 avec lamethod `getAccessTokenAsync` .

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



