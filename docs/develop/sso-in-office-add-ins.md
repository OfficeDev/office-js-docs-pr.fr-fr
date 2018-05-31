---
title: Activer l’authentification unique pour des compléments Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 45bd63150ffa8e46bf9c0fa54711ac907b8490ce
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437513"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Activer l’authentification unique pour des compléments Office (aperçu)

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365). Vous pouvez profiter de cette fonctionnalité et utiliser l’authentification unique (SSO) pour autoriser l’utilisateur à accéder à votre complément sans qu’il ne soit obligé de se connecter une seconde fois.


![Image illustrant le processus de connexion à un complément](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> L’API de l’authentification unique est actuellement prise en charge en mode aperçu pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).
> Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Pour les utilisateurs, cela permet une exécution aisée de votre complément qui ne requiert qu’une seule connexion. Pour les développeurs, cela signifie que votre complément n'a pas besoin de gérer ses propres tables utilisateur avec des mots de passe cryptés.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique.

![Diagramme illustrant le processus d’authentification unique](../images/sso-overview-diagram.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js `getAccessTokenAsync`. Cela indique à l’application hôte Office qu’elle doit obtenir un token auprès du complément. Voir [Exemple de token](#example-access-token).
2. Lorsque l'utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour lui permettre d'ouvrir une session.
3. Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.
4. L’application hôte Office demande le **jeton de complément** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.
5. Azure AD envoie le jeton de complément à l’application hôte Office.
6. L’application hôte Office envoie le **token du complément** au complément ; ce token constituant élément de l’objet de résultat renvoyé par l’appel `getAccessTokenAsync`.
7. Dans le complément, JavaScript peut analyser le token et extraire les informations dont il a besoin, telles que l'adresse e-mail de l'utilisateur. 
8. Optionnellement, le complément peut envoyer une requête HTTP à son serveur pour obtenir plus de données sur l'utilisateur, notamment les préférences de l'utilisateur. Alternativement, le token lui-même pourrait être envoyé au serveur pour analyse et validation. 

## <a name="develop-an-sso-add-in"></a>Développement d'un complément d’authentification unique

Cette section décrit les tâches impliquées dans la création d’un complément Office qui utilise l’authentification unique. Ces tâches sont décrites ici indépendamment du langage et de l’infrastructure. Pour obtenir des exemples de procédures pas à pas détaillées, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Création d'une application de service

Enregistrez le complément sur le portail d’inscription pour le point de terminaison Azure v2.0 :https://apps.dev.microsoft.com. Il s’agit d’un processus de 5 à 10 minutes qui inclut les tâches suivantes :

* Obtenir un ID client et un code secret pour le complément.
* La spécification des autorisations dont votre complément a besoin auprès du point de terminaison  Point de terminaison 2.0 (et éventuellement Microsoft Graph). L'autorisation "profil" est toujours nécessaire.
* Accordez la confiance de l’application hôte Office au complément.
* Pré-autorisez l’application hôte Office à accéder au complément à l'aide de l’autorisation par défaut *access_as_user*.

Pour plus de détails sur ce processus, voir [Enregistrer un complément Office qui utilise l'authentification unique auprès du point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Configuration du complément

Ajoutez un nouveau balisage au manifeste du complément :

* **WebApplicationInfo** : parent des éléments suivants.
* **Id** - ID du client du complément : il  s'agit d'un ID d'application que vous obtenez lors de l'enregistrement du complément. Voir [Enregistrer un complément Office avec l'authentification unique (SSO) auprès du point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md)
* **Ressource** : URL du complément.
* **Étendues** - parent d'un ou de plusieurs **éléments** d'étendues.
* **Étendue** : spécifie une autorisation nécessaire dont complément a besoin auprès d'AAD. L' `profile` autorisation est toujours nécessaire et il peut s'agir de la seule autorisation nécessaire si votre complément n'accède pas à Microsoft Graph. Si c'est le cas, vous avez également besoin des éléments d'une **étendue**pour obtenir les autorisations Microsoft Graph requises; par exemple, `User.Read`, `Mail.Read`. Les bibliothèques que vous utilisez dans votre code pour accéder à Microsoft Graph peuvent avoir des besoin d'autorisations supplémentaires. Par exemple, Microsoft Authentication Library (MSAL) pour .NET nécessite `offline_access` une autorisation. Pour plus d'informations, voir [Autoriser Microsoft Graph à partir d'un complément Office](authorize-to-microsoft-graph.md).

Pour les hôtes Office autres qu’Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

Voici un exemple de marques de révision :

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

Ajouter un code JavaScript au complément pour :

* Appeler [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).
* Analyser le token ou le transmettre au code côté serveur du complément. 

Voici un exemple simple d'un appel à `getAccessTokenAsync`. 

> [!Note]
> Cet exemple ne présente explicitement qu'un seul type d'erreur. Pour avoir des exemples de traitement des erreurs plus élaborés, voir [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) et [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). Voir également [Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

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

Voici un exemple simple d’un passage de token du complément vers le serveur. Le token est inclus en tant qu' `Authorization` en-tête lors de l'envoi d'une demande au serveur. Dans cet exemple, l'envoi de données JSON se fait en utilisant la méthode `POST`, mais `GET` est suffisant pour envoyer le token d'accès lorsque vous n'écrivez pas sur le serveur.

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

Si votre complément ne peut pas être utilisé lorsque aucun utilisateur n'est connecté à Office, vous devez appeler `getAccessTokenAsync` *lorsque le complément est lancé*.

Si le complément a une fonctionnalité qui ne nécessite pas d'utilisateur connecté, vous devez appeler `getAccessTokenAsync` *lorsque l'utilisateur effectue une action nécessitant un utilisateur connecté*. Les appels répétés à `getAccessTokenAsync` ne causent aucune dégradation importante des performances, car Office met en cache le token et le réutilise jusqu'à ce qu’il arrive à expiration, sans effectuer un autre appel vers le point de terminaison AAD v. 2.0 dès que `getAccessTokenAsync` est appelé. Ainsi, vous pouvez ajouter des appels de `getAccessTokenAsync` à l’ensemble des fonctions et gestionnaires qui lancent une action dans laquelle le token est nécessaire.

### <a name="add-server-side-code"></a>Ajouter du code côté serveur

Dans la plupart des scénarios, il n’est pas vraiment utile d’obtenir le jeton d’accès si votre complément ne le transmet pas côté serveur et ne l’utilise pas à cet emplacement. Quelques tâches côté serveur que votre complément pourrait faire :

* Créer d'une ou plusieurs méthodes d'API Web qui utilisent des informations sur l'utilisateur qui sont extraitent du token ; par exemple, une méthode qui recherche les préférences de l'utilisateur dans votre base de données hébergée. (Voir **Utilisation du token SSO en tant qu'identité** ci-dessous). En fonction de votre langue et de votre structure, des bibliothèques peuvent être disponibles pour simplifier le code que vous devez écrire.
* Obtenir des données Microsoft Graph. Votre code côté serveur doit effectuer les opérations suivantes :

    * Valider le token (voir **Valider le token** ci-dessous).
    * Démarrer le flux « de la part de » avec un appel du point de terminaison Azure AD v2.0 qui inclut le token du complément, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et code secret). Dans ce contexte, le token est appelé token de démarrage.
    * Mettre en cache le nouveau token renvoyé par le flux « de la part de ».
    * Obtenir des données à partir de Microsoft Graph en utilisant le nouveau token.

 Pour plus de détails sur l'obtention d'un accès autorisé aux données Microsoft Graph de l'utilisateur, voir [Autoriser Microsoft Graph dans votre complément Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Valider le token

Lorsque l’API web reçoit le token, elle doit valider son fonctionnement avant de l’utiliser. Le jeton est un jeton JWT. En d’autres termes, la validation se déroule comme dans la plupart des flux OAuth standard. Il existe un certain nombre de bibliothèques pouvant gérer la validation JWT qui sont toutes, au minimum, chargées de :

- vérifier que le jeton est bien formé ;
- vérifier que le jeton a été émis par l’autorité souhaitée ;
- vérifier que le jeton est destiné à l’API web.

Suivez les recommandations suivantes quand vous validez le jeton :

- Les jetons SSO valides doivent être émis par l’autorité Azure `https://login.microsoftonline.com`. La revendication `iss` dans le jeton doit commencer par cette valeur.
- Le paramètre `aud` du jeton devra correspondre à l’ID d’application de l’enregistrement du complément.
- Le paramètre `scp` du jeton devra correspondre à `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>Utilisation du jeton SSO comme identité

Si votre complément doit vérifier l’identité de l’utilisateur, le jeton SSO contient des informations utiles pour établir son identité. Les revendications suivantes présentes dans le jeton concernent l’identité de l’utilisateur.

- `name` : nom d’affichage de l’utilisateur.
- `preferred_username` Adresse e-mail de l'utilisateur
- `oid` : GUID représentant l’ID de l’utilisateur dans Azure Active Directory.
- `tid` - GUID représentant l’ID de l’organisation de l’utilisateur dans Azure Active Directory.

Depuis le `name` et `preferred_username` les valeurs pourraient changer, nous recommandons que le `oid` et les valeurs `tid` soient utilisées pour corréler l'identité avec le service d'autorisation de votre back-end.

Par exemple, votre service peut mettre en forme ces valeurs de la façon suivante `{oid-value}@{tid-value}`, puis stocker cette mise en forme sous forme de valeur dans l’enregistrement de l’utilisateur dans votre base de données utilisateur interne. Lors des demandes ultérieures, l’utilisateur pourra être récupéré grâce à cette valeur et l’accès à certaines ressources pourra être déterminé selon les mécanismes de contrôle d’accès existants.

### <a name="example-access-token"></a>Exemple de token

Voici une charge utile décodée typique de token. Pour plus d'informations sur les propriétés, voir [Référence des tokens Azure Active Directory v2.0](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-tokens).


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

## <a name="using-sso-with-and-outlook-add-in"></a>Utilisation de SSO en accompagnement et du complément Outlook

Il existe quelques différences mineures, mais importantes, en ce qui concerne l'utilisation de la connexion unique SSO dans et comme complément Outlook comparé à son utilisation comme complément Excel, PowerPoint ou Word. Assurez-vous de lire [Authentifier un utilisateur avec un token unique logé dans le complément Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-sso-token) et l' [étude de cas : Implémenter la connexion unique à votre service dans un complément Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in).