---
title: Activer l’authentification unique (SSO) dans un complément Office
description: Découvrez les étapes clés pour activer l’authentification unique (SSO) pour votre complément Office à l’aide de comptes Microsoft courants personnels, professionnels ou éducatifs.
ms.date: 05/05/2022
ms.localizationpriority: high
ms.openlocfilehash: 14b65da74cf627b7830ef013580558e8e6097ed1
ms.sourcegitcommit: fcb8d5985ca42537808c6e4ebb3bc2427eabe4d4
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2022
ms.locfileid: "65650597"
---
# <a name="enable-single-sign-on-sso-in-an-office-add-in"></a>Activer l’authentification unique (SSO) dans un complément Office

Les utilisateurs se connectent à Office à l’aide de leur compte Microsoft personnel ou de leur compte Microsoft 365 Education ou professionnel Profitez-en et utilisez l’authentification unique (SSO) pour authentifier et autoriser l’utilisateur à2 accéder à votre complément sans l’obliger à se connecter une deuxième fois.

![Image illustrant le processus de connexion pour un complément.](../images/sso-for-office-addins.png)

## <a name="how-sso-works-at-runtime"></a>Mode de fonctionnement de l’authentification unique SSO en cours d’exécution

Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique SSO. Les éléments bleus représentent Office ou la plateforme d’identité Microsoft. Les éléments gris représentent le code que vous écrivez et incluent le code côté client (volet des tâches) et le code côté serveur de votre complément.

:::image type="content" source="../images/sso-overview-diagram.svg" alt-text="Un diagramme illustrant le processus d’authentification unique." border="false":::

1. Dans le complément, votre code JavaScript appelle l’API Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)). Si l’utilisateur est déjà connecté à Office, l’hôte Office renverra le jeton d’accès avec les revendications de l’utilisateur connecté.
2. Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une boîte de dialogue permettant à l’utilisateur de se connecter. Office redirige vers la plateforme d’identité Microsoft pour terminer le processus de connexion.
3. Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.
4. L’application hôte Office demande le **jeton d’accès** à la plateforme d’identité Microsoft pour l’utilisateur actuel.
5. La plateforme d’identité Microsoft renvoie le jeton d’accès à Office. Office mettra le jeton en cache en votre nom afin que les futurs appels à **getAccessToken** renvoient simplement le jeton mis en cache.
6. L’application hôte Office renvoie le **jeton d’accès** au complément dans le cadre de l’objet de résultat renvoyé par l’appel `getAccessToken`.
7. Le jeton est à la fois un **jeton d’accès** et un **jeton d’identité**. Vous pouvez l’utiliser comme jeton d’identité pour analyser et examiner les revendications concernant l’utilisateur, telles que le nom et l’adresse e-mail de l’utilisateur.
8. Le complément peut éventuellement utiliser le jeton comme **jeton d’accès** pour envoyer des demandes HTTPS authentifiées aux API côté serveur. Étant donné que le jeton d’accès contient des revendications d’identité, le serveur peut stocker des informations associées à l’identité de l’utilisateur ; telles que les préférences de l’utilisateur.

## <a name="requirements-and-best-practices"></a>Meilleures Pratiques et Conditions Requises

### <a name="dont-cache-the-access-token"></a>Ne pas mettre en cache le jeton d’accès

Ne mettez jamais en cache ou ne stockez jamais le jeton d’accès dans votre code côté client. Appelez toujours [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) lorsque vous avez besoin d’un jeton d’accès. Office mettra en cache le jeton d’accès (ou en demandera un nouveau s’il a expiré). Cela aidera à éviter de divulguer accidentellement le jeton de votre complément.

### <a name="enable-modern-authentication-for-outlook"></a>Activer l’authentification moderne pour Outlook

Si vous travaillez avec un complément **Outlook**, assurez-vous d'activer l'authentification moderne pour la location de Microsoft 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

### <a name="implement-a-fallback-authentication-system"></a>Implémenter un système d’authentification de secours

Vous ne devez *pas* dépendre de l’authentification unique SSO comme seule méthode de votre complément d’authentification. Vous devez implémenter un système d’authentification secondaire vers lequel votre complément peut revenir dans certaines situations d’erreur. Par exemple, si votre complément est chargé sur une ancienne version d’Office qui ne prend pas en charge l’authentification unique, l’appel `getAccessToken` échouera.

Pour les compléments Excel, Word et PowerPoint, vous souhaiterez généralement utiliser la plate-forme d’identité Microsoft. Pour plus d’informations, voir [Authentification avec la plateforme d’identité Microsoft](overview-authn-authz.md#authenticate-with-the-microsoft-identity-platform).

Pour les compléments Outlook, il existe un système de secours recommandé. Pour plus d’informations, voir[Scénario : Implémenter l’authentification unique sur votre service dans un complément Outlook](../outlook/implement-sso-in-outlook-add-in.md).

Vous pouvez également utiliser un système de tables d’utilisateurs et d’authentification, ou vous pouvez tirer parti de l’un des fournisseurs de connexion sociale. Pour plus d’informations sur la procédure à suivre avec un complément Office, voir[Autoriser des services externes dans votre complément Office](auth-external-add-ins.md).

Pour des exemples de code qui utilisent la plate-forme d’identité Microsoft comme système de secours, consultez [Office Add-in NodeJS SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) et [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

## <a name="develop-an-sso-add-in"></a>Développer un complément d’authentification unique SSO

Cette section décrit les tâches impliquées dans la création d’un complément Office qui utilise l’authentification unique. Ces tâches sont décrites ici indépendamment du langage ou du framework. Pour obtenir des instructions détaillées, consultez :

- [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
- [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)

> [!NOTE]
> Vous pouvez utiliser le générateur Yeoman pour créer votre complément Office compatible avec l’authentification unique, Node.js.. Le générateur Yeoman simplifie le processus de création d’un complément avec authentification unique en automatisant les étapes nécessaires pour configurer l’authentification unique dans Azure et la génération du code nécessaire pour qu’un complément utilise l’authentification unique. Pour plus d'informations, consultez [Démarrage rapide de l'authentification unique](../quickstarts/sso-quickstart.md).

### <a name="register-your-add-in-with-the-microsoft-identity-platform"></a>Enregistrez votre complément auprès de la plateforme d’identité Microsoft

Pour travailler avec SSO, vous devez enregistrer votre complément auprès de la plateforme d’identité Microsoft. Cela permettra à la plateforme d’identité Microsoft de fournir des services d’authentification et d’autorisation pour votre complément. La création de l’enregistrement de l’application comprend les tâches suivantes.

- Obtenez un ID d’application (client) pour identifier votre complément sur la plateforme d’identité Microsoft.
- Générez un secret client pour agir comme mot de passe pour votre complément lors de la demande d’un jeton.
- Spécifiez les autorisations requises par votre complément. Les autorisations "profil" et "openid" de Microsoft Graph sont toujours requises. Vous aurez peut-être besoin d’autorisations supplémentaires en fonction de ce que votre complément doit faire.
- Accordez l’approbation des applications Office au complément.
- Pré-autorisez les applications Office sur le complément avec la portée par défaut *access_as_user*.

Pour plus de détails sur ce processus, voir [Inscrire un complément Office qui utilise l’authentification unique avec la plateforme d’identité Microsoft](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Configurer le complément

Ajoutez un nouveau balisage au manifeste du complément.

- **WebApplicationInfo**: le parent des éléments suivants.
- **ID** : ID d’application (client) que vous avez reçu lorsque vous avez enregistré le complément auprès de la plateforme d’identité Microsoft. Pour plus d’informations, voir [Inscrire un complément Office qui utilise l’authentification unique avec la plateforme d’identité Microsoft](register-sso-add-in-aad-v2.md).
- **Ressource** : URI du complément. Il s’agit du même URI (y compris le protocole `api:` ) que vous avez utilisé lors de l’enregistrement du complément auprès de la plateforme d’identité Microsoft. La partie domaine de cet URI doit correspondre au domaine, y compris tous les sous-domaines, utilisés dans les URL de la section `<Resources>` du manifeste du complément et l’URI doit se terminer par l’ID client spécifié dans l’élément `<Id>`.
- **Scopes**: le parent d’un ou plusieurs éléments **Scope**.
- **Étendue** – Spécifie une autorisation dont le complément a besoin. Les autorisations `profile` et `openID` sont toujours nécessaires et peuvent être les seules autorisations nécessaires. Si votre complément a besoin d’accéder à Microsoft Graph ou à d’autres ressources Microsoft 365, vous aurez besoin d’éléments d’**étendue** supplémentaires. Par exemple, pour les autorisations Microsoft Graph, vous pouvez demander les étendues `User.Read` et `Mail.Read`. Les biblioth?ques que vous utilisez dans votre code pour acc?der ? Microsoft Graph peuvent avoir des besoin d'autorisations suppl?mentaires. Pour plus d'informations, voir [Autoriser Microsoft Graph ? partir d'un compl?ment Office](authorize-to-microsoft-graph.md).

Pour les compléments Word, Excel et PowerPoint, ajoutez la balise à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour les compléments Outlook, ajoutez la balise à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

Voici un exemple de marques de révision.

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>openid</Scope>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

> [!NOTE]
> Si vous ne respectez pas les exigences de format dans le manifeste pour SSO, votre complément sera rejeté d’AppSource jusqu’à ce qu’il respecte le format requis.

### <a name="include-the-identity-api-requirement-set"></a>Inclure l’ensemble d’exigences de l’API Identity

Pour utiliser SSO, votre complément nécessite l’ensemble d’exigences Identity API 1.3. Pour plus d’informations, consultez [IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets).

### <a name="add-client-side-code"></a>Ajouter du code côté client

Ajoutez un code JavaScript pour le complément à :

- Appelez [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)).
- Analyser le jeton d’accès ou le transmettre au code côté serveur du complément.

Le code suivant montre un exemple simple d’appel `getAccessToken` et d’analyse du jeton pour le nom d’utilisateur et d’autres informations d’identification.

> [!NOTE]
> Cet exemple ne pr?sente explicitement qu'un seul type d'erreur. Pour des exemples de traitement des erreurs plus élaborés, voir [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) et [SSO ASP.NET pour complément Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

```js
async function getUserData() {
    try {
        let userTokenEncoded = await OfficeRuntime.auth.getAccessToken();
        let userToken = jwt_decode(userTokenEncoded); // Using the https://www.npmjs.com/package/jwt-decode library.
        console.log(userToken.name); // user name
        console.log(userToken.preferred_username); // email
        console.log(userToken.oid); // user id     
    }
    catch (exception) {
        if (exception.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // Microsoft 365 Education or work account, or a Microsoft account.
        } else {
            // Handle error
        }
    }
}
```


#### <a name="when-to-call-getaccesstoken"></a>Quand appeler getAccessToken

Si votre complément nécessite un utilisateur connecté, vous devez appeler `getAccessToken` depuis l’intérieur de `Office.initialize`. Vous devez également passer `allowSignInPrompt: true` le `options` paramètre de `getAccessToken`. Par exemple; `OfficeRuntime.auth.getAccessToken( { allowSignInPrompt: true });` Cela garantira que si l’utilisateur n’est pas encore connecté, Office invite l’utilisateur via l’interface utilisateur à se connecter maintenant.

Si le complément possède certaines fonctionnalités qui ne nécessitent pas d’utilisateur connecté, vous pouvez appeler `getAccessToken` *lorsque l’utilisateur effectue une action nécessitant un utilisateur connecté*. Il n’y a pas de dégradation significative des performances avec des appels redondants, `getAccessToken` car Office met en cache le jeton d’accès et le réutilisera, jusqu’à son expiration, sans effectuer un autre appel à la [plateforme d’identité Microsoft](/azure/active-directory/develop/) chaque fois que `getAccessToken` est appelé. Ainsi, vous pouvez ajouter des appels de `getAccessToken` à l’ensemble des fonctions et gestionnaires qui lancent une action dans laquelle le jeton est nécessaire.

> [!IMPORTANT]
> Comme meilleure pratique de sécurité, appelez toujours `getAccessToken` lorsque vous avez besoin d’un jeton d’accès. Office le mettra en cache pour vous. Ne mettez pas en cache ou ne stockez pas le jeton d’accès en utilisant votre propre code.

### <a name="pass-the-access-token-to-server-side-code"></a>Passer le jeton d’accès au code côté serveur

Si vous devez accéder aux API Web sur votre serveur ou à des services supplémentaires tels que Microsoft Graph, vous devrez transmettre le jeton d’accès à votre code côté serveur. Le jeton d’accès permet d’accéder (pour l’utilisateur authentifié) à vos API Web. De plus, le code côté serveur peut analyser le jeton pour les informations d’identité s’il en a besoin. (Voir **Utiliser le jeton d’accès comme jeton d’identité** ci-dessous.) Il existe de nombreuses bibliothèques disponibles pour différents langages et plateformes qui peuvent aider à simplifier le code que vous écrivez. Pour plus d’informations, consultez [Présentation de la bibliothèque d’authentification Microsoft (MSAL)](/azure/active-directory/develop/msal-overview).

Si vous devez accéder aux données Microsoft Graph, votre code côté serveur doit effectuer les opérations suivantes :

- Valider le jeton d’accès (voir **Valider du jeton d’accès** ci-dessous).
- Lancez le flux [OAuth 2.0 On-Behalf-Of](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) avec un appel à la plateforme d’identité Microsoft qui inclut le jeton d’accès, certaines métadonnées sur l’utilisateur et les informations d’identification du complément (son ID et son secret). La plateforme d’identité Microsoft renverra un nouveau jeton d’accès qui peut être utilisé pour accéder à Microsoft Graph.
- Obtenir des données à partir de Microsoft Graph en utilisant le nouveau jeton.
- Si vous devez mettre en cache le nouveau jeton d’accès pour plusieurs appels, nous vous recommandons d’utiliser la [sérialisation du cache de jeton dans MSAL.NET](/azure/active-directory/develop/msal-net-token-cache-serialization?tabs=aspnet).

> [!IMPORTANT]
> Comme meilleure pratique de sécurité, utilisez toujours le code côté serveur pour effectuer des appels Microsoft Graph ou d’autres appels nécessitant la transmission d’un jeton d’accès. Ne renvoyez jamais le jeton OBO au client pour permettre au client d’effectuer des appels directs vers Microsoft Graph. Cela aide à protéger le jeton contre l’interception ou la fuite. Pour plus d’informations sur le flux de protocole approprié, consultez le [diagramme de protocole OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#protocol-diagram)

Le code suivant montre un exemple de transmission du jeton d’accès côté serveur. Le jeton est transmis dans un en-tête `Authorization` lors de l’envoi d’une requête à une API Web côté serveur. Cet exemple envoie des données JSON, il utilise donc la méthode `POST`, mais `GET` est suffisant pour envoyer le jeton d’accès lorsque vous n’écrivez pas sur le serveur.

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + accessToken
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

Pour plus de d?tails sur l'obtention d'un acc?s autoris? aux donn?es Microsoft Graph de l'utilisateur, voir [Autoriser Microsoft Graph dans votre compl?ment Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Valider le jeton d’accès

Les API Web de votre serveur doivent valider le jeton d'accès s’il est envoyé depuis le client. Le jeton est un jeton JWT. En d’autres termes, la validation se déroule comme dans la plupart des flux OAuth standard. Il existe un certain nombre de bibliothèques pouvant gérer la validation JWT qui sont toutes, au minimum, chargées de :

- vérifier que le jeton est bien formé ;
- vérifier que le jeton a été émis par l’autorité souhaitée ;
- vérifier que le jeton est destiné à l’API web.

Suivez les recommandations suivantes quand vous validez le jeton.

- Les jetons SSO valides doivent être émis par l’autorité Azure `https://login.microsoftonline.com`. La revendication `iss` dans le jeton doit commencer par cette valeur.
- Le paramètre `aud` du jeton sera défini sur l’ID d’application de l’inscription de l’application Azure du complément.
- Le paramètre `scp` du jeton devra correspondre à `access_as_user`.

Pour plus d’informations sur la validation des jetons, consultez [Jetons d’accès à la plateforme d’identité Microsoft](/azure/active-directory/develop/access-tokens#validating-tokens).

#### <a name="use-the-access-token-as-an-identity-token"></a>Utiliser le jeton d'accès comme jeton d’identité

Si votre complément doit vérifier l’identité de l’utilisateur, le jeton d’accès retourné par `getAccessToken()` contient des informations qui peuvent être utilisées pour établir l’identité. Les revendications suivantes présentes dans le jeton concernent l’identité de l’utilisateur.

- `name`: le nom d’affichage de l’utilisateur.
- `preferred_username`: l’adresse de messagerie de l’utilisateur.
- `oid` – Un GUID représentant l’ID de l’utilisateur dans le système d’identité Microsoft.
- `tid` – Un GUID représentant le locataire auquel l’utilisateur se connecte.

Pour plus de détails sur ces revendications et d’autres, consultez [Jetons d’ID de plateforme d’identité Microsoft](/azure/active-directory/develop/id-tokens). Si vous devez créer un ID unique pour représenter l’utilisateur dans votre système, reportez-vous à [Utilisation des revendications pour identifier de manière fiable un utilisateur](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id) pour plus d’informations.

### <a name="example-access-token"></a>Exemple de token

Voici une charge utile d?cod?e typique de token. Pour plus d’informations sur les propriétés, voir [Jetons d’accès à la plateforme d’identité Microsoft](/azure/active-directory/develop/active-directory-v2-tokens).

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

## <a name="using-sso-with-an-outlook-add-in"></a>Utilisation de l’authentification unique SSO en accompagnement d’un complément Outlook

Il existe quelques différences mineures, mais importantes, en ce qui concerne l'utilisation de la connexion unique SSO dans un complément Outlook à partir de son utilisation dans un complément Excel, PowerPoint ou Word. Assurez-vous de lire [Authentifier un utilisateur avec un token unique log? dans le compl?ment Outlook](../outlook/authenticate-a-user-with-an-sso-token.md) et l' [?tude de cas : Impl?menter la connexion unique ? votre service dans un compl?ment Outlook](../outlook/implement-sso-in-outlook-add-in.md).

## <a name="see-also"></a>Voir aussi

- [Documentation de la plateforme d’identités Microsoft](/azure/active-directory/develop/)
- [Ensembles de conditions requises](specify-office-hosts-and-api-requirements.md)
- [IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets)