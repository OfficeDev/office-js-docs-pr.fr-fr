---
title: Vue d’ensemble de l’authentification et de l’autorisation dans les compléments Office
description: Découvrez le fonctionnement de l’authentification et de l’autorisation dans les compléments Office.
ms.date: 01/25/2022
ms.localizationpriority: high
ms.openlocfilehash: 1dab5e7e4cd1d5a32115bdecca3fa742699a53b9
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320122"
---
# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a>Vue d’ensemble de l’authentification et de l’autorisation dans les compléments Office

Les compléments Office autorisent l’accès anonyme par défaut, mais vous pouvez demander aux utilisateurs de se connecter pour utiliser votre complément avec un compte Microsoft, un compte Microsoft 365 Éducation ou professionnel ou un autre compte commun. Cette tâche est appelée authentification des utilisateurs, car elle permet au complément de déterminer l’identité de l’utilisateur.

Votre add-in peut également obtenir le consentement de l'utilisateur pour accéder à ses données Microsoft Graphique (telles que son profil Microsoft 365, ses fichiers OneDrive et ses données SharePoint) ou à des données d'autres sources externes telles que Google, Facebook, LinkedIn, SalesForce et GitHub. Cette tâche est appelée autorisation de complément (ou d’application), car il s’agit du *complément* qui est autorisé et non l’utilisateur.

## <a name="key-resources-for-authentication-and-authorization"></a>Ressources clés pour l’authentification et l’autorisation

Cette documentation explique comment créer et configurer des compléments Office pour implémenter correctement l’authentification et l’autorisation. Toutefois, de nombreux concepts et technologies de sécurité mentionnés ne sont pas concernés par cette documentation. Par exemple, les concepts de sécurité généraux tels que les flux OAuth, la mise en cache de jetons ou la gestion des identités ne sont pas expliqués ici. Cette documentation ne documente pas non plus quoi que ce soit spécifique à Microsoft Azure ou à la plateforme d’identités Microsoft. Nous vous recommandons de consulter les ressources suivantes si vous avez besoin d’informations supplémentaires dans ces domaines.

- [Plateforme d’identité Microsoft](/azure/active-directory/develop)
- [La prise en charge de la plateforme d’identités Microsoft et les options d’aide pour les développeurs](/azure/active-directory/develop/developer-support-help-options)
- [Protocoles OAuth 2.0 et OpenID Connect sur la plateforme d’identités Microsoft](/azure/active-directory/develop/active-directory-v2-protocols)

## <a name="sso-scenarios"></a>Scénarios d’authentification unique

L’utilisation de l’authentification unique (SSO) est pratique pour l’utilisateur, car il ne doit se connecter qu’une seule fois à Office. Ils n’ont pas besoin de se connecter séparément à votre complément. L’authentification unique n’étant pas prise en charge sur toutes les versions d’Office, vous devez toujours implémenter une autre approche de connexion, par [l’utilisation de la plateforme d’identités Microsoft](#authenticate-with-the-microsoft-identity-platform). Pour plus d’informations sur les versions d’Office prises en charge, consultez [Définir les conditions de l’API d’identité](../reference/requirement-sets/identity-api-requirement-sets.md)

### <a name="get-the-users-identity-through-sso"></a>Obtenir l’identité de l’utilisateur via l’authentification unique

Souvent, votre complément a uniquement besoin de l’identité de l’utilisateur. Par exemple, vous pouvez simplement personnaliser votre complément et afficher le nom de l’utilisateur dans le volet des tâches. Vous pouvez également souhaiter qu’un ID unique associe l’utilisateur à ses données dans votre base de données. Pour ce faire, il suffit d’obtenir le jeton d’accès pour l’utilisateur auprès d’Office.

Pour obtenir l’identité de l’utilisateur via l’authentification unique, appelez la méthode [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_). La méthode retourne un jeton d’accès qui est également un jeton d’identité contenant plusieurs revendications, unique à l’utilisateur connecté actuel, y compris `preferred_username`, `name`, `sub` et `oid`. Pour plus d’informations sur ces propriétés, consultez [jetons d’ID de la Plateforme d’identités Microsoft](/azure/active-directory/develop/id-tokens). Pour obtenir un exemple du jeton retourné par [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_), consultez [Exemple de jeton d’accès](sso-in-office-add-ins.md#example-access-token).

Si l’utilisateur n’est pas connecté, Office ouvrira une boîte de dialogue et utilise la plateforme d’identités Microsoft pour demander à l’utilisateur de se connecter. Ensuite, la méthode retournera un jeton d’accès ou génère une erreur si elle ne parvient pas à connecter l’utilisateur.

Dans un scénario où vous devez stocker des données pour l’utilisateur, reportez-vous à [jetons d’ID de la plateforme d’identités Microsoft](/azure/active-directory/develop/id-tokens) pour plus d’informations sur la façon d’obtenir une valeur à partir du jeton pour identifier l’utilisateur de manière unique. Utilisez cette valeur pour rechercher l’utilisateur dans une table utilisateur ou une base de données utilisateur que vous gérez. Utilisez la base de données pour stocker les informations relatives aux utilisateurs, comme les préférences utilisateur ou l’état du compte utilisateur. Étant donné que vous utilisez l’authentification unique, vos utilisateurs ne se connectent pas séparément à votre complément. vous n’avez donc pas besoin de stocker de mot de passe pour l’utilisateur.

Avant de commencer l’implémentation de l’authentification des utilisateurs avec l’authentification unique, assurez-vous que vous êtes familiarisé avec l’article [Activer l’authentification unique pour les compléments Office](sso-in-office-add-ins.md).

### <a name="access-your-web-apis-through-sso"></a>Accéder à vos API Web via l’authentification unique

Si votre complément a des API côté serveur qui nécessitent un utilisateur autorisé, appelez la méthode [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_) pour obtenir un jeton d’accès. Le jeton d’accès fournit l’accès à votre propre serveur web (configuré via un [inscription d’application Microsoft Azure ](register-sso-add-in-aad-v2.md).) Lorsque vous appelez des API sur votre serveur Web, vous transmettez également le jeton d’accès pour autoriser l’utilisateur.

Le code suivant montre comment construire une requête HTTPS GET vers l’API de serveur Web du complément pour obtenir des données. Le code s’exécute côté client, par exemple dans un volet Office. Il obtient d’abord le jeton d’accès en appelant `getAccessToken`. Il construit ensuite un appel AJAX avec l’en-tête et l’URL d’autorisation appropriés pour l’API serveur.

```javascript
function getOneDriveFileNames() {

    let accessToken = await Office.auth.getAccessToken();

    $.ajax({
        url: "/api/data",
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
        .done(function (result) {
            //... work with data from the result...
        });
}
```

Le code suivant montre un exemple de gestionnaire /api/data pour l’appel REST à partir de l’exemple de code précédent. Le code est ASP.NET, code s’exécutant sur un serveur Web. L’attribut `[Authorize]` nécessitera qu’un jeton d’accès valide soit passé à partir du client, ou il retournera une erreur au client.

```csharp
    [Authorize]
    // GET api/data
    public async Task<HttpResponseMessage> Get()
    {
        //... obtain and return data to the client-side code...
    }
```

### <a name="access-microsoft-graph-through-sso"></a>Accès Microsoft Graph via l’authentification unique

Dans certains scénarios, non seulement vous avez besoin de l’identité de l’utilisateur, mais vous devez également accéder aux ressources [Microsoft Graph](/graph) pour le compte de l’utilisateur. Par exemple, vous devrez peut-être envoyer un e-mail ou créer une conversation dans Teams pour le compte de l’utilisateur. Ces actions, et bien plus encore, peuvent être effectuées via Microsoft Graph. Vous devrez suivre ces étapes :

1. Obtenez le jeton d’accès pour l’utilisateur actuel via l’authentification unique en appelant [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_). Si l’utilisateur n’est pas connecté, Office ouvrira une boîte de dialogue et connectera l’utilisateur avec la plateforme d’identités Microsoft. Une fois que l’utilisateur se connecte, ou si l’utilisateur est déjà connecté, la méthode retourne un jeton d’accès.
1. Transmettez le jeton d’accès à votre code côté serveur.
1. Du côté serveur, utilisez le [flux on-Behalf-Of OAuth 2.0 On-Behalf-Of](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) pour échanger le jeton d’accès contre un nouveau jeton d’accès contenant l’identité d’utilisateur déléguée et les autorisations nécessaires pour appeler Microsoft Graph.

> [!NOTE]
> Pour une sécurité optimale afin d’éviter toute fuite du jeton d’accès, effectuez toujours le flux On-Behalf-Of côté serveur. Appelez Microsoft Graph API à partir de votre serveur, et non du client. Ne retournez pas le jeton d’accès au code côté client.

Avant de commencer à implémenter l’authentification unique pour accéder à Microsoft Graph dans votre complément, assurez-vous que vous connaissez bien les articles suivants.

- [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md)
- [Autoriser la connexion à Microsoft Graph avec l’authentification unique](authorize-to-microsoft-graph.md)

Vous devez également lire au moins l’un des articles suivants qui vous guideront dans la création d’un complément Office pour utiliser l’authentification unique et accéder à Microsoft Graph. Même si vous ne suivez pas la procédure, celles-ci contiennent des informations utiles sur la façon dont vous implémentez l’authentification unique et le flux On Behalf Of.

- [Créez un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md) qui vous guide dans l’exemple de [complément Office ASP.NET d’authentification unique](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).
- [Créez un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md) qui vous guide dans l’exemple vers [Office Add-in NodeJS SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

## <a name="non-sso-scenarios"></a>Scénarios autres que l’authentification unique

Dans certains scénarios, vous ne souhaiterez peut-être pas utiliser l’authentification unique. Par exemple, vous devrez peut-être vous authentifier à l’aide d’un fournisseur d’identité différent de celui de la plateforme d’identités Microsoft. En outre, l’authentification unique n’est pas prise en charge dans tous les scénarios. Par exemple, les versions antérieures d’Office ne prennent pas en charge l’authentification unique. Dans ce cas, vous devez revenir à un autre système d’authentification pour votre complément.

### <a name="authenticate-with-the-microsoft-identity-platform"></a>S’authentifier auprès de la Plateforme d’identités Microsoft.

Votre complément peut connecter des utilisateurs à l’aide de la [Plateforme d’identités Microsoft](/azure/active-directory/develop) en tant que fournisseur d’authentification. Une fois que vous êtes connecté à l’utilisateur, vous pouvez utiliser la plateforme d’identités Microsoft pour autoriser le complément de [Microsoft Graph](/graph) ou d’autres services gérés par Microsoft. Utilisez cette approche comme autre méthode de connexion lorsque l’authentification unique via Office n’est pas disponible. Il existe également des scénarios dans lesquels vous voulez que vos utilisateurs se connectent à votre complément séparément, même lorsque l’authentification unique est disponible. Par exemple, si vous voulez qu’ils aient la possibilité de se connecter au complément avec un ID différent de celui avec lequel ils sont actuellement connectés à Office.

Il est important de noter que la plateforme d’identités Microsoft n’autorise pas l’ouverture de sa page de connexion dans un iframe. Lorsqu’un complément Office est exécuté sur *Office sur le Web*, le volet des tâches est un IFrame. Cela signifie que vous devrez ouvrir la page de connexion à l’aide d’une boîte de dialogue ouverte avec l’API de boîte de dialogue Office. Cela a une incidence sur la manière dont vous utilisez les bibliothèques d’aide à l’authentification. Pour plus d’informations, consultez [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

Pour plus d’informations sur l’implémentation de l’authentification avec la plateforme d’identités Microsoft, consultez la[vue d’ensemble de la Plateforme d’identités Microsoft (v2.0)](/azure/active-directory/develop/v2-overview). La documentation contient de nombreux didacticiels et guides, ainsi que des liens vers des exemples et des bibliothèques pertinents. Comme expliqué dans [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md), vous devrez peut-être ajuster le code dans les exemples pour qu’il s’exécute dans la boîte de dialogue Office.

### <a name="access-to-microsoft-graph-without-sso"></a>Accès à Microsoft Graph sans authentification unique

Vous pouvez obtenir l’autorisation données Microsoft Graph pour votre complément en obtenant un jeton d’accès pour Microsoft Graph à partir de la plateforme d’identités Microsoft. Vous pouvez le faire sans vous appuyer sur l’authentification unique via Office (ou si l’authentification unique a échoué ou n’est pas prise en charge). Pour plus d’informations sur la manière de procéder, consultez [Accès à Microsoft Graph sans authentification unique](authorize-to-microsoft-graph-without-sso.md) qui contient davantage de détails et des liens vers des exemples.

### <a name="access-to-non-microsoft-data-sources"></a>Accès à des sources de données non-Microsoft

Les services en ligne populaires, dont Google, Facebook, LinkedIn, SalesForce et GitHub, permettent aux développeurs d’accorder aux utilisateurs l’accès à leurs comptes dans d’autres applications. Vous avez ainsi la possibilité d’inclure ces services dans votre complément Office. Pour obtenir une vue d’ensemble des méthodes que votre complément peut utiliser, voir [Autoriser des services externes dans votre complément Office](auth-external-add-ins.md).

> [!IMPORTANT]
> Avant de commencer le codage, déterminez si la source de données autorise l’ouverture de sa page de connexion dans un IFrame. Lorsqu’un complément Office est exécuté sur *Office sur le Web*, le volet des tâches est un IFrame. Si la source de données n’autorise pas l’ouverture de la page connexion dans un IFrame, vous devrez ouvrir la page de connexion dans une boîte de dialogue ouverte avec l’API de dialogue Office. Pour plus d’informations, consultez [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

## <a name="see-also"></a>Voir aussi

- [Documentation de la plateforme d’identités Microsoft](/azure/active-directory/develop/)
- [Jetons d’accès de plateforme d’identité Microsoft](/azure/active-directory/develop/access-tokens)
- [Protocoles OAuth 2.0 et OpenID Connect sur la plateforme d’identités Microsoft](/azure/active-directory/develop/active-directory-v2-protocols)
- [Plateforme d’identités Microsoft et flux OAuth 2.0 On-Behalf-Of](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
- [Jeton Web JSON (JWT)](https://en.wikipedia.org/wiki/JSON_Web_Token)
- [Visionneuse de jetons Web JSON](https://jwt.ms/)
