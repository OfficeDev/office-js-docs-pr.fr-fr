---
title: Vue d’ensemble de l’authentification et de l’autorisation dans les compléments Office
description: ''
ms.date: 08/09/2019
localization_priority: Priority
ms.openlocfilehash: dab5eec14a95aea9c27e1d26151b121ac2ed82ac
ms.sourcegitcommit: 24303ca235ebd7144a1d913511d8e4fb7c0e8c0d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2019
ms.locfileid: "36838507"
---
# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a>Vue d’ensemble de l’authentification et de l’autorisation dans les compléments Office

Les applications Web et par conséquent les compléments Office autorisent l’accès anonyme par défaut, mais vous pouvez demander aux utilisateurs de s’authentifier avec une connexion. Par exemple, vous pouvez demander à vos utilisateurs de se connecter à l’aide d’un compte Microsoft, d’un compte professionnel ou scolaire Office 365 ou d’un autre compte commun. Cette tâche est appelée authentification des utilisateurs, car elle permet au complément de déterminer l’identité de l’utilisateur.

Votre complément peut également obtenir l’autorisation de l’utilisateur à accéder à ses données Microsoft Graph (par exemple, son profil Office 365, ses fichiers OneDrive et ses données SharePoint) ou aux données d’autres sources externes comme Google, Facebook, LinkedIn, SalesForce et GitHub. Cette tâche est appelée autorisation de complément (ou d’application), car il s’agit du *complément* qui est autorisé et non l’utilisateur.

Vous avez le choix entre deux méthodes d’authentification.

- **Authentification unique (SSO) d’Office** : système *actuellement en préversion* qui permet à la connexion Office d’un utilisateur de fonctionner également comme connexion au complément. Si vous le souhaitez, le complément peut également utiliser les informations d’identification Office de l’utilisateur pour autoriser le complément à accéder à Microsoft Graph. (Les sources non-Microsoft ne sont pas accessibles par ce biais.)
- **Authentification et autorisation des applications web avec Azure Active Directory** : il ne s’agit pas d’une nouveauté, ni d’un comportement spécial. C’est la manière dont le complément Office (et les autres applications web) authentifiait les utilisateurs et autorisait les applications avant l’arrivée d’un système d’authentification unique pour Office et elle reste utilisée dans les scénarios où l’authentification unique d’Office est impossible.

Le diagramme suivant montre les décisions que vous devez prendre en tant que développeur de compléments. Cet article contient d’autres détails plus avant.

![Image illustrant un organigramme des décisions pour activer l’authentification et l’autorisation dans les compléments Office](../images/auth-decisions-flowchart.gif)

## <a name="user-authentication-without-sso"></a>Authentification utilisateur sans authentification unique

Vous pouvez authentifier un utilisateur dans un complément Office avec Azure Active Directory (AAD) comme vous le feriez dans n’importe quelle autre application web, à une exception près : AAD n’autorise pas la page de connexion à s’ouvrir dans un iFrame. Lorsqu’un complément Office est exécuté sur *Office sur le Web*, le volet Office est un iFrame. Cela signifie que vous devez ouvrir l’écran de connexion AAD dans une boîte de dialogue ouverte avec l’API de boîte de dialogue Office. Cela a une incidence sur la manière dont vous utilisez les bibliothèques d’aide à l’authentification. Pour plus d’informations, consultez [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

Pour plus d’informations sur la programmation de l’authentification avec AAD, commencez par la [vue d’ensemble de Microsoft Identity Platform (v 2.0)](/azure/active-directory/develop/v2-overview). Cet article présente de nombreux tutoriels et guides, ainsi que des liens vers des exemples et des bibliothèques pertinents. Comme expliqué dans [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md), vous devrez peut-être ajuster le code dans les exemples pour qu’il s’exécute dans la boîte de dialogue Office.

## <a name="access-to-microsoft-graph-without-sso"></a>Accès à Microsoft Graph sans authentification unique

Vous pouvez obtenir l’autorisation d’accès aux données Microsoft Graph pour votre complément en obtenant un jeton d’accès auprès d’Azure Active Directory (AAD). Vous pouvez effectuer cette opération sans utiliser l’authentification unique d’Office. Pour plus d’informations sur la manière de procéder, voir [Accès à Microsoft Graph sans authentification unique](authorize-to-microsoft-graph-without-sso.md) qui contient davantage de détails et des liens vers des exemples.

## <a name="user-authentication-with-sso"></a>Authentification utilisateur avec authentification unique

Pour utiliser l’authentification unique pour authentifier l’utilisateur, votre code dans le volet Office ou dans un fichier de fonction appelle la méthode [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-). Si l’utilisateur n’est pas connecté à Office, Office ouvre une boîte de dialogue et le redirige vers la page de connexion Azure Active Directory. Une fois l’utilisateur connecté, ou si l’utilisateur est déjà connecté, la méthode retourne un jeton d’accès. Ce jeton est un jeton d’amorçage dans le flux **On Behalf Of**. (Voir [accès à Microsoft Graph sans authentification unique](#access-to-microsoft-graph-with-sso).) Il peut toutefois être utilisé en tant que jeton d’ID, car il contient plusieurs revendications uniques pour l’utilisateur actuel, notamment `preferred_username`, `name`, `sub` et `oid`. Pour obtenir des instructions sur la propriété à utiliser en tant qu’ID utilisateur final, voir [Jetons d’accès à la plateforme d’identité Microsoft](https://docs.microsoft.com/fr-FR/azure/active-directory/develop/access-tokens#payload-claims). Pour obtenir un exemple d’un de ces jetons, consultez l’[exemple de jeton d’accès](sso-in-office-add-ins.md#example-access-token).

Une fois que votre code a extrait la revendication souhaitée du jeton, il utilise cette valeur pour rechercher l’utilisateur dans une table des utilisateurs ou une base de données des utilisateurs. Utilisez la base de données pour stocker les informations relatives aux utilisateurs, comme les préférences utilisateur ou l’état du compte utilisateur. Étant donné que vous utilisez l’authentification unique, vos utilisateurs ne se connectent pas séparément à votre complément. vous n’avez donc pas besoin de stocker de mot de passe pour l’utilisateur.

Avant de commencer l’implémentation de l’authentification des utilisateurs avec l’authentification unique, assurez-vous que vous êtes familiarisé avec l’article [Activer l’authentification unique pour les compléments Office](sso-in-office-add-ins.md). Notez également les exemples suivants :

- [Authentification unique NodeJS de complément Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), notamment le fichier [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) qui utilise la bibliothèque [jswebtoken](https://github.com/auth0/node-jsonwebtoken) pour décoder et analyser le jeton. (Toutefois, cet exemple n’utilise pas le jeton comme jeton d’identité. Il l’utilise pour obtenir l’accès à Microsoft Graph avec le flux **On Behalf Of**.)
- [Authentification unique ASP.NET de complément Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO), notamment le fichier [ValuesController.ts](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs) qui utilise la classe [System.Security.Claims.ClaimsPrincipal](https://docs.microsoft.com/dotnet/api/system.security.claims.claimsprincipal) de la bibliothèque pour extraire les revendications du jeton. (Toutefois, cet exemple n’utilise pas le jeton comme jeton d’identité. Il extrait une revendication `scope` du jeton et l’utilise pour obtenir l’accès à Microsoft Graph avec le flux **On Behalf Of**.)

## <a name="access-to-microsoft-graph-with-sso"></a>Accès à Microsoft Graph avec l’authentification unique

Pour utiliser l’authentification unique pour accéder à Microsoft Graph, votre complément dans le volet Office ou dans un fichier de fonction appelle la méthode [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-). Si l’utilisateur n’est pas connecté à Office, Office ouvre une boîte de dialogue et le redirige vers la page de connexion Azure Active Directory. Une fois l’utilisateur connecté, ou si l’utilisateur est déjà connecté, la méthode retourne un jeton d’accès. Ce jeton est un jeton d’amorçage dans le flux **On Behalf Of**. Plus précisément, il possède une `scope`revendication avec la valeur `access_as_user`. Pour plus d’informations sur les revendications dans le jeton, voir [Jetons d’accès à la plateforme d’identité Microsoft](https://docs.microsoft.com/fr-FR/azure/active-directory/develop/access-tokens#payload-claims). Pour obtenir un exemple d’un de ces jetons, consultez l’[exemple de jeton d’accès](sso-in-office-add-ins.md#example-access-token).

Une fois que votre code a obtenu le jeton, il l’utilise dans le flux **On Behalf Of** pour obtenir un deuxième jeton : un jeton d’accès à Microsoft Graph.

Avant de commencer l’implémentation de l’authentification unique Office, assurez-vous de bien vous familiariser avec les deux articles suivants :

- [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md)
- [Autoriser la connexion à Microsoft Graph avec l’authentification unique](authorize-to-microsoft-graph.md)

Vous devez également lire au moins l’un des articles de procédure de procédure mentionnés ici. Même si vous ne suivez pas la procédure, celle-ci contient des informations utiles sur la façon dont vous implémentez l’authentification unique Office et le flux **On Behalf Of**. 

- [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
- [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)

Notez également les exemples suivants :

- [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [SSO ASP.NET pour complément Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)

## <a name="access-to-non-microsoft-data-sources"></a>Accès à des sources de données non-Microsoft

Les services en ligne populaires, dont Google, Facebook, LinkedIn, SalesForce et GitHub, permettent aux développeurs d’accorder aux utilisateurs l’accès à leurs comptes dans d’autres applications. Vous avez ainsi la possibilité d’inclure ces services dans votre complément Office. Pour obtenir une vue d’ensemble des méthodes que votre complément peut utiliser, voir [Autoriser des services externes dans votre complément Office](auth-external-add-ins.md).

> [!IMPORTANT]
> Avant de commencer à coder, déterminez si la source de données autorise l’ouverture de son écran de connexion dans un iFrame. Lorsqu’un complément Office est exécuté sur *Office sur le Web*, le volet Office est un iFrame. Si la source de données n’autorise pas l’ouverture de l’écran de connexion dans un iFrame, vous devez ouvrir l’écran de connexion dans une boîte de dialogue ouverte avec l’API de boîte de dialogue Office. Pour plus d’informations, consultez [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

