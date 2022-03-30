---
title: 'Scénario: implémenter l’authentification unique dans votre service'
description: Découvrez comment utiliser le jeton d’authentification unique et le jeton d’identité Exchange fournis par un complément Outlook afin d’implémenter l’authentification unique (SSO) pour votre service.
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2b9c4031a0011d2333582b4a10abe42f6844f763
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496921"
---
# <a name="scenario-implement-single-sign-on-to-your-service-in-an-outlook-add-in"></a>Scénario : Implémenter l’authentification unique sur votre service dans un complément Outlook

Dans cet article, nous allons vous expliquer comment utiliser le [jeton d’accès à authentification unique](authenticate-a-user-with-an-sso-token.md) et le [jeton d’identité Exchange](authenticate-a-user-with-an-identity-token.md) pour implémenter une authentification unique sur votre service principal. En utilisant ces jetons, vous pouvez profiter des avantages du jeton d’accès SSO quand il est disponible, tout en garantissant le fonctionnement de votre complément quand il ne l’est pas, par exemple quand l’utilisateur bascule vers un client qui ne les prend pas en charge , ou quand la boîte aux lettres de l’utilisateur se trouve sur un serveur Exchange local.

Pour obtenir un exemple de add-in qui implémente les idées de cet article, voir Outlook [SSO de l’application](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).


> [!NOTE]
> L’API d’authentification unique est actuellement prise en charge pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’emplacement où l’API d’authentification unique est actuellement prise en charge, consultez [ensembles de conditions requises IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location Microsoft 365. Pour plus d’informations sur la procédure à suivre, consultez [Exchange Online : comment activer votre locataire pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).


## <a name="why-use-the-sso-access-token"></a>Pourquoi utiliser le jeton d’accès SSO ?

Le jeton d’identité Exchange étant demandé dans tous les ensembles de conditions requises de l’API du complément, il peut être tentant d’utiliser uniquement ce jeton. Toutefois, le jeton SSO présente des avantages par rapport au jeton d’identité Exchange, c’est pourquoi nous vous recommandons de l’utiliser quand il est disponible.

- Le jeton SSO utilise un format OpenID standard et est émis par Azure, ce qui simplifie considérablement le processus de validation de ces jetons. En revanche, les jetons d’identité Exchange utilisent un format personnalisé basé sur la norme JWT (JSON Web Token), ce qui nécessite un travail de personnalisation pour valider ce jeton.
- Le jeton SSO peut être employé par votre service principal pour récupérer un jeton d’accès à Microsoft Graph afin d’éviter à l’utilisateur d’entrer une nouvelle fois ses informations d’identification.
- Le jeton SSO fournit des informations d’identité plus riche, comme le nom d’affichage de l’utilisateur.

## <a name="add-in-scenario"></a>Scénario du complément

Prenons l’exemple d’un complément composé d’une interface utilisateur et de scripts (HTML + JavaScript), et de l’API web principale appelée par le complément. L’API web principale appelle à la fois l’[API Microsoft Graph](/graph/overview) et l’API Contoso Data, une API fictive créée par un tiers. Tout comme l’API Microsoft Graph, l’API Contoso Data nécessite une authentification OAuth, qui exige que l’API web principale soit capable d’appeler les deux API sans inviter l’utilisateur à renseigner ses informations d’identification après chaque expiration du jeton d’accès.

Pour cela, l’API principale crée une base de données utilisateurs sécurisée. Chaque utilisateur obtient une entrée dans la base de données où l’API principale peut stocker des jetons d’actualisation à durée de vie longue pour l’API Microsoft Graph et l’API Contoso Data. Les marques JSON suivantes représentent l’entrée d’un utilisateur dans la base de données.

```JSON
{
  "userDisplayName": "...",
  "ssoId": "...",
  "exchangeId": "...",
  "graphRefreshToken": "...",
  "contosoRefreshToken": "..."
}
```

Le complément inclut le jeton d’accès SSO (s’il est disponible) ou le jeton d’identité Exchange (si le jeton SSO n’est pas disponible), ainsi que tous les appels passés à l’API web principale.

### <a name="add-in-startup"></a>Démarrage du complément

1. Quand le complément démarre, il envoie une demande à l’API web principale pour déterminer si l’utilisateur est enregistré (c’est-à-dire, s’il est associé à un enregistrement dans la base de données utilisateur) et si l’API dispose de jetons d’actualisation pour Graph et Contoso. Dans cet appel, le complément inclut le jeton SSO (s’il est disponible) et le jeton d’identité.

1. L’API web utilise les méthodes décrites dans les rubriques [Authentifier un utilisateur avec un jeton d’authentification unique dans un complément Outlook](authenticate-a-user-with-an-sso-token.md) et [Authentifier un utilisateur avec un jeton d’identité pour Exchange](authenticate-a-user-with-an-identity-token.md) pour valider et générer un identificateur unique à partir des deux jetons.

1. Si un jeton SSO est fourni, l’API web recherche dans la base de données utilisateur une entrée dont la valeur `ssoId` correspond à l’identificateur unique généré à partir du jeton SSO.
   - Si aucune entrée ne correspond, passez à l’étape suivante.
   - Si l’API trouve cette entrée, passez à l’étape 5.

1. L’API web recherche dans la base de données une entrée dont la valeur `exchangeId` correspond à l’identificateur unique généré à partir du jeton d’identité Exchange.
   - Si aucune entrée ne correspond et qu’un jeton SSO a été fourni, mettez à jour l’enregistrement de l’utilisateur dans la base de données pour que la valeur `ssoId` corresponde à l’identificateur unique généré à partir du jeton SSO, puis passez à l’étape 5.
   - Si l’API trouve une entrée et qu’aucun jeton SSO n’a été fourni, passez à l’étape 5.
   - Si aucune entrée ne correspond, créez une entrée. Attribuez à la valeur `ssoId` l’identificateur unique généré à partir de jeton SSO (s’il est disponible) et attribuez à la valeur `exchangeId` l’identificateur unique généré à partir du jeton d’identité Exchange.

1. Recherchez un jeton d’actualisation valide dans la valeur `graphRefreshToken` de l’utilisateur.
   - Si la valeur est manquante ou non valide et qu’un jeton SSO a été fourni, utilisez le [flux Pour le compte de Oauth2](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) pour obtenir un jeton d’accès et un jeton d’actualisation pour Microsoft Graph. Enregistrez le jeton d’actualisation dans la valeur `graphRefreshToken` de l’utilisateur.

1. Recherchez les jetons d’actualisation valides dans les valeurs `graphRefreshToken` et `contosoRefreshToken`.
   - Si ces deux valeurs sont valides, répondez au complément pour indiquer que l’utilisateur est déjà enregistré et configuré.
   - Si l’une de ces valeurs n’est pas valide, répondez au complément pour indiquer que l’utilisateur doit être installé et signaler les services (Graph ou Contoso) à configurer.

1. Le complément vérifie la réponse.
   - Si l’utilisateur est déjà enregistré et configuré, le complément continue de fonctionner normalement.
   - Si l’utilisateur doit être configuré, le complément passe en mode Installation et invite l’utilisateur à autoriser l’accès au complément.

### <a name="authorize-the-backend-web-api"></a>Autorisation de l’API web principale

Dans l’idéal, il conviendrait d’autoriser une seule fois l’API web principale à appeler l’API Microsoft Graph et l’API Contoso Data, pour éviter d’inviter l’utilisateur à se connecter à chaque fois.

En fonction de la réponse de l’API web principale, le complément doit autoriser l’utilisateur à accéder à l’API Microsoft Graph et/ou à l’API Contoso Data. Étant donné que les deux API ont recours à l’authentification OAuth2, la même méthode doit être employée pour les deux.

1. Le complément informe l’utilisateur qu’il doit autoriser l’utilisation de l’API et lui demande de cliquer sur un lien ou un bouton pour démarrer la procédure.

    > [!NOTE]
    > L’exemple de add-in de [Outlook Add-in SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) indique comment utiliser [l’API de](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) dialogue et la bibliothèque [office-js-helpers](https://github.com/OfficeDev/office-js-helpers) comme options pour démarrer le flux de code d’autorisation [OAuth2](/azure/active-directory/develop/active-directory-protocols-oauth-code) pour l’API.

1. Une fois le flux terminé, le complément envoie le jeton d’actualisation à l’API web principale et inclut le jeton SSO (s’il est disponible) ou le jeton d’identité Exchange.

1. L’API web principale recherche l’utilisateur dans la base de données et met à jour le jeton d’actualisation approprié.

1. Le complément continue de fonctionner normalement.

### <a name="normal-operation"></a>Conditions normales de fonctionnement

Quand le complément appelle l’API web principale, il inclut le jeton SSO ou le jeton d’identité Exchange. L’API web principale localise l’utilisateur en fonction de ce jeton, puis utilise les jetons d’actualisation stockés pour obtenir les jetons d’accès à l’API Microsoft Graph et à l’API Contoso Data. Tant que les jetons d’actualisation sont valides, l’utilisateur n’a pas besoin de se reconnecter.
