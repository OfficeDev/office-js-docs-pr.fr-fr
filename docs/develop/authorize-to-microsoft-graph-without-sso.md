---
title: Autoriser l’accès à Microsoft Graph à partir d’un Office de conférence
description: Découvrez comment autoriser microsoft à Graph à partir d’un Office de conférence.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8166b7a71767abd0456662dbe8573f59bb2c7e82
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743583"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>Autoriser l’accès à Microsoft Graph à partir d’un Office de conférence

Votre add-in peut obtenir l’autorisation d’accès aux données microsoft Graph en obtenant un jeton d’accès à Microsoft Graph à partir du Plateforme d'identités Microsoft. Utilisez le flux de code d’autorisation ou le flux implicite comme vous le feriez dans d’autres applications web, mais à une exception près : le Plateforme d'identités Microsoft n’autorise pas l’ouverture de sa page de signature dans un iFrame. Lorsqu’un complément Office est exécuté sur *Office sur le Web*, le volet des tâches est un IFrame. Cela signifie que vous devez ouvrir la page de connexion dans une boîte de dialogue à l’aide de l’API Office dialogue. Cela a un effet sur votre utilisation des bibliothèques d’aide à l’authentification et l’autorisation. Pour plus d’informations, consultez l'[Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

> [!NOTE]
> Si vous implémentez l’oD SSO et prévoyez d’accéder à Microsoft Graph, consultez Autoriser [l’accès à Microsoft Graph avec sso](authorize-to-microsoft-graph.md).

Pour plus d’informations sur la programmation de l’authentification à l’Plateforme d'identités Microsoft, voir [Plateforme d'identités Microsoft documentation](/azure/active-directory/develop). Vous trouverez des didacticiels et des guides dans cet ensemble de documentation, ainsi que des liens vers des exemples pertinents. Une fois encore, vous devrez peut-être ajuster le code des exemples à exécuter dans la boîte de dialogue Office pour prendre en compte la boîte de dialogue Office qui s’exécute dans un processus distinct du volet Des tâches.

Une fois que votre code a obtenu le jeton d’accès à Microsoft Graph, il transmet le jeton d’accès de la boîte de dialogue au volet Des tâches, ou il stocke le jeton dans une base de données et signale au volet Des tâches que le jeton est disponible. (Pour plus [d’informations, voir l’authentification Office’API de boîte](auth-with-office-dialog-api.md) de dialogue.) Le code du volet Des tâches demande des données à Microsoft Graph et inclut le jeton dans ces demandes. Pour plus d’informations sur l’appel de Microsoft Graph et des SDK Microsoft Graph, consultez la [documentation de Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Bibliothèques et exemples recommandés

Nous vous recommandons d’utiliser les bibliothèques suivantes lors de l’accès à Microsoft Graph.

- Pour les compléments utilisant un élément côté serveur avec une infrastructure .NET basée sur le réseau, comme .NET Core ou ASP.NET, utilisez [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Pour les compléments utilisant un élément côté serveur NodeJS, utilisez [Passport Azure AD](https://github.com/AzureAD/passport-azure-ad).
- Pour les compléments utilisant le flux implicite, utilisez [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Pour plus d’informations sur les bibliothèques recommandées avec la plateforme d’identité Microsoft (anciennement AAD v. 2.0), voir [Bibliothèques d’authentification de la plateforme d’identité Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Les exemples suivants obtiennent des données microsoft Graph à partir d’un Office de gestion.

- [Complément Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Complément Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Complément Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
