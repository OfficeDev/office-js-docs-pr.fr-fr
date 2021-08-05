---
title: Autoriser la connexion à Microsoft Graph sans authentification unique
description: Découvrir l'autorisation de connexion à Microsoft Graph sans authentification unique
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: c16af84bf63ead9acb81cf92be0a14ab92a6def3
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773929"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Autoriser la connexion à Microsoft Graph sans authentification unique

Votre add-in peut obtenir l’autorisation d’accès aux données microsoft Graph en obtenant un jeton d’accès à Microsoft Graph auprès de Azure Active Directory (Azure AD). Utilisez le flux de code d’autorisation ou le flux implicite comme vous le feriez dans d’autres applications web, mais à une exception près : Azure AD n’autorise pas l’ouverture de sa page de signature dans un iFrame. Lorsqu’un complément Office est exécuté sur *Office sur le Web*, le volet des tâches est un IFrame. Cela signifie que vous devez ouvrir l’écran de connexion Azure AD dans une boîte de dialogue ouverte avec l’API Office dialogue. Cela a un effet sur votre utilisation des bibliothèques d’aide à l’authentification et l’autorisation. Pour plus d’informations, consultez l'[Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

Pour plus d’informations sur la programmation de l’authentification avec Azure AD, commencez par la vue d’ensemble de [Plateforme d’identités Microsoft (v2.0),](/azure/active-directory/develop/v2-overview)où vous trouverez des didacticiels et des guides dans cet ensemble de documentation, ainsi que des liens vers des exemples pertinents. Une fois encore, vous devrez peut-être ajuster le code dans les exemples pour qu’il s’exécute dans la boîte de dialogue Office, afin de tenir compte du fait que la boîte de dialogue Office s’exécute dans un processus distinct de celui du volet Office.

Une fois que votre code a obtenu le jeton d’accès à Microsoft Graph, il transmet le jeton d’accès de la boîte de dialogue au volet Des tâches, ou il stocke le jeton dans une base de données et signale au volet Des tâches que le jeton est disponible. (Pour plus [d’informations, voir Authentification Office boîte](auth-with-office-dialog-api.md) de dialogue.) Le code du volet Des tâches demande des données à Microsoft Graph et inclut le jeton dans ces demandes. Pour plus d’informations sur l’appel de Microsoft Graph et des SDK Microsoft Graph, voir [la documentation microsoft Graph.](/graph/)

## <a name="recommended-libraries-and-samples"></a>Bibliothèques et exemples recommandés

Nous vous recommandons d’utiliser les bibliothèques suivantes lorsque vous accédez à Microsoft Graph sans utiliser l' sso.

- Pour les compléments utilisant un élément côté serveur avec une infrastructure .NET basée sur le réseau, comme .NET Core ou ASP.NET, utilisez [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Pour les compléments utilisant un élément côté serveur NodeJS, utilisez [Passport Azure AD](https://github.com/AzureAD/passport-azure-ad).
- Pour les compléments utilisant le flux implicite, utilisez [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Pour plus d’informations sur les bibliothèques recommandées avec la plateforme d’identité Microsoft (anciennement AAD v. 2.0), voir [Bibliothèques d’authentification de la plateforme d’identité Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Les exemples suivants obtiennent des données microsoft Graph à partir d’un Office de gestion.

- [Complément Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Complément Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Complément Office Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
