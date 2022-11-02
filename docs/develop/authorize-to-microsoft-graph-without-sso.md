---
title: Autoriser sur Microsoft Graph à partir d’un complément Office
description: Découvrez comment autoriser Microsoft Graph à partir d’un complément Office.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 37dd4be3acb92dc7884972de923d94936fa870f4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810168"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>Autoriser sur Microsoft Graph à partir d’un complément Office

Votre complément peut obtenir l’autorisation pour les données Microsoft Graph en obtenant un jeton d’accès à Microsoft Graph à partir du Plateforme d'identités Microsoft. Utilisez le flux de code d’autorisation ou le flux implicite comme vous le feriez dans d’autres applications web, mais à une exception près : le Plateforme d'identités Microsoft n’autorise pas l’ouverture de sa page de connexion dans un iframe. Lorsqu’un complément Office s’exécute dans *Office sur le Web*, le volet Office est un iframe. Cela signifie que vous devez ouvrir la page de connexion dans une boîte de dialogue à l’aide de l’API de boîte de dialogue Office. Cela a un effet sur votre utilisation des bibliothèques d’aide à l’authentification et l’autorisation. Pour plus d’informations, consultez l'[Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

> [!NOTE]
> Si vous implémentez l’authentification unique et envisagez d’accéder à Microsoft Graph, consultez [Autoriser Microsoft Graph avec l’authentification unique](authorize-to-microsoft-graph.md).

Pour plus d’informations sur la programmation de l’authentification à l’aide du Plateforme d'identités Microsoft, consultez [Plateforme d'identités Microsoft documentation](/azure/active-directory/develop). Vous trouverez des tutoriels et des guides dans cet ensemble de documentation, ainsi que des liens vers des exemples pertinents. Une fois de plus, vous devrez peut-être ajuster le code dans les exemples à exécuter dans la boîte de dialogue Office pour tenir compte de la boîte de dialogue Office qui s’exécute dans un processus distinct du volet Office.

Une fois que votre code a obtenu le jeton d’accès à Microsoft Graph, il transmet le jeton d’accès de la boîte de dialogue au volet Office, ou il stocke le jeton dans une base de données et signale au volet Office que le jeton est disponible. (Pour plus d’informations, consultez [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md) .) Le code dans le volet Office demande des données à Microsoft Graph et inclut le jeton dans ces demandes. Pour plus d’informations sur l’appel de Microsoft Graph et des Kits de développement logiciel (SDK) Microsoft Graph, consultez [la documentation Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Bibliothèques et exemples recommandés

Nous vous recommandons d’utiliser les bibliothèques suivantes lors de l’accès à Microsoft Graph.

- Pour les compléments utilisant un élément côté serveur avec une infrastructure .NET basée sur le réseau, comme .NET Core ou ASP.NET, utilisez [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Pour les compléments utilisant un élément côté serveur NodeJS, utilisez [Passport Azure AD](https://github.com/AzureAD/passport-azure-ad).
- Pour les compléments utilisant le flux implicite, utilisez [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Pour plus d’informations sur les bibliothèques recommandées avec la plateforme d’identité Microsoft (anciennement AAD v. 2.0), voir [Bibliothèques d’authentification de la plateforme d’identité Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Les exemples suivants obtiennent des données Microsoft Graph à partir d’un complément Office.

- [Complément Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Complément Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Complément Office Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
