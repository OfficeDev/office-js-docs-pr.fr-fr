---
title: Autoriser la connexion à Microsoft Graph sans authentification unique
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 1d696783003fc475f98d5d1d37f60348aacec5eb
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053311"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Autoriser la connexion à Microsoft Graph sans authentification unique

Vous pouvez obtenir l’autorisation d’accès aux données Microsoft Graph pour votre complément en obtenant un jeton d’accès auprès d’Azure Active Directory (AAD). Pour ce faire, utilisez le flux de codes d’autorisation ou le flux implicite, comme vous le feriez dans n’importe quelle autre application web, à une exception près : AAD n’autorise pas la page de connexion à s’ouvrir dans un iFrame. Lorsqu’un complément Office est exécuté sur *Office sur le Web*, le volet Office est un iFrame. Cela signifie que vous devez ouvrir l’écran de connexion AAD dans une boîte de dialogue ouverte avec l’API de boîte de dialogue Office. Cela a un effet sur votre utilisation des bibliothèques d’aide à l’authentification et l’autorisation. Pour plus d’informations, consultez [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

Pour plus d’informations sur la programmation de l’authentification avec AAD, commencez par la [vue d’ensemble de Microsoft Identity Platform (v 2.0)](/azure/active-directory/develop/v2-overview). Cet article présente de nombreux tutoriels et guides, ainsi que des liens vers des exemples pertinents. Une fois encore, rappelez-vous que vous devrez peut-être ajuster le code dans les exemples pour qu’il s’exécute dans la boîte de dialogue Office, afin de tenir compte du fait que la boîte de dialogue s’exécute dans un processus distinct de celui du volet Office.

Une fois que votre code a obtenu le jeton d’accès à Microsoft Graph, soit il passe le jeton d’accès de la boîte de dialogue au volet Office, soit il stocke le jeton dans une base de données et signale au volet Office que le jeton y est disponible. (Pour plus d’informations, voir [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).) Le code dans le volet Office demande les données de Microsoft Graph et inclut le jeton dans ces demandes. Pour plus d’informations sur l’appel de Microsoft Graph et des kits de développement pour Microsoft Graph, voir la [documentation de Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Bibliothèques et exemples recommandés

Nous vous recommandons d’utiliser les bibliothèques suivantes lorsque vous accédez à Microsoft Graph sans utiliser l’authentification unique :

- Pour les compléments utilisant un élément côté serveur avec une infrastructure .NET basée sur le réseau, comme .NET Core ou ASP.NET, utilisez [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Pour les compléments utilisant un élément côté serveur NodeJS, utilisez [Passport Azure AD](https://github.com/AzureAD/passport-azure-ad).
- Pour les compléments utilisant le flux implicite, utilisez [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Pour plus d’informations sur les bibliothèques recommandées avec la plateforme d’identité Microsoft (anciennement AAD v. 2.0), voir [Bibliothèques d’authentification de la plateforme d’identité Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Les exemples suivants obtiennent les données Microsoft Graph d’un complément Office :

- [Complément Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Complément Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)

