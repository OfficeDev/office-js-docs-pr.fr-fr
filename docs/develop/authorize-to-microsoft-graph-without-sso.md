---
title: Autoriser la connexion à Microsoft Graph sans authentification unique
description: Découvrir l'autorisation de connexion à Microsoft Graph sans authentification unique
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 828779a766c41088435ff5fdfa693e1d9939c710
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41949660"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>Autoriser la connexion à Microsoft Graph sans authentification unique

Votre complément peut recevoir une autorisation d’accès aux données Microsoft Graph en obtenant un jeton d’accès auprès d’Azure Active Directory (AAD). Utilisez le flux de codes d’autorisation ou le flux implicite, comme vous le feriez dans d'autres applications web, mais à une exception près : AAD n’autorise pas la page de connexion à s’ouvrir dans un iFrame. Lorsqu’un complément Office est exécuté sur *Office sur le web*, le volet Office est un IFrame. Cela signifie que vous devez ouvrir l’écran de connexion Azure Active Directory dans une boîte de dialogue ouverte avec l’API de boîte de dialogue Office. Cela a un effet sur votre utilisation des bibliothèques d’aide à l’authentification et l’autorisation. Pour plus d’informations, consultez l'[Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).

Pour plus d’informations sur la programmation de l’authentification avec Azure Active Directory, commencez par la [Vue d’ensemble de la plateforme d’identité Microsoft (v 2.0)](/azure/active-directory/develop/v2-overview) où se trouvent les didacticiels et guides de cet ensemble de documentation, ainsi que des liens vers des exemples pertinents. Une fois encore, vous devrez peut-être ajuster le code dans les exemples pour qu’il s’exécute dans la boîte de dialogue Office, afin de tenir compte du fait que la boîte de dialogue Office s’exécute dans un processus distinct de celui du volet Office.

Une fois que votre code obtient le jeton d’accès à Graph, soit il passe le jeton d’accès de la boîte de dialogue au volet Office, soit il stocke le jeton dans une base de données et signale au volet Office que le jeton est disponible. (Pour plus d’informations, voir [Authentification avec l’API de boîte de dialogue Office](auth-with-office-dialog-api.md).) Le code dans le volet Office demande les données de Graph et inclut le jeton dans ces demandes. Pour plus d’informations sur l’appel de Graph et des kits de développement Graph, voir la [documentation de Microsoft Graph](/graph/).

## <a name="recommended-libraries-and-samples"></a>Bibliothèques et exemples recommandés

Nous vous recommandons d’utiliser les bibliothèques suivantes lorsque vous accédez à Microsoft Graph sans utiliser l’authentification unique :

- Pour les compléments utilisant un élément côté serveur avec une infrastructure .NET basée sur le réseau, comme .NET Core ou ASP.NET, utilisez [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation).
- Pour les compléments utilisant un élément côté serveur NodeJS, utilisez [Passport Azure AD](https://github.com/AzureAD/passport-azure-ad).
- Pour les compléments utilisant le flux implicite, utilisez [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki).

Pour plus d’informations sur les bibliothèques recommandées avec la plateforme d’identité Microsoft (anciennement AAD v. 2.0), voir [Bibliothèques d’authentification de la plateforme d’identité Microsoft](/azure/active-directory/develop/reference-v2-libraries).

Les exemples suivants obtiennent les données Microsoft Graph d’un complément Office :

- [Complément Office Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Complément Outlook Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Complément Office Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
