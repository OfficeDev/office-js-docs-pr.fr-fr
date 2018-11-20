---
title: Autoriser des services externes dans votre complément Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 6cdf07886ba883a7dfe935b59c918948c2b45afa
ms.sourcegitcommit: 86724e980f720ed05359c9525948cb60b6f10128
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/09/2018
ms.locfileid: "26237478"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autoriser des services externes dans votre complément Office

Les services en ligne populaires, y compris Office 365, Google, Facebook, LinkedIn, SalesForce et GitHub, permettent aux développeurs d’accorder aux utilisateurs l’accès à leurs comptes dans d’autres applications. Vous avez ainsi la possibilité d’inclure ces services dans votre complément Office.

L’infrastructure standard dans le secteur permettant d’activer l’accès d’une application web à un service en ligne est appelée **OAuth 2.0**. En règle générale, vous n’avez pas besoin de connaître les détails du fonctionnement de l’infrastructure pour pouvoir l’utiliser dans votre complément. Ces détails sont simplifiés pour vous dans de nombreuses bibliothèques disponibles.

L’un des fondements d’OAuth est qu’une application peut être un principal de sécurité en elle-même, de la même façon qu’un utilisateur ou un groupe, avec sa propre identité et son ensemble d’autorisations. Le plus souvent, quand l’utilisateur exécute une action dans le complément Office ayant besoin du service en ligne, le complément envoie une demande au service portant sur un ensemble spécifique d’autorisations pour le compte de l’utilisateur. Le service invite ensuite l’utilisateur à octroyer ces autorisations au complément. Une fois que les autorisations sont accordées, le service envoie un petit *jeton d’accès* codé au complément. Le complément peut utiliser le service en incluant le jeton dans toutes ses demandes aux API du service. Toutefois, le complément agit uniquement dans la limite des autorisations que l’utilisateur lui a accordées. En outre, le jeton expire après un certain délai.

Plusieurs modèles OAuth, appelés *flux* ou *types d’accès accordé*, sont conçus pour différents scénarios. Les deux modèles suivants sont les plus couramment implémentés :

- **Flux implicite** : la communication entre le complément et le service en ligne est mise en œuvre avec JavaScript côté client.
- **Flux de code d’autorisation** : la communication est effectuée de *serveur à serveur* entre l’application web de votre complément et le service en ligne. Par conséquent, elle est mise en œuvre avec du code côté serveur.

L’objectif d’un flux OAuth est de sécuriser l’identité et l’autorisation de l’application. Dans le flux de code d’autorisation, une *clé secrète client* devant rester masquée vous est fournie. Comme une application monopage (SPA) ne permet pas de protéger la clé secrète, nous vous recommandons d’utiliser le flux implicite dans ce type d’application.

Vous devez être familiarisé avec les avantages et inconvénients du flux implicite et du flux de code d’autorisation. Pour plus d’informations sur ces deux flux, reportez-vous à [Code d’autorisation](https://tools.ietf.org/html/rfc6749#section-1.3.1) et [Implicite](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> Vous avez aussi la possibilité de charger un service intermédiaire d’effectuer tout ce qui concerne les autorisations et de transmettre le jeton d’accès à votre complément. Pour plus d’informations sur ce scénario, consultez la rubrique **Services intermédiaires** plus loin dans cet article.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Utilisation du flux implicite dans des compléments Office
La meilleure façon de déterminer si un service en ligne prend en charge le flux implicite est de consulter la documentation. Pour les services qui prennent en charge le flux implicite, vous pouvez charger la bibliothèque JavaScript **Office-js-helpers** d’effectuer à votre place toutes les tâches détaillées :

- [Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

Pour plus d’informations sur les autres bibliothèques prenant en charge le flux implicite, consultez la rubrique **bibliothèques** plus loin dans cet article.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Utilisation du flux de code d’autorisation dans les compléments Office

De nombreuses bibliothèques sont disponibles pour l’implémentation du flux de code d’autorisation dans différentes langues et infrastructures. Pour plus d’informations sur ces bibliothèques, reportez-vous à la section **Bibliothèques** plus loin dans cet article.

Les aperçus suivants fournissent des exemples de compléments qui implémentent le flux de code d’autorisation :

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

### <a name="relayproxy-functions"></a>Fonctions de relais/proxy

Vous pouvez utiliser le flux de code d’autorisation même avec une application web sans serveur en stockant les valeurs d’**identifiant client** et de **clé secrète client** dans une fonction simple, hébergée dans un service tel qu’[Azure Functions](https://azure.microsoft.com/services/functions) ou [Amazon Lambda](https://aws.amazon.com/lambda). La fonction remplace un code donné par un **jeton d’accès** et le transmet au client. La sécurité de cette approche dépend de la surveillance de l’accès à la fonction.

Pour utiliser cette technique, votre complément ouvre une interface utilisateur/un menu contextuel pour afficher l’écran de connexion au service en ligne (Google, Facebook, etc.). Lorsque l’utilisateur est connecté et accorde l’autorisation au complément d’accéder à ses ressources dans le service en ligne, le complément reçoit un code qui peut être envoyé à la fonction en ligne. Les services décrits dans la section **Services intermédiaires** plus loin dans cet article utilisent un flux semblable à celui-ci.

## <a name="libraries"></a>Bibliothèques

Des bibliothèques sont disponibles dans de nombreuses langues et sur de nombreuses plateformes, aussi bien pour le flux implicite que pour le flux de code d’autorisation. Certaines sont destinées à un usage général, d’autres sont propres à des services en ligne bien spécifiques.

**Office 365 et autres services utilisant Azure Active Directory en tant que fournisseur d’autorisation** : [bibliothèques d’authentification Azure Active Directory](https://azure.microsoft.com/documentation/articles/active-directory-authentication-libraries/). Un aperçu est également disponible pour la [bibliothèque d’authentification Microsoft](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google** : cherchez « auth » ou le nom de votre langue sur [GitHub.com/Google](https://github.com/google). La plupart des référentiels pertinents sont nommés `google-auth-library-[name of language]`.

**Facebook** : cherchez « bibliothèque » ou « sdk » sur le site [Facebook pour les développeurs](https://developers.facebook.com).

**OAuth 2.0 général** : une page contenant des liens vers des bibliothèques pour plus d’une dizaine de langues est conservée par le groupe de travail OAuth de l’IETF sur une page relative au [code OAuth](http://oauth.net/code/). Notez que certaines de ces bibliothèques sont destinées à l’implémentation d’un service compatible OAuth. Les bibliothèques qui vous sont utiles en tant que développeur de compléments sont appelées bibliothèques *client* sur cette page car votre serveur web est un client du service compatible OAuth.

## <a name="middleman-services"></a>Services intermédiaires

Votre complément peut utiliser un service intermédiaire tel qu’OAuth.io ou Auth0 pour gérer des autorisations. Un service intermédiaire peut fournir des jetons d’accès pour de nombreux services en ligne populaires ou simplifier la procédure de connexion aux réseaux sociaux pour votre complément, ou qui effectue ces deux opérations. Avec très peu de code, votre complément peut utiliser un script côté client ou du code côté serveur pour se connecter au service intermédiaire et envoyer les jetons requis à votre complément pour le service en ligne. L’ensemble du code de mise en œuvre des autorisations se trouve dans le service intermédiaire.

Pour obtenir des exemples de compléments qui utilisent un service intermédiaire d’autorisation, voir les exemples suivants :

- [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0) utilise Auth0 pour activer la connexion aux réseaux sociaux avec les comptes Facebook, Google et Microsoft.

- [Office-Add-in-OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io) utilise OAuth.io pour obtenir des jetons d’accès à partir de Facebook et Google.

## <a name="what-is-cors"></a>Que signifie l’acronyme CORS ?

CORS est l’acronyme de [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS) (partage des ressources d’origines croisées). Pour plus d’informations sur l’utilisation de CORS dans les compléments, reportez-vous à la rubrique relative à la [résolution des limites de stratégie d’origine identique dans les compléments Office](addressing-same-origin-policy-limitations.md).
