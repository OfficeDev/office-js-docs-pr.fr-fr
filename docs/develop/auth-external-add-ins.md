---
title: Autoriser des services externes dans votre compl?ment Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 34e8119d4ecf6432cde7f06552584d164b8c9b8e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autoriser des services externes dans votre compl?ment Office

Les services en ligne populaires, y compris Office 365, Google, Facebook, LinkedIn, SalesForce et GitHub, permettent aux d?veloppeurs d?accorder aux utilisateurs l?acc?s ? leurs comptes dans d?autres applications. Vous avez ainsi la possibilit? d?inclure ces services dans votre compl?ment Office.

L?infrastructure standard dans le secteur permettant d?activer l?acc?s d?une application web ? un service en ligne est appel?e **OAuth 2.0**. En r?gle g?n?rale, vous n?avez pas besoin de conna?tre les d?tails du fonctionnement de l?infrastructure pour pouvoir l?utiliser dans votre compl?ment. Ces d?tails sont simplifi?s pour vous dans de nombreuses biblioth?ques disponibles.

L?un des fondements d?OAuth est qu?une application peut ?tre un principal de s?curit? en elle-m?me, de la m?me fa?on qu?un utilisateur ou un groupe, avec sa propre identit? et son ensemble d?autorisations. Dans les sc?narios les plus courants, lorsque l?utilisateur ex?cute une action dans le compl?ment Office ayant besoin du service en ligne, le compl?ment envoie une demande au service portant sur un ensemble sp?cifique d?autorisations pour le compte de l?utilisateur. Le service invite ensuite l?utilisateur ? octroyer ces autorisations au compl?ment. Une fois que les autorisations sont accord?es, le service envoie un petit *jeton d?acc?s* cod? au compl?ment. Le compl?ment peut utiliser le service en incluant le jeton dans toutes ses demandes aux API du service. Toutefois, le compl?ment agit uniquement dans la limite des autorisations que l?utilisateur lui a accord?es. En outre, le jeton expire apr?s un certain d?lai.

Plusieurs mod?les OAuth, appel?s *flux* ou *types d?acc?s accord?*, sont con?us pour diff?rents sc?narios. Les deux mod?les suivants sont les plus couramment impl?ment?s :

- **Flux implicite** : la communication entre le compl?ment et le service en ligne est mise en ?uvre avec JavaScript c?t? client.
- **Flux de code d?autorisation** : la communication est effectu?e de *serveur ? serveur* entre l?application web de votre compl?ment et le service en ligne. Par cons?quent, elle est mise en ?uvre avec du code c?t? serveur.

L?objectif d?un flux OAuth est de s?curiser l?identit? et l?autorisation de l?application. Dans le flux de code d?autorisation, une *cl? secr?te client* devant rester masqu?e vous est fournie. Comme une application monopage (SPA) ne permet pas de prot?ger la cl? secr?te, nous vous recommandons d?utiliser le flux implicite dans ce type d?application.

Vous devez ?tre familiaris? avec les avantages et inconv?nients du flux implicite et du flux de code d?autorisation. Pour plus d?informations sur ces deux flux, reportez-vous ? [Code d?autorisation](https://tools.ietf.org/html/rfc6749#section-1.3.1) et [Implicite](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> Vous avez aussi la possibilit? de charger un service interm?diaire d?effectuer tout ce qui concerne les autorisations et de transmettre le jeton d?acc?s ? votre compl?ment. Pour plus d?informations sur ce sc?nario, consultez la rubrique **Services interm?diaires** plus loin dans cet article.

## <a name="authorization-to-microsoft-graph"></a>Autorisation d?acc?s ? Microsoft Graph

Si le service externe est accessible via Microsoft Graph, par exemple Office 365 ou OneDrive, vous pouvez fournir ? vos utilisateurs la meilleure exp?rience possible tout en profitant vous-m?me d?une exp?rience de d?veloppement la plus simple possible, en utilisant le syst?me d?authentification unique d?crit sur la page [Autoriser Microsoft Graph dans vos compl?ments Office](authorize-to-microsoft-graph.md) et ses articles connexes. Les techniques d?crites dans cet article trouvent leur meilleur usage dans des services externes qui ne sont pas accessibles avec Microsoft Graph. Toutefois, elles *peuvent* ?tre utilis?es pour acc?der ? Microsoft Graph, et vous pouvez pr?f?rer leurs avantages ? ceux de l?authentification unique. Par exemple, le syst?me d?authentification unique requiert un code c?t? serveur, et ne peut donc pas ?tre utilis? dans une application web monopage. En outre, le syst?me de l?authentification unique n?est pas encore pris en charge sur toutes les plateformes.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Utilisation du flux implicite dans des compl?ments Office
La meilleure fa?on de d?terminer si un service en ligne prend en charge le flux implicite est de consulter la documentation. Pour les services qui prennent en charge le flux implicite, vous pouvez charger la biblioth?que JavaScript **Office-js-helpers** d?effectuer ? votre place toutes les t?ches d?taill?es :

- [Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

Pour plus d?informations sur les autres biblioth?ques prenant en charge le flux implicite, consultez la rubrique **biblioth?ques** plus loin dans cet article.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Utilisation du flux de code d?autorisation dans les compl?ments Office

De nombreuses biblioth?ques sont disponibles pour l?impl?mentation du flux de code d?autorisation dans diff?rentes langues et infrastructures. Pour plus d?informations sur ces biblioth?ques, reportez-vous ? la section **Biblioth?ques** plus loin dans cet article.

Les aper?us suivants fournissent des exemples de compl?ments qui impl?mentent le flux de code d?autorisation :

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

### <a name="relayproxy-functions"></a>Fonctions de relais/proxy

Vous pouvez utiliser le flux de code d?autorisation m?me avec une application web sans serveur en stockant les valeurs d?**identifiant client** et de **cl? secr?te client** dans une fonction simple, h?berg?e dans un service tel qu?[Azure Functions](https://azure.microsoft.com/en-us/services/functions) ou [Amazon Lambda](https://aws.amazon.com/lambda). La fonction remplace un code donn? par un **jeton d?acc?s** et le transmet au client. La s?curit? de cette approche d?pend de la surveillance de l?acc?s ? la fonction.

Pour utiliser cette technique, votre compl?ment ouvre une interface utilisateur/un menu contextuel pour afficher l??cran de connexion au service en ligne (Google, Facebook, etc.). Lorsque l?utilisateur est connect? et accorde l?autorisation au compl?ment d?acc?der ? ses ressources dans le service en ligne, le compl?ment re?oit un code qui peut ?tre envoy? ? la fonction en ligne. Les services d?crits dans la section **Services interm?diaires** plus loin dans cet article utilisent un flux semblable ? celui-ci.

## <a name="libraries"></a>Biblioth?ques

Des biblioth?ques sont disponibles dans de nombreuses langues et sur de nombreuses plateformes, aussi bien pour le flux implicite que pour le flux de code d?autorisation. Certaines sont destin?es ? un usage g?n?ral, d?autres sont propres ? des services en ligne bien sp?cifiques.

**Office 365 et autres services utilisant Azure Active Directory en tant que fournisseur d?autorisation** : [biblioth?ques d?authentification Azure Active Directory](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/). Un aper?u est ?galement disponible pour la [biblioth?que d?authentification Microsoft](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google** : cherchez ? auth ? ou le nom de votre langue sur [GitHub.com/Google](https://github.com/google). La plupart des r?f?rentiels pertinents sont nomm?s `google-auth-library-[name of language]`.

**Facebook** : cherchez ? biblioth?que ? ou ? sdk ? sur le site [Facebook pour les d?veloppeurs](https://developers.facebook.com).

**OAuth 2.0 g?n?ral** : une page contenant des liens vers des biblioth?ques pour plus d?une dizaine de langues est conserv?e par le groupe de travail OAuth de l?IETF sur une page relative au [code OAuth](http://oauth.net/code/). Notez que certaines de ces biblioth?ques sont destin?es ? l?impl?mentation d?un service compatible OAuth. Les biblioth?ques qui vous sont utiles en tant que d?veloppeur de compl?ments sont appel?es biblioth?ques *client* sur cette page car votre serveur web est un client du service compatible OAuth.

## <a name="middleman-services"></a>Services interm?diaires

Votre compl?ment peut utiliser un service interm?diaire tel qu?OAuth.io ou Auth0 pour g?rer des autorisations. Un service interm?diaire peut fournir des jetons d?acc?s pour de nombreux services en ligne populaires ou simplifier la proc?dure de connexion aux r?seaux sociaux pour votre compl?ment, ou qui effectue ces deux op?rations. Avec tr?s peu de code, votre compl?ment peut utiliser un script c?t? client ou du code c?t? serveur pour se connecter au service interm?diaire et envoyer les jetons requis ? votre compl?ment pour le service en ligne. L?ensemble du code de mise en ?uvre des autorisations se trouve dans le service interm?diaire.

Pour obtenir des exemples de compl?ments qui utilisent un service interm?diaire d?autorisation, voir les exemples suivants :

- [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0) utilise Auth0 pour activer la connexion aux r?seaux sociaux avec les comptes Facebook, Google et Microsoft.

- [Office-Add-in-OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io) utilise OAuth.io pour obtenir des jetons d?acc?s ? partir de Facebook et Google.

## <a name="what-is-cors"></a>Que signifie l?acronyme CORS ?

CORS est l?acronyme de [Cross Origin Resource Sharing](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS) (partage des ressources d?origines crois?es). Pour plus d?informations sur l?utilisation de CORS dans les compl?ments, reportez-vous ? la rubrique relative ? la [r?solution des limites de strat?gie d?origine identique dans les compl?ments Office](addressing-same-origin-policy-limitations.md).
