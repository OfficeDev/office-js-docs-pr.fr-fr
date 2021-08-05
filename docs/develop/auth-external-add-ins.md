---
title: Autoriser des services externes dans votre complément Office
description: Autorisation d’accès à des sources de données non-Microsoft comme Google, Facebook, LinkedIn, SalesForce et GitHub à l’aide d’OAuth 2.0, du code d’autorisation et des flux implicites.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 4927837caf96268421cfbc79ce3859b9dc4043f9
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773936"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autoriser des services externes dans votre complément Office

Les services en ligne populaires, notamment Microsoft 365, Google, Facebook, LinkedIn, SalesForce et GitHub, donnent aux développeurs l’accès à leurs comptes dans d’autres applications. Vous avez ainsi la possibilité d’inclure ces services dans votre complément Office.

> [!NOTE]
> Le reste de cet article concerne l’accès aux services non-Microsoft. Pour plus d’informations sur l’accès à Microsoft Graph (y compris Microsoft 365), consultez l’accès à Microsoft Graph avec [ssO](overview-authn-authz.md#access-to-microsoft-graph-with-sso) et l’accès à Microsoft Graph sans [SSO](overview-authn-authz.md#access-to-microsoft-graph-without-sso).

L’infrastructure standard dans le secteur permettant d’activer l’accès d’une application web à un service en ligne est appelée **OAuth 2.0**. En règle générale, vous n’avez pas besoin de connaître les détails du fonctionnement de l’infrastructure pour pouvoir l’utiliser dans votre complément. Ces détails sont simplifiés pour vous dans de nombreuses bibliothèques disponibles.

L’une des idées fondamentales d’OAuth est qu’une application peut être un [principal de sécurité](/windows/security/identity-protection/access-control/security-principals) en elle-même, de la même façon qu’un utilisateur ou un groupe, avec sa propre identité et son ensemble d’autorisations. Le plus souvent, quand l’utilisateur exécute une action dans le complément Office ayant besoin du service en ligne, le complément envoie une demande au service portant sur un ensemble spécifique d’autorisations pour le compte de l’utilisateur. Le service invite ensuite l’utilisateur à octroyer ces autorisations au complément. Une fois que les autorisations sont accordées, le service envoie un petit *jeton d’accès* codé au complément. Le complément peut utiliser le service en incluant le jeton dans toutes ses demandes aux API du service. Toutefois, le complément agit uniquement dans la limite des autorisations que l’utilisateur lui a accordées. En outre, le jeton expire après un certain délai.

Plusieurs modèles OAuth, appelés *flux* ou *types d’accès accordé*, sont conçus pour différents scénarios. Les deux modèles suivants sont les plus couramment implémentés.

- **Flux implicite** : la communication entre le complément et le service en ligne est mise en œuvre avec JavaScript côté client. Ce flux est couramment utilisé dans les applications à page unique (SPA).
- **Flux de code d’autorisation** : la communication est effectuée de *serveur à serveur* entre l’application web de votre complément et le service en ligne. Par conséquent, elle est mise en œuvre avec du code côté serveur.

L’objectif d’un flux OAuth est de sécuriser l’identité et l’autorisation de l’application. Dans le flux de code d’autorisation, une *clé secrète client* devant rester masquée vous est fournie. Les applications sans élément principal côté serveur, comme les applications monopages, ne permettent pas de protéger la clé secrète et nous vous recommandons d’utiliser le flux implicite dans ce type d’application.

Vous devez être familiarisé avec les avantages et inconvénients du flux implicite et du flux de code d’autorisation. Pour plus d’informations sur ces deux flux, reportez-vous à [Code d’autorisation](https://tools.ietf.org/html/rfc6749#section-1.3.1) et [Implicite](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> Vous avez aussi la possibilité de charger un service intermédiaire d’effectuer tout ce qui concerne les autorisations et de transmettre le jeton d’accès à votre complément. Pour plus d’informations sur ce scénario, consultez la rubrique **Services intermédiaires** plus loin dans cet article.

## <a name="use-the-implicit-flow-in-office-add-ins"></a>Utiliser le flux implicite dans les Office des modules

La meilleure façon de déterminer si un service en ligne prend en charge le flux implicite est de consulter la documentation.

Pour plus d’informations sur les bibliothèques prenant en charge le flux implicite, consultez la rubrique **bibliothèques** plus loin dans cet article.

## <a name="use-the-authorization-code-flow-in-office-add-ins"></a>Utiliser le flux de code d’autorisation dans Office des modules

De nombreuses bibliothèques sont disponibles pour l’implémentation du flux de code d’autorisation dans différentes langues et infrastructures. Pour plus d’informations sur ces bibliothèques, reportez-vous à la section **Bibliothèques** plus loin dans cet article.

## <a name="libraries"></a>Bibliothèques

Des bibliothèques sont disponibles dans de nombreuses langues et sur de nombreuses plateformes, aussi bien pour le flux implicite que pour le flux de code d’autorisation. Certaines sont destinées à un usage général, d’autres sont propres à des services en ligne bien spécifiques.

**Google** : cherchez « auth » ou le nom de votre langue sur [GitHub.com/Google](https://github.com/google). La plupart des référentiels pertinents sont nommés `google-auth-library-[name of language]`.

**Facebook** : cherchez « bibliothèque » ou « sdk » sur le site [Facebook pour les développeurs](https://developers.facebook.com).

**OAuth 2.0 général** : une page contenant des liens vers des bibliothèques pour plus d’une dizaine de langues est conservée par le groupe de travail OAuth de l’IETF sur une page relative au [code OAuth](https://oauth.net/code/). Notez que certaines de ces bibliothèques sont destinées à l’implémentation d’un service compatible OAuth. Les bibliothèques qui vous sont utiles en tant que développeur de compléments sont appelées bibliothèques *client* sur cette page car votre serveur web est un client du service compatible OAuth.

## <a name="middleman-services"></a>Services intermédiaires

Votre complément peut utiliser un service intermédiaire tel qu’[OAuth.io](https://oauth.io) ou [Auth0](https://auth0.com) pour gérer des autorisations. Un service intermédiaire peut fournir des jetons d’accès pour de nombreux services en ligne populaires ou simplifier la procédure de connexion aux réseaux sociaux pour votre complément, ou qui effectue ces deux opérations. Avec très peu de code, votre complément peut utiliser un script côté client ou du code côté serveur pour se connecter au service intermédiaire et envoyer les jetons requis à votre complément pour le service en ligne. L’ensemble du code de mise en œuvre des autorisations se trouve dans le service intermédiaire.

Nous vous recommandons que l’interface utilisateur de l’authentification/autorisation dans votre complément utilise nos boîte de dialogue API pour ouvrir une page de connexion. Voir[ Utilisation des API de dialogue dans un flux d’authentification](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow)pour plus d’informations. Lorsque vous ouvrez une boîte de dialogue Office de cette façon, la boîte de dialogue a une instance distincte et complètement nouvelle du moteur JavaScript à partir de l’instance de navigateur et dans la page parent (par exemple, volet Office du complément ou FunctionFile). Un jeton et toute autre information peut être converti(e) en chaîne, est transmis(e) au parent à l’aide d’une API appelée `messageParent`. La page parent peut ensuite utiliser le jeton pour passer des appels autorisés à la ressource. En raison de cette architecture, vous devez être vigilant de l’utilisation API fournis par un service intermédiaires. Le service fournit souvent une API définir dans lequel votre code crée un type d’objet de contexte qui obtient un jeton et utilise ce jeton afin de passer des appels conséquents à la ressource. Souvent le service a une méthode API unique qui effectue l’appel initiale *et* crée l’objet de contexte. Un objet comme suit ne peut pas être complètement mis sous forme de chaîne, il ne peut donc pas être transmis à partir de la boîte de dialogue Office à la page parent. En règle générale, le service intermédiaires fournit un ensemble de second API, du niveau inférieur d’abstraction, par exemple, une API REST. Cette seconde série comportera une API qui récupère un jeton à partir du service et autres API qui passe le jeton au service lorsque vous utilisez pour accéder à la ressource autorisée. Vous devez travailler avec une API à ce niveau inférieur d’abstraction afin que vous puissiez obtenir le jeton dans la boîte de dialogue Office, puis utiliser `messageParent` afin de le passer à la page parent.

## <a name="what-is-cors"></a>Que signifie l’acronyme CORS ?

CORS est l’acronyme de [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS) (partage des ressources d’origines croisées). Pour plus d’informations sur l’utilisation de CORS dans les compléments, reportez-vous à la rubrique relative à la [résolution des limites de stratégie d’origine identique dans les compléments Office](addressing-same-origin-policy-limitations.md).
