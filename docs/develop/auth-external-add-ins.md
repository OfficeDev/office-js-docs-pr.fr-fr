---
title: Autoriser des services externes dans votre complément Office
description: Autorisation d’accès à des sources de données non-Microsoft comme Google, Facebook, LinkedIn, SalesForce et GitHub à l’aide d’OAuth 2.0, du code d’autorisation et des flux implicites.
ms.date: 08/07/2019
localization_priority: Normal
ms.openlocfilehash: fd180e11106e7e1e2f20f539746535c4310ad81e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093741"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autoriser des services externes dans votre complément Office

Popular online services, including Microsoft 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in.

> [!NOTE]
> Le reste de cet article concerne l’accès aux services non-Microsoft. Pour plus d’informations sur l’accès à Microsoft Graph (y compris Microsoft 365), voir [Access to Microsoft Graph with SSO](overview-authn-authz.md#access-to-microsoft-graph-with-sso) and [Access to Microsoft Graph with SSO](overview-authn-authz.md#access-to-microsoft-graph-without-sso).

The industry standard framework for enabling web application access to an online service is **OAuth 2.0**. In most situations, you don't need to know the details of how the framework works to use it in your add-in. Many libraries are available that simplify the details for you.

L’une des idées fondamentales d’OAuth est qu’une application peut être un [principal de sécurité](/windows/security/identity-protection/access-control/security-principals) en elle-même, de la même façon qu’un utilisateur ou un groupe, avec sa propre identité et son ensemble d’autorisations. Le plus souvent, quand l’utilisateur exécute une action dans le complément Office ayant besoin du service en ligne, le complément envoie une demande au service portant sur un ensemble spécifique d’autorisations pour le compte de l’utilisateur. Le service invite ensuite l’utilisateur à octroyer ces autorisations au complément. Une fois que les autorisations sont accordées, le service envoie un petit *jeton d’accès* codé au complément. Le complément peut utiliser le service en incluant le jeton dans toutes ses demandes aux API du service. Toutefois, le complément agit uniquement dans la limite des autorisations que l’utilisateur lui a accordées. En outre, le jeton expire après un certain délai.

Plusieurs modèles OAuth, appelés *flux* ou *types d’accès accordé*, sont conçus pour différents scénarios. Les deux modèles suivants sont les plus couramment implémentés :

- **Flux implicite** : la communication entre le complément et le service en ligne est mise en œuvre avec JavaScript côté client. Ce flux est couramment utilisé dans les applications à page unique (SPA).
- **Flux de code d’autorisation** : la communication est effectuée de *serveur à serveur* entre l’application web de votre complément et le service en ligne. Par conséquent, elle est mise en œuvre avec du code côté serveur.

L’objectif d’un flux OAuth est de sécuriser l’identité et l’autorisation de l’application. Dans le flux de code d’autorisation, une *clé secrète client* devant rester masquée vous est fournie. Les applications sans élément principal côté serveur, comme les applications monopages, ne permettent pas de protéger la clé secrète et nous vous recommandons d’utiliser le flux implicite dans ce type d’application.

Vous devez être familiarisé avec les avantages et inconvénients du flux implicite et du flux de code d’autorisation. Pour plus d’informations sur ces deux flux, reportez-vous à [Code d’autorisation](https://tools.ietf.org/html/rfc6749#section-1.3.1) et [Implicite](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> You also have the option of using a middleman service to perform authorization and pass the access token to your add-in. For details about this scenario, see the **Middleman services** section later in this article.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Utilisation du flux implicite dans des compléments Office

La meilleure façon de déterminer si un service en ligne prend en charge le flux implicite est de consulter la documentation.

Pour plus d’informations sur les bibliothèques prenant en charge le flux implicite, consultez la rubrique **bibliothèques** plus loin dans cet article.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Utilisation du flux de code d’autorisation dans les compléments Office

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For more information about some of these libraries, see the **Libraries** section later in this article.

## <a name="libraries"></a>Bibliothèques

Libraries are available for many languages and platforms, for both the Implicit flow and the Authorization Code flow. Some libraries are general purpose, while others are for specific online services.

**Google**: Search [GitHub.com/Google](https://github.com/google) for "auth" or the name of your language. Most of the relevant repos are named `google-auth-library-[name of language]`.

**Facebook** : cherchez « bibliothèque » ou « sdk » sur le site [Facebook pour les développeurs](https://developers.facebook.com).

**General OAuth 2.0**: A page of links to libraries for over a dozen languages is maintained by the IETF OAuth Working Group at: [OAuth Code](https://oauth.net/code/). Note that some of these libraries are for implementing an OAuth compliant service. The libraries of interest to you as a an add-in developer are called *client* libraries on this page because your web server is a client of the OAuth compliant service.

## <a name="middleman-services"></a>Services intermédiaires

Your add-in can use a middleman service such as [OAuth.io](https://oauth.io) or [Auth0](https://auth0.com) to perform authorization. A middleman service may either provide access tokens for popular online services or simplify the process of enabling social login for your add-in, or both. With very little code, your add-in can use either client-side script or server-side code to connect to the middleman service and it will send your add-in any required tokens for the online service. All of the authorization implementation code is in the middleman service. 

Nous vous recommandons que l’interface utilisateur de l’authentification/autorisation dans votre complément utilise nos boîte de dialogue API pour ouvrir une page de connexion. Voir[ Utilisation des API de dialogue dans un flux d’authentification](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow)pour plus d’informations. Lorsque vous ouvrez une boîte de dialogue Office de cette façon, la boîte de dialogue a une instance distincte et complètement nouvelle du moteur JavaScript à partir de l’instance de navigateur et dans la page parent (par exemple, volet Office du complément ou FunctionFile). Un jeton et toute autre information peut être converti(e) en chaîne, est transmis(e) au parent à l’aide d’une API appelée `messageParent`. La page parent peut ensuite utiliser le jeton pour passer des appels autorisés à la ressource. En raison de cette architecture, vous devez être vigilant de l’utilisation API fournis par un service intermédiaires. Le service fournit souvent une API définir dans lequel votre code crée un type d’objet de contexte qui obtient un jeton et utilise ce jeton afin de passer des appels conséquents à la ressource. Souvent le service a une méthode API unique qui effectue l’appel initiale *et* crée l’objet de contexte. Un objet comme suit ne peut pas être complètement mis sous forme de chaîne, il ne peut donc pas être transmis à partir de la boîte de dialogue Office à la page parent. En règle générale, le service intermédiaires fournit un ensemble de second API, du niveau inférieur d’abstraction, par exemple, une API REST. Cette seconde série comportera une API qui récupère un jeton à partir du service et autres API qui passe le jeton au service lorsque vous utilisez pour accéder à la ressource autorisée. Vous devez travailler avec une API à ce niveau inférieur d’abstraction afin que vous puissiez obtenir le jeton dans la boîte de dialogue Office, puis utiliser `messageParent` afin de le passer à la page parent. 

## <a name="what-is-cors"></a>Que signifie l’acronyme CORS ?

CORS stands for [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). For information about how to use CORS inside add-ins, see [Addressing same-origin policy limitations in Office Add-ins](addressing-same-origin-policy-limitations.md).
