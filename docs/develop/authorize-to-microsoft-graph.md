---
title: Autoriser Microsoft Graph dans votre complément Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 6d0b6f2002b71c4680b72d2f40492fff1abf15e2
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505858"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Autorisez Microsoft Graph dans votre complément Office (préversion)

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365). Le meilleur moyen pour un complément Office d'obtenir un accès autorisé à [Microsoft Graph](https://developer.microsoft.com/graph/docs) est d'utiliser les informations d'identification de connexion Office de l'utilisateur. Cela leur permet d’accéder à leurs données Microsoft Graph sans avoir besoin de se connecter une seconde fois. 

> [!NOTE]
> L’API d’authentification unique est actuellement prise en charge en mode préversion pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur la prise en charge de l’API d’authentification unique reportez-vous à [Ensembles de conditions requises d’IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Architecture de complément pour SSO et Microsoft Graph

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](https://docs.microsoft.com/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du complément contient un balisage qui spécifie comment le complément est enregistré dans le point de terminaison Azure Active Directory (Azure AD) v2.0 et il indique les autorisations à Microsoft Graph dont le complément a besoin.

### <a name="how-it-works-at-runtime"></a>Fonctionnement à l’exécution

Le diagramme suivant montre le fonctionnement du processus de connexion et d'accès à Microsoft Graph.

![Diagramme illustrant le processus d’authentification unique](../images/sso-access-to-microsoft-graph.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Cette option indique à l’application hôte Office qu'il faut obtenir un jeton d’accès pour le complément (ci-après, il s’agit du **jeton d’accès des données d’amorçage** , car il est remplacé par un deuxième jeton plus loin dans le processus. Pour obtenir un exemple d’un jeton d’accès d’amorçage décodé, voir [Exemple de jeton d’accès](sso-in-office-add-ins.md#example-access-token).)
1. Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.
1. Si c’est la première fois que l’utilisateur utilise votre complément, il est invité à donner son consentement.
1. L’application hôte Office demande le **jeton d'accès de démarrage** au point de terminaison Azure AD version 2.0 pour l’utilisateur actuel.
1. Azure AD envoie le jeton de démarrage à l’application hôte Office.
1. L’application hôte Office envoie le **jeton d'accès de démarrage** au complément en tant que partie de l’objet résultat renvoyé par l’appel `getAccessTokenAsync`.
1. Un code JavaScript dans le complément effectue une requête HTTP vers une API web hébergée sur le même domaine complet que le complément et inclut le **jeton d'accès de démarrage** comme preuve d’autorisation.  
1. Le code côté serveur valide le **jeton d'accès de démarrage**entrant.
1. Le code côté serveur utilise le flux « au nom de » (défini par l'[Échange de jetons OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) et l'[application démon ou serveur pour un scénario d'API Web Azure](https://docs.microsoft.com/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)) pour obtenir un jeton d'accès pour Microsoft Graph en échange du jeton d'accès de démarrage.
1. Azure AD renvoie le jeton d'accès à Microsoft Graph (et un jeton d’actualisation si le complément demande l’autorisation *offline_access*) au complément.
1. Le code côté serveur met en cache le jeton d'accès pour Microsoft Graph.
1. Le code côté serveur envoie des requêtes à Microsoft Graph et inclut le jeton d'accès à Microsoft Graph.
1. Microsoft Graph renvoie des données au complément, qui peut les transmettre à l’interface utilisateur du complément.
1. Lorsque le jeton d'accès à Microsoft Graph expire, le code côté serveur peut utiliser son jeton d'actualisation pour obtenir un nouveau jeton d'accès à Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Développer un complément SSO qui accède à Microsoft Graph

Vous développez un complément qui accède à Microsoft Graph, de la même façon que le feriez pour n’importe quel autre complément qui utilise l’authentification unique. Pour obtenir une description détaillée, voir [Activer l’authentification unique pour les compléments Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins). La différence est qu’il est obligatoire que le complément ait une API de Web côté serveur, et ce qu’on appelle le jeton d’accès appelé, dans cet article, le « jeton d’accès d’amorçage ». 

Selon votre langage et votre infrastructure, des bibliothèques peuvent être disponibles pour simplifier le code côté serveur que vous devez écrire. Votre code doit effectuer les opérations suivantes :

* Valider le jeton d’accès d’amorçage de complément qui est reçu à partir du gestionnaire de jeton que vous avez créé précédemment. Pour plus d’informations, voir [Valider le jeton d’accès](sso-in-office-add-ins.md#validate-the-access-token). 
* Démarrer le flux « de la part de » avec un appel au point de terminaison Azure AD version 2.0 qui inclut le jeton d’accès d’amorçage, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et secret).
* Mettre en cache le jeton d’accès renvoyé à Microsoft Graph. Pour plus d’informations sur ce flux voir [Azure Active Directory v2.0  et flux OAuth 2.0  « de la part de »](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Créer une ou plusieurs méthodes d’API Web qui obtiennent des données Microsoft Graph, en transmettant le jeton d’accès mis en cache vers Microsoft Graph.

> [!NOTE]
> Pour des exemples de jetons d'accès décodés pour Microsoft Graph qui ont été obtenus « de la part du » flux, consultez la rubrique [Azure Active Directory version 2.0 et OAuth 2.0 de la part du lux](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Pour une aide détaillée pas à pas et des exemples de scénarios, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
* [Scénario : implémenter l’authentification unique sur votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



