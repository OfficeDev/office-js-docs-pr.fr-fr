---
title: Autoriser Microsoft Graph dans votre complément Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 4f584010d38a5e96a9863233854300184a24660f
ms.sourcegitcommit: eea7f2b1679cf9a209d35880b906e311bdf1359c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/26/2018
ms.locfileid: "21241165"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Autorisez Microsoft Graph dans votre complément Office (préversion)

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365). Le meilleur moyen pour un complément Office d'obtenir un accès autorisé à [Microsoft Graph](https://developer.microsoft.com/graph/docs) est d'utiliser les informations d'identification de connexion Office de l'utilisateur. Cela leur permet d'accéder à leurs données Microsoft Graph sans avoir besoin de se connecter une seconde fois. 

> [!NOTE]
> L’API de l’authentification unique est actuellement prise en charge en mode préversion pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).
> Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour en savoir plus sur la manière de procéder, consultez la rubrique [Exchange Online : Activativez votre client pour une authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Architecture de complément pour SSO et Microsoft Graph

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](https://msdn.microsoft.com/en-us/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du complément contient un balisage qui spécifie comment le complément est enregistré dans le point de terminaison Azure Active Directory (Azure AD) v2.0 et il indique les autorisations à Microsoft Graph dont le complément a besoin.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant montre le fonctionnement du processus de connexion et d'accès à Microsoft Graph.

![Diagramme illustrant le processus d’authentification unique](../images/sso-access-to-microsoft-graph.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js `getAccessTokenAsync`. Cela indique à l’application hôte Office qu’elle doit obtenir un jeton d’accès au complément. (Ci-après, cela s'appelle le **jeton d'accès de démarrage** car il est remplacé par un second jeton plus tard dans le processus. Pour un exemple de jeton d'accès de démarrage décodé, consultez la rubrique [Exemple de jeton d'accès](sso-in-office-add-ins.md#example-access-token).
1. Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.
1. Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.
1. L’application hôte Office demande le **jeton d'accès de démarrage** au point de terminaison Azure AD version 2.0 pour l’utilisateur à jour.
1. Azure AD envoie le jeton de démarrage à l’application hôte Office.
1. L’application hôte Office envoie le **jeton d'accès de démarrage** au complément en tant que partie de l’objet résultat renvoyé par l’appel `getAccessTokenAsync`.
1. Un code JavaScript dans le complément effectue une requête HTTP vers une API web hébergée sur le même domaine complet que le complément et inclut le **jeton d'accès de démarrage** comme preuve d’autorisation.  
1. Le code côté serveur valide le **jeton d'accès de démarrage**entrant.
1. Le code côté serveur utilise le flux « au nom de » (défini par l'[Échange de jetons OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) et l'[application démon ou serveur pour un scénario d'API Web Azure](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)) pour obtenir un jeton d'accès pour Microsoft Graph en échange du jeton d'accès de démarrage.
1. Azure AD renvoie le jeton d'accès à Microsoft Graph (et un jeton d’actualisation si le complément demande l’autorisation offline_access) au complément.**
1. Le code côté serveur met en cache le jeton d'accès pour Microsoft Graph.
1. Le code côté serveur envoie des requêtes à Microsoft Graph et inclut le jeton d'accès à Microsoft Graph.
1. Microsoft Graph renvoie des données au complément, qui peut les transmettre à l’interface utilisateur du complément.
1. Lorsque le jeton d'accès à Microsoft Graph expire, le code côté serveur peut utiliser son jeton d'actualisation pour obtenir un nouveau jeton d'accès à Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Développez un complément SSO qui accède à Microsoft Graph

Vous développez un complément qui accède à Microsoft Graph comme vous le feriez pour tout autre complément utilisant l'authentification unique. Pour une description détaillée, voir [Activez la connexion unique pour les compléments Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/sso-in-office-add-ins). La différence est qu'il est obligatoire pour le complément d'avoir une API Web côté serveur, et ce qu'on appelle le jeton d'accès dans cet article s'appelle le « jeton d'accès de démarrage ». 

Selon votre langue et votre infrastructure, des bibliothèques peuvent être disponibles pour simplifier le code côté serveur que vous devez écrire. Votre code doit effectuer les opérations suivantes :

* Validez le jeton d'accès de démarrage du complément, reçu du gestionnaire de jetons que vous avez créé précédemment. Pour plus d'informations, consultez la rubrique [Valider le jeton d'accès](sso-in-office-add-ins.md#validate-the-access-token). 
* Démarrez le « de la part du » flux avec un appel au point de terminaison Azure AD version 2.0 qui inclut le jeton d'accès de démarrage, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et clé secrète).
* Mettez en cache le jeton d'accès renvoyé à Microsoft Graph. Pour obtenir plus d'informations sur ce flux, consultez la rubrique[Azure Active Directory v2.0 et flux OAuth 2.0 « de la part de »](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Créez une ou plusieurs méthodes d'API web qui obtiennent des données Microsoft Graph, en transmettant le jeton d'accès mis en cache vers Microsoft Graph.

> [!NOTE]
> Pour des exemples de jetons d'accès décodés pour Microsoft Graph qui ont été obtenus « de la part du » flux, consultez la rubrique [Azure Active Directory version 2.0 et OAuth 2.0 de la part du lux](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Pour une aide détaillée pas à pas et des exemples de scénarios, consultez les rubriques suivantes :

* [Créez un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créez un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
* [Scénario : implémentez l’authentification unique sur votre service dans un complément Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in)



