---
title: Autoriser Microsoft Graph dans votre complément Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: f6e7de146d2f03256aa673a0653c1e03f9340d86
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876591"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Autorisez Microsoft Graph dans votre complément Office (préversion)

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365). Le meilleur moyen pour un complément Office d'obtenir un accès autorisé à [Microsoft Graph](https://developer.microsoft.com/graph/docs) est d'utiliser les informations d'identification de connexion Office de l'utilisateur. Cela leur permet d'accéder à leurs données Microsoft Graph sans avoir besoin de se connecter une seconde fois. 

> [!NOTE]
> L’API de l’authentification unique est actuellement prise en charge en mode aperçu pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).
> Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Architecture de complément pour SSO et Microsoft Graph

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](https://msdn.microsoft.com/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du complément contient un balisage qui spécifie comment le complément est enregistré dans le point de terminaison Azure Active Directory (Azure AD) v2.0 et il indique les autorisations à Microsoft Graph dont le complément a besoin.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant montre le fonctionnement du processus de connexion et d'accès à Microsoft Graph.

![Diagramme illustrant le processus d’authentification unique](../images/sso-access-to-microsoft-graph.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Cette option indique à l’application hôte Office qu'il faut obtenir un jeton d’accès pour le complément (ci-après, il s’agit du **jeton d’accès des données d’amorçage** , car il est remplacé par un deuxième jeton plus loin dans le processus. Pour obtenir un exemple d’un jeton d’accès d’amorçage décodé, voir [Exemple de jeton d’accès](sso-in-office-add-ins.md#example-access-token).)
1. Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.
1. Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.
1. L’application hôte Office demande le **jeton d'accès d'amorçage** au point de terminaison Azure AD version 2.0 pour l’utilisateur actuel.
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

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Développez un complément SSO qui accède à Microsoft Graph

Vous développez un complément qui accède à Microsoft Graph comme vous le feriez pour tout autre complément utilisant l'authentification unique. Pour une description détaillée, voir [Activez la connexion unique pour les compléments Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins). La différence est qu'il est obligatoire pour le complément d'avoir une API Web côté serveur, et ce qu'on appelle le jeton d'accès dans cet article s'appelle le « jeton d'accès de démarrage ». 

Selon votre langue et votre environnement de travail, des bibliothèques peuvent être disponibles pour simplifier le code côté serveur que vous devez écrire. Votre code doit effectuer les opérations suivantes :

* Validez le jeton d'accès d'amorçage du complément, reçu du gestionnaire de jetons que vous avez créé précédemment. Pour plus d'informations, consultez la rubrique [Valider le jeton d'accès](sso-in-office-add-ins.md#validate-the-access-token). 
* Démarrez le flux « de la part de » avec un appel au point de terminaison Azure AD version 2.0 qui inclut le jeton d'accès de démarrage, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et clé secrète).
* Mettez en cache le jeton d'accès renvoyé à Microsoft Graph. Pour obtenir plus d'informations sur ce flux, consultez la rubrique[Azure Active Directory v2.0 et flux OAuth 2.0 « de la part de »](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Créez une ou plusieurs méthodes d'API web qui obtiennent des données Microsoft Graph, en transmettant le jeton d'accès mis en cache vers Microsoft Graph.

> [!NOTE]
> Pour des exemples de jetons d'accès décodés pour Microsoft Graph qui ont été obtenus « de la part du » flux, consultez la rubrique [Azure Active Directory version 2.0 et OAuth 2.0 de la part du lux](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Pour une aide détaillée pas à pas et des exemples de scénarios, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
* [Scénario : implémentez l’authentification unique sur votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



