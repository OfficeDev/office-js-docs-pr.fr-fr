---
title: Autoriser la connexion à Microsoft Graph avec l’authentification unique
description: Découvrez comment les utilisateurs d’un complément Office peuvent utiliser l’authentification unique (SSO) pour extraire des données de Microsoft Graph.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: c04026324c71bcfbef978106aecf607d46d55fc3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608403"
---
# <a name="authorize-to-microsoft-graph-with-sso-preview"></a>Autoriser la connexion à Microsoft Graph avec l’authentification unique (préversion)

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365). Le meilleur moyen pour un complément Office d’obtenir un accès autorisé à [Microsoft Graph](https://developer.microsoft.com/graph/docs) est d’utiliser les informations d’identification Office de l’utilisateur. Cela leur permet d’accéder à leurs données Microsoft Graph sans avoir à se connecter une deuxième fois. 

> [!NOTE]
> L’API de l’authentification unique est actuellement prise en charge en mode aperçu pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](../reference/requirement-sets/identity-api-requirement-sets.md). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Architecture de complément pour l’authentification unique et Microsoft Graph

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du complément contient un balisage qui spécifie comment le complément est enregistré dans le point de terminaison Azure Active Directory (Azure AD) v2.0 et il indique les autorisations à Microsoft Graph dont le complément a besoin.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant montre comment fonctionne le processus de connexion et l’accès à Microsoft Graph.

![Un diagramme illustrant le processus d’authentification unique](../images/sso-access-to-microsoft-graph.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-). Cela indique à l’application hôte Office qu’elle doit obtenir un jeton d’accès au complément. (Ci-après, il est appelé **jeton d’accès bootstrap**, car il est remplacé par un deuxième jeton plus loin dans le processus. Pour consulter un exemple de jeton d’accès bootstrap décodé, voir [Exemple jeton d’accès](sso-in-office-add-ins.md#example-access-token).)
2. Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.
3. Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.
4. L’application hôte Office demande le **jeton d’accès bootstrap** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.
5. Azure AD envoie le jeton bootstrap à l’application hôte Office.
6. L’application hôte Office envoie le**jeton d’accès bootstrap** au complément dans le cadre de l’objet de résultat renvoyé par l’appel`getAccessToken`.
7. JavaScript dans le complément effectue une requête HTTP à une API web qui est hébergée sur le même domaine complet que le complément et inclut le **jeton d’accès bootstrap** comme preuve d’autorisation.
8. Le code côté serveur valide le **jeton d’accès bootstrap** entrant.
9. Le code côté serveur utilise le flux « de la part de » (défini dans [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) et le [scénario ou l’application serveur vers le scénario Azure de l’API Web](/azure/active-directory/develop/active-directory-authentication-scenarios)) pour obtenir un jeton d’accès pour Microsoft Graph dans Exchange pour le jeton d’accès bootstrap.
10. Azure AD renvoie le jeton d’accès à Microsoft Graph (et un jeton d’actualisation si le complément demande l’autorisation *offline_access*) au complément.
11. Le code côté serveur met en cache le jeton d’accès à Microsoft Graph.
12. Le code côté serveur effectue des requêtes à Microsoft Graph et inclut le jeton d’accès à Microsoft Graph.
13. Microsoft Graph renvoie des données au complément, qui peuvent le transmettre à l’interface utilisateur du complément.
14. Lorsque le jeton d’accès à Microsoft Graph expire, le code côté serveur peut utiliser son jeton d’actualisation pour obtenir un nouveau jeton d’accès à Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Développer un complément authentification unique qui accède à Microsoft Graph

Vous développez un complément qui accède à Microsoft Graph comme vous le feriez pour n’importe quel autre complément qui utilise l’authentification unique. Pour obtenir une description complète, voir [Activer l’authentification unique pour les compléments Office](../develop/sso-in-office-add-ins.md). La différence est qu’il est obligatoire que le complément ait une API Web côté serveur, et ce qu’on appelle le jeton d’accès dans cet article s’appelle le « jeton d’accès bootstrap ».

Selon votre langue et votre infrastructure, des bibliothèques peuvent être disponibles pour simplifier le code côté serveur que vous devez rédiger. Votre code côté serveur doit effectuer les opérations suivantes :

* Lancez le flux « de la part de » avec un appel vers le point de terminaison Azure AD v 2.0 qui inclut le jeton d’accès bootstrap, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (son ID et sa clé secrète).
* Créer une ou plusieurs méthodes API Web qui obtiennent des données de Microsoft Graph en transmettant le jeton d’accès (potentiellement mis en cache) à Microsoft Graph.
* De manière facultative, avant d’initier le flux, validez le jeton d’accès bootstrap reçu à partir du gestionnaire de jetons que vous avez créé précédemment. Pour plus d’informations, voir [Valider le jeton d’accès](sso-in-office-add-ins.md#validate-the-access-token). 
* De manière facultative, une fois le flux terminé, mettez en cache le jeton d’accès renvoyé vers Microsoft Graph. Nous vous conseillons de le faire si le complément effectue plusieurs appels à Microsoft Graph. Pour plus d’informations sur ce flux, voir [Azure Active Directory v2.0 et OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> Pour consulter des exemples de jeton d’accès décodés pour Microsoft Graph qui ont été obtenus par le flux « de la part de », voir [Azure Active Directory v2.0 et OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Pour obtenir des exemples de scénarios et procédures détaillées, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
* [Scénario : Implémenter l’authentification unique pour votre service dans un complément Outlook](../outlook/implement-sso-in-outlook-add-in.md)
