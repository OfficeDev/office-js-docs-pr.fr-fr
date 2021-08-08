---
title: Autoriser la connexion à Microsoft Graph avec l’authentification unique
description: Découvrez comment les utilisateurs d’un Office peuvent utiliser l' sign-on unique (SSO) pour extraire des données de Microsoft Graph.
ms.date: 07/27/2021
localization_priority: Normal
ms.openlocfilehash: a4302d05d796b53f6db602dcd12f8c03469fc240927b2bc326ffa9f07a5d8954
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081234"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>Autoriser la connexion à Microsoft Graph avec l’authentification unique

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel, de leur compte professionnel, ou scolaire (Office 365). Le meilleur moyen pour un complément Office d’obtenir un accès autorisé à [Microsoft Graph](https://developer.microsoft.com/graph/docs) est d’utiliser les informations d’identification Office de l’utilisateur. Cela leur permet d’accéder à leurs données Microsoft Graph sans avoir à se connecter une deuxième fois.

> [!NOTE]
> La connexion unique sur API est actuellement prise en charge pour Word, Excel et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](../reference/requirement-sets/identity-api-requirement-sets.md).
> Si vous travaillez avec un add-in Outlook, assurez-vous d'activer l'authentification moderne pour la location de Microsoft 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Architecture de complément pour l’authentification unique et Microsoft Graph

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du complément contient un balisage qui spécifie comment le complément est enregistré dans le point de terminaison Azure Active Directory (Azure AD) v2.0 et il indique les autorisations à Microsoft Graph dont le complément a besoin.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant montre comment fonctionne le processus de connexion et l’accès à Microsoft Graph.

![Diagramme montrant le processus DSO.](../images/sso-access-to-microsoft-graph.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_). Cela indique à l’application cliente Office qu’elle doit obtenir un jeton d’accès au complément. (Ci-après, il est appelé **jeton d’accès bootstrap**, car il est remplacé par un deuxième jeton plus loin dans le processus. Pour consulter un exemple de jeton d’accès bootstrap décodé, voir [Exemple jeton d’accès](sso-in-office-add-ins.md#example-access-token).)
2. Si l’utilisateur n’est pas connecté, l’application cliente Office ouvre une fenêtre contextuelle pour qu’il se connecte.
3. Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.
4. L Office application cliente demande le jeton d’accès **bootstrap** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.
5. Azure AD envoie le jeton d’a bootstrap à l Office application cliente.
6. L Office’application cliente envoie le jeton d’accès **bootstrap** au module dans le cadre de l’objet de résultat renvoyé par `getAccessToken` l’appel.
7. JavaScript dans le complément effectue une requête HTTP à une API web qui est hébergée sur le même domaine complet que le complément et inclut le **jeton d’accès bootstrap** comme preuve d’autorisation.
8. Le code côté serveur valide le **jeton d’accès bootstrap** entrant.
9. Le code côté serveur utilise le flux « de la part de » (défini dans le Exchange de jeton [OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) et l’application de [daemon](/azure/active-directory/develop/active-directory-authentication-scenarios)ou serveur pour le scénario Azure de l’API web) pour obtenir un jeton d’accès pour Microsoft Graph en échange du jeton d’accès bootstrap.
10. Azure AD renvoie le jeton d’accès à Microsoft Graph (et un jeton d’actualisation si le complément demande l’autorisation *offline_access*) au complément.
11. Le code côté serveur met en cache le jeton d’accès à Microsoft Graph.
12. Le code côté serveur effectue des requêtes à Microsoft Graph et inclut le jeton d’accès à Microsoft Graph.
13. Microsoft Graph renvoie des données au complément, qui peut les transmettre à l’interface utilisateur du complément.
14. Lorsque le jeton d’accès à Microsoft Graph expire, le code côté serveur peut utiliser son jeton d’actualisation pour obtenir un nouveau jeton d’accès à Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Développer un complément authentification unique qui accède à Microsoft Graph

Vous développez un complément qui accède à Microsoft Graph comme vous le feriez pour n’importe quel autre complément qui utilise l’authentification unique. Pour obtenir une description complète, voir [Activer l’authentification unique pour les compléments Office](../develop/sso-in-office-add-ins.md). La différence est qu’il est obligatoire que le complément ait une API Web côté serveur, et ce qu’on appelle le jeton d’accès dans cet article s’appelle le « jeton d’accès bootstrap ».

Selon votre langue et votre infrastructure, des bibliothèques peuvent être disponibles pour simplifier le code côté serveur que vous devez rédiger. Votre code côté serveur doit effectuer les opérations suivantes :

* Lancez le flux « de la part de » avec un appel au point de terminaison Azure AD v2.0 qui inclut le jeton d’accès bootstrap, certaines métadonnées sur l’utilisateur et les informations d’identification du module (son ID et sa question secrète).
* Créer une ou plusieurs méthodes API Web qui obtiennent des données de Microsoft Graph en transmettant le jeton d’accès (potentiellement mis en cache) à Microsoft Graph.
* De manière facultative, avant d’initier le flux, validez le jeton d’accès bootstrap reçu à partir du gestionnaire de jetons que vous avez créé précédemment. Pour plus d’informations, voir [Valider le jeton d’accès](sso-in-office-add-ins.md#validate-the-access-token). 
* De manière facultative, une fois le flux terminé, mettez en cache le jeton d’accès renvoyé vers Microsoft Graph. Nous vous conseillons de le faire si le complément effectue plusieurs appels à Microsoft Graph. Pour plus d’informations sur ce flux, voir [Azure Active Directory v2.0 et OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> Pour consulter des exemples de jeton d’accès décodés pour Microsoft Graph qui ont été obtenus par le flux « de la part de », voir [Azure Active Directory v2.0 et OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Pour obtenir des exemples de scénarios et procédures détaillées, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
* [Scénario : Implémenter l’authentification unique pour votre service dans un complément Outlook](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Distribution de modules ssO dans Microsoft AppSource

Lorsqu’un administrateur Microsoft 365 acquiert un add-in à partir d’AppSource, il peut le redistribuer via les applications intégrées et accorder l’autorisation d’administrateur au add-in pour accéder aux étendues microsoft Graph. [](https://appsource.microsoft.com) [](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) Toutefois, il est également possible pour l’utilisateur final d’acquérir le add-in directement à partir d’AppSource, auquel cas l’utilisateur doit donner son consentement au module. Cela peut créer un problème de performances potentiel pour lequel nous avons fourni une solution.

Si votre code passe l’option dans l’appel de , par exemple, Office peut demander à l’utilisateur son consentement si Azure AD signale à Office que ce consentement `allowConsentPrompt` n’a pas encore été accordé au module. `getAccessToken` `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );` Toutefois, pour des raisons de sécurité, Office peut uniquement invite l’utilisateur à consentir à l’étendue Azure `profile` AD. *Office pas être invité à consentir* à des étendues Graph Microsoft, pas même `User.Read` . Cela signifie que si l’utilisateur donne son consentement à l’invite, Office renvoyer un jeton d’a bootstrap. Toutefois, la tentative d’échange du jeton d’a bootstrap contre un jeton d’accès à Microsoft Graph échouera avec l’erreur AADSTS65001, ce qui signifie que le consentement (aux étendues Microsoft Graph) n’a pas été accordé.

Votre code peut et doit gérer cette erreur en revenir à un autre système d’authentification, qui invite l’utilisateur à donner son consentement aux étendues Graph Microsoft. (Pour obtenir des exemples de code, voir Créer un Node.js Office qui utilise l' [sign-on](create-sso-office-add-ins-nodejs.md) unique et [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md) and the samples they link to.) Toutefois, l’ensemble du processus nécessite plusieurs allers-retours vers Azure AD. Vous pouvez éviter cette pénalité de performances en incluant `forMSGraphAccess` l’option dans l’appel de ; par `getAccessToken` exemple, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` .  Cela signale Office que votre application a besoin de Microsoft Graph étendues. Office demander à Azure AD de vérifier que le consentement aux étendues Graph Microsoft a déjà été accordé au module. Si c’est le cas, le jeton d’a bootstrap est renvoyé. Si ce n’est pas le cas, l’appel de `getAccessToken` retournera l’erreur 13012. Votre code peut gérer cette erreur en revenir immédiatement à un autre système d’authentification, sans tenter d’échanger des jetons avec Azure AD.

En tant que meilleure pratique, passez toujours aux moments où votre application sera distribuée dans AppSource et nécessite des étendues Graph `forMSGraphAccess` `getAccessToken` Microsoft.

> [!TIP]
> Si vous développez un Outlook qui utilise l' luiso et que vous  chargez une version test, Office retourne toujours l’erreur 13012 lorsqu’il est passé, même si le consentement de l’administrateur a été `forMSGraphAccess` `getAccessToken` accordé. Pour cette raison, vous devez commenter `forMSGraphAccess` l’option lors du développement **d’un** Outlook de développement. N’oubliez pas de désafcommenter l’option lorsque vous déployez pour la production. La fausse version 13012 se produit uniquement lorsque vous chargez une version de version Outlook.
