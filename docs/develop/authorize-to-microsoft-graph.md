---
title: Autoriser la connexion à Microsoft Graph avec l’authentification unique
description: Découvrez comment les utilisateurs d’un Office peuvent utiliser l’sign-on unique (SSO) pour extraire des données de Microsoft Graph.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 538648e96233bd0c2b497ef588d10c4f708e8522
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320264"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>Autoriser la connexion à Microsoft Graph avec l’authentification unique

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel, de leur compte professionnel, ou scolaire (Office 365). Le meilleur moyen pour un complément Office d’obtenir un accès autorisé à [Microsoft Graph](https://developer.microsoft.com/graph/docs) est d’utiliser les informations d’identification Office de l’utilisateur. Cela leur permet d’accéder à leurs données Microsoft Graph sans avoir à se connecter une deuxième fois.

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Architecture de complément pour l’authentification unique et Microsoft Graph

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du add-in contient un élément **WebApplicationInfo** qui fournit des informations d’inscription d’application Azure importantes à Office, y compris les autorisations sur Microsoft Graph dont le module complémentaire a besoin.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant montre les étapes nécessaires pour se connecter et accéder à Microsoft Graph. L’ensemble du processus utilise les jetons d’accès OAuth 2.0 et JWT.

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="Diagramme montrant le processus DSO." border="false":::

1. Le code côté client du add-in appelle l’API Office.js [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_). Cela indique à l Office’hôte d’obtenir un jeton d’accès pour le module.

    Si l’utilisateur n’est pas signé, l’hôte Office conjointement avec le Plateforme d'identités Microsoft fournit une interface utilisateur pour la signature et le consentement de l’utilisateur.

2. L’Office demande un jeton d’accès à l’Plateforme d'identités Microsoft.
3. Le Plateforme d'identités Microsoft renvoie le jeton *d’accès A* à l Office hôte. Le *jeton d’accès A* fournit uniquement l’accès aux API côté serveur du add-in. Il ne fournit pas d’accès à Microsoft Graph.
4. L Office’hôte renvoie le jeton *d’accès A* au code côté client du module. Le code côté client peut désormais effectuer des appels authentifiés aux API côté serveur.
5. Le code côté client effectue une demande HTTP à une API web côté serveur qui nécessite une authentification. Il inclut le jeton *d’accès A* comme preuve d’autorisation. Le code côté serveur valide le jeton *d’accès A*.
6. Le code côté serveur utilise le flux OAuth 2.0 On-Behalf-Of (OBO) pour demander un nouveau jeton d’accès avec des autorisations pour Microsoft Graph.
7. Le Plateforme d'identités Microsoft renvoie le nouveau jeton d’accès *B* avec des autorisations pour Microsoft Graph (et un jeton d’actualisation, si le *offline_access demande une* autorisation). Le serveur peut éventuellement mettre en cache le jeton *d’accès B*.
8. Le code côté serveur effectue une demande à une API Microsoft Graph et inclut le jeton *d’accès B* avec des autorisations pour Microsoft Graph.
9. Microsoft Graph les données dans le code côté serveur.
10. Le code côté serveur renvoie les données au code côté client.

Lors des demandes suivantes, le code client passe toujours le jeton d’accès *A* lors des appels authentifiés au code côté serveur. Le code côté serveur peut mettre en cache le jeton *B* afin qu’il n’a pas besoin de le demander à nouveau lors des futurs appels d’API.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Développer un complément authentification unique qui accède à Microsoft Graph

Vous développez un add-in qui accède à Microsoft Graph comme vous le feriez pour n’importe quelle autre application qui utilise l’luiso. Pour obtenir une description détaillée, voir [Activer l’sign-on Office des modules.](../develop/sso-in-office-add-ins.md) La différence est qu’il est obligatoire que le add-in a une API Web côté serveur.

Selon votre langue et votre infrastructure, des bibliothèques peuvent être disponibles pour simplifier le code côté serveur que vous devez rédiger. Votre code côté serveur doit effectuer les opérations suivantes :

* Validez le jeton *d’accès A* chaque fois qu’il est transmis à partir du code côté client. Pour plus d’informations, voir [Valider le jeton d’accès](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).
* Lancez le flux OAuth 2.0 On-Behalf-Of (OBO) avec un appel au Plateforme d'identités Microsoft qui inclut le jeton d’accès, certaines métadonnées sur l’utilisateur et les informations d’identification du module (son ID et sa question secrète). Pour plus d’informations sur le flux OBO, [voir Plateforme d'identités Microsoft et OAuth 2.0 On-Behalf-Of flow](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).
* Éventuellement, une fois le flux terminé, mettre en cache le jeton *d’accès renvoyé B* avec des autorisations à Microsoft Graph. Nous vous conseillons de le faire si le complément effectue plusieurs appels à Microsoft Graph. Pour plus d’informations, voir [Acquérir et mettre en cache des jetons à l’aide de la bibliothèque d’authentification Microsoft (MSAL)](/azure/active-directory/develop/msal-acquire-cache-tokens)
* Créez une ou plusieurs méthodes d’API Web qui obtiennent des données microsoft Graph en passant le jeton d’accès (éventuellement mis en cache) *B* à Microsoft Graph.

Pour obtenir des exemples de scénarios et procédures détaillées, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
* [Scénario : Implémenter l’authentification unique pour votre service dans un complément Outlook](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Distribution de modules ssO dans Microsoft AppSource

Lorsqu’un administrateur Microsoft 365 acquiert un add-in à [](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) partir [d’AppSource](https://appsource.microsoft.com), il peut le redistribuer via les applications intégrées et accorder l’autorisation à l’administrateur d’accéder aux étendues Graph Microsoft. Toutefois, il est également possible pour l’utilisateur final d’acquérir le add-in directement à partir d’AppSource, auquel cas l’utilisateur doit donner son consentement au module. Cela peut créer un problème de performances potentiel pour lequel nous avons fourni une solution.

Si votre code `allowConsentPrompt` passe l’option `getAccessToken`dans l’appel de , `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`par exemple, Office peut demander à l’utilisateur son consentement si l’Plateforme d'identités Microsoft signale à Office que ce consentement n’a pas encore été accordé au module. Toutefois, pour des raisons de sécurité, Office peut uniquement invite l’utilisateur à consentir à l’étendue Graph `profile` Microsoft. *Office pas être invité à consentir à d’autres étendues Graph Microsoft*, pas même `User.Read`. Cela signifie que si l’utilisateur donne son consentement à l’invite, Office renvoie un jeton d’accès. Toutefois, la tentative d’échange du jeton d’accès contre un nouveau jeton d’accès avec des étendues Microsoft Graph supplémentaires échoue avec l’erreur AADSTS65001, ce qui signifie que le consentement (aux étendues Microsoft Graph) n’a pas été accordé.

> [!NOTE]
> La demande de consentement avec peut `{ allowConsentPrompt: true }` toujours échouer même `profile` pour l’étendue si l’administrateur a désactivé le consentement de l’utilisateur final. Pour plus d’informations, voir [Configurer la façon dont les utilisateurs finaux consentent aux applications à](/azure/active-directory/manage-apps/configure-user-consent) l’aide Azure Active Directory.

Votre code peut et doit gérer cette erreur en revenir à un autre système d’authentification, qui invite l’utilisateur à donner son consentement aux étendues Graph Microsoft. Pour obtenir des exemples de code, voir Créer un Node.js Office qui utilise [l’sign-on](create-sso-office-add-ins-nodejs.md) unique et [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md) and the samples they link to. L’ensemble du processus nécessite plusieurs allers-retours vers le Plateforme d'identités Microsoft. Pour éviter cette pénalité de performances, incluez l’option `forMSGraphAccess` dans l’appel `getAccessToken`de ; par exemple, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`. Cela signale Office que votre application a besoin de Microsoft Graph étendues. Office demande au Plateforme d'identités Microsoft de vérifier que le consentement aux étendues Graph Microsoft a déjà été accordé au module. Si c’est le cas, le jeton d’accès est renvoyé. Si ce n’est pas le cas, l’appel de `getAccessToken` renvoie l’erreur 13012. Votre code peut gérer cette erreur en revenir immédiatement à un autre système d’authentification, sans tenter d’échanger des jetons avec le Plateforme d'identités Microsoft.

En tant que meilleure pratique, passez toujours aux moments où votre application sera distribuée dans AppSource et nécessite des étendues `forMSGraphAccess` `getAccessToken` Graph Microsoft.

## <a name="details-on-sso-with-an-outlook-add-in"></a>Détails sur l’sso avec un Outlook de données

Si vous développez un Outlook qui utilise l’luiso et que vous chargez une version test, Office retourne toujours l’erreur 13012  `forMSGraphAccess` `getAccessToken` lorsqu’il est passé, même si le consentement de l’administrateur a été accordé. Pour cette raison, vous devez commenter l’option `forMSGraphAccess` lors **du développement d’un** Outlook de développement. N’oubliez pas de désafcommenter l’option lorsque vous déployez pour la production. La fausse version 13012 se produit uniquement lorsque vous chargez une version de version Outlook.

Pour Outlook des applications, assurez-vous d’activer l’authentification moderne pour Microsoft 365 location. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="see-also"></a>Voir aussi

* [Jeton OAuth2 Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Plateforme d'identités Microsoft flux « De la part de » et OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [Ensembles de conditions requises IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md)
