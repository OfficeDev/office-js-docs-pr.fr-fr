---
title: Autoriser la connexion à Microsoft Graph avec l’authentification unique
description: Découvrez comment les utilisateurs d’un complément Office peuvent utiliser l’authentification unique (SSO) pour extraire des données de Microsoft Graph.
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4c7bfc51e67755c2a50875f11d3a5477bd5885a4
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/15/2022
ms.locfileid: "66090942"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>Autoriser la connexion à Microsoft Graph avec l’authentification unique

Les utilisateurs se connectent à Office à l’aide de leur compte Microsoft personnel, de leur Microsoft 365 Éducation ou de leur compte professionnel. Le meilleur moyen pour un complément Office d’obtenir un accès autorisé à [Microsoft Graph](https://developer.microsoft.com/graph/docs) est d’utiliser les informations d’identification Office de l’utilisateur. Cela leur permet d’accéder à leurs données Microsoft Graph sans avoir à se connecter une deuxième fois.

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Architecture de complément pour l’authentification unique et Microsoft Graph

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du complément contient un élément **WebApplicationInfo** qui fournit des informations d’inscription d’application Azure importantes à Office, y compris les autorisations à Microsoft Graph requises par le complément.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant montre les étapes à suivre pour se connecter et accéder à Microsoft Graph. L’ensemble du processus utilise des jetons d’accès OAuth 2.0 et JWT.

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="Diagramme montrant le processus d’authentification unique." border="false":::

1. Le code côté client du complément appelle l’API [Office.js getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)). Cela indique à l’hôte Office d’obtenir un jeton d’accès pour le complément.

    Si l’utilisateur n’est pas connecté, l’hôte Office conjointement avec le Plateforme d'identités Microsoft fournit une interface utilisateur permettant à l’utilisateur de se connecter et de donner son consentement.

2. L’hôte Office demande un jeton d’accès auprès du Plateforme d'identités Microsoft.
3. Le Plateforme d'identités Microsoft retourne le jeton *d’accès A* à l’hôte Office. Le jeton *d’accès A* fournit uniquement l’accès aux propres API côté serveur du complément. Il ne fournit pas d’accès à Microsoft Graph.
4. L’hôte Office retourne le jeton *d’accès A* au code côté client du complément. À présent, le code côté client peut effectuer des appels authentifiés aux API côté serveur.
5. Le code côté client envoie une requête HTTP à une API web côté serveur qui nécessite une authentification. Il inclut le jeton d’accès *A* comme preuve d’autorisation. Le code côté serveur valide le jeton d’accès *A*.
6. Le code côté serveur utilise le flux OBO (On-Behalf-Of) OAuth 2.0 pour demander un nouveau jeton d’accès avec des autorisations à Microsoft Graph.
7. Le Plateforme d'identités Microsoft retourne le nouveau jeton d’accès *B* avec des autorisations à Microsoft Graph (et un jeton d’actualisation, si le complément demande *offline_access* autorisation). Le serveur peut éventuellement mettre en cache le jeton d’accès *B*.
8. Le code côté serveur effectue une demande à un API Graph Microsoft et inclut le jeton d’accès *B* avec des autorisations sur Microsoft Graph.
9. Microsoft Graph renvoie les données au code côté serveur.
10. Le code côté serveur retourne les données au code côté client.

Lors des demandes suivantes, le code client passe toujours le jeton *d’accès A* lors de l’exécution d’appels authentifiés au code côté serveur. Le code côté serveur peut mettre en cache le jeton *B* afin qu’il n’ait pas besoin de le demander à nouveau lors de futurs appels d’API.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Développer un complément authentification unique qui accède à Microsoft Graph

Vous développez un complément qui accède à Microsoft Graph comme vous le feriez pour toute autre application qui utilise l’authentification unique. Pour obtenir une description détaillée, consultez [Activer l’authentification unique pour Office compléments](../develop/sso-in-office-add-ins.md). La différence est qu’il est obligatoire que le complément dispose d’une API web côté serveur.

Selon votre langue et votre infrastructure, des bibliothèques peuvent être disponibles pour simplifier le code côté serveur que vous devez rédiger. Votre code côté serveur doit effectuer les opérations suivantes :

* Validez le jeton d’accès *A* chaque fois qu’il est passé à partir du code côté client. Pour plus d’informations, voir [Valider le jeton d’accès](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).
* Lancez le flux OBO (On-Behalf-Of) OAuth 2.0 avec un appel au Plateforme d'identités Microsoft qui inclut le jeton d’accès, certaines métadonnées sur l’utilisateur et les informations d’identification du complément (son ID et son secret). Pour plus d’informations sur le flux OBO, consultez [Plateforme d'identités Microsoft et le flux On-Behalf-Of OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow).
* Si vous le souhaitez, une fois le flux terminé, mettez en cache le jeton d’accès *retourné B* avec des autorisations sur Microsoft Graph. Nous vous conseillons de le faire si le complément effectue plusieurs appels à Microsoft Graph. Pour plus d’informations, consultez [Acquérir et mettre en cache des jetons à l’aide de la bibliothèque d’authentification Microsoft (MSAL)](/azure/active-directory/develop/msal-acquire-cache-tokens)
* Créez une ou plusieurs méthodes d’API web qui obtiennent des données Microsoft Graph en transmettant le jeton d’accès (éventuellement mis en cache) *B* à Microsoft Graph.

Pour obtenir des exemples de scénarios et procédures détaillées, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md)
* [Scénario : Implémenter l’authentification unique pour votre service dans un complément Outlook](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Distribution de compléments prenant en charge l’authentification unique dans Microsoft AppSource

Lorsqu’un administrateur Microsoft 365 acquiert un complément à partir [d’AppSource](https://appsource.microsoft.com), il peut le redistribuer par le biais [d’applications intégrées](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) et accorder au complément le consentement de l’administrateur pour accéder aux étendues Microsoft Graph. Toutefois, il est également possible pour l’utilisateur final d’acquérir le complément directement à partir d’AppSource, auquel cas l’utilisateur doit accorder son consentement au complément. Cela peut créer un problème de performances potentiel pour lequel nous avons fourni une solution.

Si votre code passe l’option `allowConsentPrompt` dans l’appel de `getAccessToken`, par exemple`OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`, Office pouvez demander le consentement de l’utilisateur si le Plateforme d'identités Microsoft signale à Office que le consentement n’a pas encore été accordé au complément. Toutefois, pour des raisons de sécurité, Office pouvez uniquement inviter l’utilisateur à donner son consentement à l’étendue microsoft Graph`profile`. *Office ne peut pas demander le consentement à d’autres étendues Graph Microsoft*, même `User.Read`pas . Cela signifie que si l’utilisateur donne son consentement à l’invite, Office retourne un jeton d’accès. Toutefois, la tentative d’échange du jeton d’accès contre un nouveau jeton d’accès avec des étendues microsoft Graph supplémentaires échoue avec l’erreur AADSTS65001, ce qui signifie que le consentement (aux étendues Microsoft Graph) n’a pas été accordé.

> [!NOTE]
> La demande de consentement peut `{ allowConsentPrompt: true }` toujours échouer même pour l’étendue `profile` si l’administrateur a désactivé le consentement de l’utilisateur final. Pour plus d’informations, consultez [Configurer la façon dont les utilisateurs finaux consentent aux applications à l’aide de Azure Active Directory](/azure/active-directory/manage-apps/configure-user-consent).

Votre code peut et doit gérer cette erreur en rebasculant vers un autre système d’authentification, qui invite l’utilisateur à donner son consentement aux étendues Microsoft Graph. Pour obtenir des exemples de code, consultez [Créer un complément Node.js Office qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md) et [Créer un complément ASP.NET Office qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md) et les exemples auxquels ils sont liés. L’ensemble du processus nécessite plusieurs allers-retours vers le Plateforme d'identités Microsoft. Pour éviter cette pénalité de performances, incluez l’option `forMSGraphAccess` dans l’appel de `getAccessToken`; par exemple, `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`. Cela signale à Office que votre complément a besoin d’étendues Microsoft Graph. Office demandez au Plateforme d'identités Microsoft de vérifier que le consentement aux étendues Microsoft Graph a déjà été accordé au complément. Si c’est le cas, le jeton d’accès est retourné. Si ce n’est pas le cas, l’appel de `getAccessToken` retour renvoie l’erreur 13012. Votre code peut gérer cette erreur en rebasculant immédiatement vers un autre système d’authentification, sans tenter d’échanger des jetons avec le Plateforme d'identités Microsoft.

En guise de bonne pratique, passez `forMSGraphAccess` toujours à `getAccessToken` quel moment votre complément sera distribué dans AppSource et nécessite des étendues Microsoft Graph.

## <a name="details-on-sso-with-an-outlook-add-in"></a>Détails sur l’authentification unique avec un complément Outlook

Si vous développez un complément Outlook qui utilise l’authentification unique et que vous la chargez de manière indépendante pour le test, Office retourne *toujours* l’erreur 13012 quand `forMSGraphAccess` elle est passée même `getAccessToken` si le consentement de l’administrateur a été accordé. Pour cette raison, vous devez commenter l’option **lors** du `forMSGraphAccess` développement d’un complément Outlook. Veillez à annuler les marques de commentaire de l’option lorsque vous déployez pour la production. Le faux 13012 ne se produit que lorsque vous chargez des versions test dans Outlook.

Pour Outlook compléments, veillez à activer l’authentification moderne pour la location Microsoft 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="see-also"></a>Voir aussi

* [Exchange de jeton OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Plateforme d’identités Microsoft et flux OAuth 2.0 On-Behalf-Of](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [Ensembles de conditions requises IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
