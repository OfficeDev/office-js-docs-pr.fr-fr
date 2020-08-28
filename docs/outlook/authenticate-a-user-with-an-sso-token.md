---
title: Authentifier un utilisateur avec un jeton à authentification unique
description: Découvrez comment utiliser le jeton d’authentification unique fourni par un complément Outlook pour implémenter l’authentification unique (SSO) sur votre service.
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: e0925979d26f6b3145658d71b1edaf30431e0c7e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293981"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>Authentifier un utilisateur avec un jeton d’authentification unique dans un complément Outlook

L’authentification unique (SSO) permet à votre complément d’authentifier les utilisateurs en toute transparence (et éventuellement d’obtenir des jetons d’accès pour appeler l’[API Microsoft Graph](/graph/overview)).

Grâce à cette méthode, votre complément peut obtenir un jeton d’accès inclus dans l’API principale de votre serveur. Le complément l’utilise comme un jeton du porteur dans l’en-tête `Authorization` pour authentifier un rappel de votre API. Si vous le souhaitez, vous pouvez également obtenir votre code côté serveur. Pour ce faire, procédez comme suit :

- renseigner le flux De la part de pour obtenir un jeton d’accès inclus dans l’API Microsoft Graph ;
- utiliser les informations d’identité dans le jeton pour établir l’identité de l’utilisateur et s’authentifier à vos services principaux.

Pour une vue d’ensemble de l’authentification unique dans les compléments Office, reportez-vous à [Activer l’authentification unique pour des compléments Office ](../develop/sso-in-office-add-ins.md) et [Autorisation de l’accès à Microsoft Graph dans votre complément Office](../develop/authorize-to-microsoft-graph.md).

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Activer l’authentification moderne dans votre location Microsoft 365

Pour utiliser l’authentification unique (SSO) avec un complément Outlook, vous devez activer l’authentification moderne pour la location Microsoft 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="register-your-add-in"></a>Inscription de votre complément

Pour utiliser l’authentification unique, votre complément Outlook devra avoir une API web côté serveur enregistrée auprès d’Azure Active Directory (AAD) v2.0. Pour obtenir plus d’informations, reportez-vous à [Enregistrer un complément Office qui utilise l’authentification unique auprès du point de terminaison Azure AD v2.0](../develop/register-sso-add-in-aad-v2.md).

### <a name="provide-consent-when-sideloading-an-add-in"></a>Consentement fourni pendant le chargement indépendant d’un complément

Lorsque vous développez un complément, vous devez fournir un consentement à l’avance. Pour plus d’informations, consultez [la rubrique accorder le consentement de l’administrateur au complément](../develop/grant-admin-consent-to-an-add-in.md).

## <a name="update-the-add-in-manifest"></a>Mise à jour du manifeste de complément

Pour activer l’authentification unique dans le complément, vous devez ensuite ajouter un élément `WebApplicationInfo` à la fin de l’élément `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md). Pour plus d’informations, reportez-vous à [Configurer le complément](../develop/sso-in-office-add-ins.md#configure-the-add-in).

## <a name="get-the-sso-token"></a>Obtention du jeton SSO

Le complément obtient un jeton SSO avec le script côté client. Pour plus d’informations, reportez-vous à [Ajouter du code côté client](../develop/sso-in-office-add-ins.md#add-client-side-code).

## <a name="use-the-sso-token-at-the-back-end"></a>Utilisation du jeton SSO dans le back-end

Dans la plupart des scénarios, il n’est pas vraiment utile d’obtenir le jeton d’accès si votre complément ne le transmet pas côté serveur et ne l’utilise pas à cet emplacement. Pour plus d’informations sur ce que votre côté serveur peut et doit faire, reportez-vous à la section [Ajouter du code côté serveur](../develop/sso-in-office-add-ins.md#add-server-side-code).

> [!IMPORTANT]
> Quand vous utilisez le jeton SSO sous forme d’identité dans le complément *Outlook*, nous vous recommandons d’[utiliser le jeton d’identité Exchange](authenticate-a-user-with-an-identity-token.md) comme identité alternative. Les utilisateurs de votre complément peuvent utiliser plusieurs clients, dont certains ne fourniront peut-être pas de jeton SSO. En utilisant le jeton d’identité Exchange comme alternative, vous pouvez éviter d’inviter les utilisateurs à entrer leurs informations d’identification plusieurs fois. Pour plus d’informations, voir[Scénario : Implémenter l’authentification unique sur votre service dans un complément Outlook](implement-sso-in-outlook-add-in.md).

## <a name="see-also"></a>Voir aussi

- Pour obtenir un exemple de complément Outlook qui utilise le jeton SSO pour accéder à l’API Microsoft Graph, consultez la rubrique relative à l' [authentification unique du complément Outlook](https://github.com/OfficeDev/Outlook-Add-in-SSO).
- [Référence d’API SSO](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [Ensemble d’exigences IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md)
