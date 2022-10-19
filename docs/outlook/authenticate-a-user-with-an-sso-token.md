---
title: Authentifier un utilisateur avec un jeton à authentification unique
description: Découvrez comment utiliser le jeton d’authentification unique fourni par un complément Outlook pour implémenter l’authentification unique (SSO) sur votre service.
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 23b7936cc0ba4453a2a10cbfe0731941a913c118
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607442"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>Authentifier un utilisateur avec un jeton d’authentification unique dans un complément Outlook

L’authentification unique (SSO) permet à votre complément d’authentifier les utilisateurs en toute transparence (et éventuellement d’obtenir des jetons d’accès pour appeler l’[API Microsoft Graph](/graph/overview)).

Grâce à cette méthode, votre complément peut obtenir un jeton d’accès inclus dans l’API principale de votre serveur. Le complément l’utilise comme un jeton du porteur dans l’en-tête `Authorization` pour authentifier un rappel de votre API. Si vous le souhaitez, vous pouvez également avoir votre code côté serveur.

- renseigner le flux De la part de pour obtenir un jeton d’accès inclus dans l’API Microsoft Graph ;
- utiliser les informations d’identité dans le jeton pour établir l’identité de l’utilisateur et s’authentifier à vos services principaux.

Pour une vue d’ensemble de l’authentification unique dans les compléments Office, reportez-vous à [Activer l’authentification unique pour des compléments Office ](../develop/sso-in-office-add-ins.md) et [Autorisation de l’accès à Microsoft Graph dans votre complément Office](../develop/authorize-to-microsoft-graph.md).

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Activer l’authentification moderne dans votre location Microsoft 365

Pour utiliser l’authentification unique avec un complément Outlook, vous devez activer l’authentification moderne pour la location Microsoft 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="register-your-add-in"></a>Inscription de votre complément

Pour utiliser l’authentification unique, votre complément Outlook devra avoir une API web côté serveur enregistrée auprès d’Azure Active Directory (AAD) v2.0. Pour obtenir plus d’informations, reportez-vous à [Enregistrer un complément Office qui utilise l’authentification unique auprès du point de terminaison Azure AD v2.0](../develop/register-sso-add-in-aad-v2.md).

### <a name="provide-consent-when-sideloading-an-add-in"></a>Consentement fourni pendant le chargement indépendant d’un complément

Lorsque vous développez un complément, vous devez donner votre consentement à l’avance. Pour plus d’informations, consultez [Accorder le consentement de l’administrateur au complément](../develop/grant-admin-consent-to-an-add-in.md).

## <a name="update-the-add-in-manifest"></a>Mise à jour du manifeste de complément

L’étape suivante pour activer l’authentification unique dans le complément consiste à ajouter des informations au manifeste à partir de l’inscription Plateforme d'identités Microsoft du complément. Le balisage varie en fonction du type de manifeste.

- **Manifeste XML** : ajoutez un `WebApplicationInfo` élément à la fin de l’élément `VersionOverridesV1_1` [VersionOverrides](/javascript/api/manifest/versionoverrides) . Ajoutez ensuite ses éléments enfants requis. Pour plus d’informations sur le balisage, consultez [Configurer le complément](../develop/sso-in-office-add-ins.md#configure-the-add-in).
- **Manifeste Teams (préversion)** : ajoutez une propriété « webApplicationInfo » à l’objet racine `{ ... }` dans le manifeste. Attribuez à cet objet une propriété « id » enfant définie sur l’ID d’application de l’application web du complément tel qu’il a été généré dans le Portail Azure lorsque vous avez inscrit le complément. (Consultez la section [Inscrire votre complément](#register-your-add-in) plus haut dans cet article.) Attribuez également une propriété « ressource » enfant qui est définie sur le même **URI d’ID d’application** que celui que vous avez défini lors de l’inscription du complément. Cet URI doit avoir la forme `api://<fully-qualified-domain-name>/<application-id>`. Voici un exemple.

   ```json
   "webApplicationInfo": {
        "id": "a661fed9-f33d-4e95-b6cf-624a34a2f51d",
        "resource": "api://addin.contoso.com/a661fed9-f33d-4e95-b6cf-624a34a2f51d"
    },
   ```

  > [!NOTE]
  > Les compléments prenant en charge l’authentification unique qui utilisent le manifeste Teams peuvent être chargés de manière indépendante, mais ne peuvent pas être déployés d’une autre manière pour l’instant.

## <a name="get-the-sso-token"></a>Obtention du jeton SSO

Le complément obtient un jeton SSO avec le script côté client. Pour plus d’informations, reportez-vous à [Ajouter du code côté client](../develop/sso-in-office-add-ins.md#add-client-side-code).

## <a name="use-the-sso-token-at-the-back-end"></a>Utilisation du jeton SSO dans le back-end

Dans la plupart des scénarios, il n’est pas vraiment utile d’obtenir le jeton d’accès si votre complément ne le transmet pas côté serveur et ne l’utilise pas à cet emplacement. Pour plus d’informations sur ce que votre côté serveur peut et doit faire, reportez-vous à la section [Ajouter du code côté serveur](../develop/sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code).

> [!IMPORTANT]
> Quand vous utilisez le jeton SSO sous forme d’identité dans le complément *Outlook*, nous vous recommandons d’[utiliser le jeton d’identité Exchange](authenticate-a-user-with-an-identity-token.md) comme identité alternative. Les utilisateurs de votre complément peuvent utiliser plusieurs clients, dont certains ne fourniront peut-être pas de jeton SSO. En utilisant le jeton d’identité Exchange comme alternative, vous pouvez éviter d’inviter les utilisateurs à entrer leurs informations d’identification plusieurs fois. Pour plus d’informations, voir[Scénario : Implémenter l’authentification unique sur votre service dans un complément Outlook](implement-sso-in-outlook-add-in.md).

## <a name="sso-for-event-based-activation"></a>Authentification unique pour l’activation basée sur les événements

Il existe des étapes supplémentaires à suivre si votre complément utilise l’activation basée sur les événements. Pour plus d’informations, consultez Activer l’authentification [unique (SSO) dans les compléments Outlook qui utilisent l’activation basée sur les événements](use-sso-in-event-based-activation.md).

## <a name="see-also"></a>Voir aussi

- [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1))
- Pour obtenir un exemple de complément Outlook qui utilise le jeton d’authentification unique pour accéder à Microsoft API Graph, consultez l’authentification unique du [complément Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).
- [Référence d’API SSO](/javascript/api/office/office.auth#office-office-auth-getaccesstoken-member(1))
- [Ensemble d’exigences IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [Activer l’authentification unique (SSO) dans les compléments Outlook qui utilisent l’activation basée sur les événements](use-sso-in-event-based-activation.md)
