---
title: Options d’authentification dans les compléments Outlook
description: Les compléments Outlook offrent différentes méthodes qui permettent de s’authentifier en fonction de votre scénario.
ms.date: 10/17/2022
ms.localizationpriority: high
ms.openlocfilehash: d8ae8971c4095e5314885514226cd8f52728fb07
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607526"
---
# <a name="authentication-options-in-outlook-add-ins"></a>Options d’authentification dans les compléments Outlook

Votre complément Outlook peut accéder à des informations à partir de n’importe quel emplacement sur Internet, qu’il s’agisse du serveur qui héberge le complément, de votre réseau interne ou du cloud. Si ces informations sont protégées, votre complément doit trouver un moyen d’authentifier votre utilisateur. Les compléments Outlook offrent différentes méthodes qui permettent de s’authentifier en fonction de votre scénario.

## <a name="single-sign-on-access-token"></a>Jeton d’accès à authentification unique

Les jetons d’accès à authentification unique permettent à votre complément de s’authentifier en toute transparence et d’obtenir des jetons d’accès pour appeler l’[API Microsoft Graph](/graph/overview). Cette fonctionnalité réduit la friction étant donné que l’utilisateur n’a pas besoin de saisir ses informations d’identification.

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Vous pouvez utiliser des jetons d’accès d’authentification unique si votre complément :

- Est principalement utilisé par les utilisateurs de Microsoft 365
- doit accéder à ce qui suit :
  - Services Microsoft exposés dans le cadre de Microsoft Graph ;
  - Service non-Microsoft que vous contrôlez.

La méthode d’authentification unique utilise le flux [OAuth2 De la part de fourni par Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of). Cela nécessite l’enregistrement du complément dans le [portail d’inscription des applications](https://apps.dev.microsoft.com/) et la spécification de toute étendue Microsoft Graph requise dans son manifeste.

> [!NOTE]
> Si le complément utilise le [manifeste Teams pour les compléments Office (préversion),](../develop/json-manifest-overview.md) il existe une configuration de manifeste, mais les étendues Microsoft Graph ne sont pas spécifiées. Les compléments prenant en charge l’authentification unique qui utilisent le manifeste Teams peuvent être chargés de manière indépendante, mais ne peuvent pas être déployés d’une autre manière pour l’instant.

Grâce à cette méthode, votre complément peut obtenir un jeton d’accès inclus dans l’API principale de votre serveur. Le complément l’utilise comme un jeton du porteur dans l’en-tête `Authorization` pour authentifier un rappel de votre API. À ce stade, votre serveur peut :

- renseigner le flux De la part de pour obtenir un jeton d’accès inclus dans l’API Microsoft Graph ;
- utiliser les informations d’identité dans le jeton pour établir l’identité de l’utilisateur et s’authentifier à vos propres services principaux.

Pour obtenir un aperçu plus détaillé, consultez la [présentation complète de la méthode d’authentification unique](../develop/sso-in-office-add-ins.md).

Pour plus d’informations sur l’utilisation du jeton à authentification unique dans un complément Outlook, consultez la section [Authentifier un utilisateur avec un jeton à authentification unique dans un complément Outlook](authenticate-a-user-with-an-sso-token.md).

Pour un complément échantillon qui utilise le jeton à authentification unique, consultez [Authentification unique d’un complément Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO).

## <a name="exchange-user-identity-token"></a>Jeton d’identité d’utilisateur Exchange

Les jetons d’identité d’utilisateur Exchange permettent à votre complément d’établir l’identité de l’utilisateur. En vérifiant l’identité de l’utilisateur, vous pouvez ensuite effectuer une authentification unique dans votre système principal, puis accepter le jeton d’identité d’utilisateur comme une autorisation pour les demandes futures. Utilisez le jeton d’identité d’utilisateur Exchange :

- quand le complément est utilisé principalement par des utilisateurs locaux Exchange ;
- quand le complément doit accéder à un service non-Microsoft que vous contrôlez ;
- En tant qu’authentification de secours quand le complément est exécuté sur une version d’Office qui ne prend pas en charge SSO.

Votre complément peut appeler la méthode [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getuseridentitytokenasync-member(1)) pour obtenir des jetons d’identité d’utilisateur Exchange. Pour plus d’informations sur l’utilisation de ces jetons, voir [Authentifier un utilisateur avec un jeton d’identité pour Exchange](authenticate-a-user-with-an-identity-token.md).

## <a name="access-tokens-obtained-via-oauth2-flows"></a>Jetons d’accès obtenus via les flux OAuth2

Les compléments peuvent également accéder à des services de Microsoft et d’autres entreprises qui prennent en charge OAuth2 pour l’autorisation. Vous pouvez utiliser les jetons OAuth2 si votre complément :

- Doit accéder à un service en dehors de votre contrôle.

Grâce à cette méthode, votre module complémentaire invite l'utilisateur à se connecter au service en utilisant la méthode [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) pour initialiser le flux OAuth2.

## <a name="callback-tokens"></a>Jetons de rappel

Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api). Consider using callback tokens if your add-in:

- Doit accéder à la boîte aux lettres de l’utilisateur à partir de votre serveur principal.

Les compléments obtiennent des jetons de rappel à l’aide d’une méthode [getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods). Le niveau d’accès est contrôlé par les autorisations spécifiées dans le manifeste du complément.
