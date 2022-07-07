---
title: Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)
description: Conseils sur la résolution des problèmes liés à l’authentification unique (SSO) dans les compléments Office et la gestion de conditions ou d’erreurs spéciales.
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e155b1da472e9e9e081bf43b1660996583f97cc1
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659949"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso"></a>Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)

Cet article fournit des conseils sur la résolution des problèmes liés à l’authentification unique (SSO) dans les compléments Office, et explique comment faire en sorte que votre complément gère correctement les conditions particulières ou les erreurs.

> [!NOTE]
> L’API d’authentification unique est actuellement prise en charge pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’emplacement où l’API d’authentification unique est actuellement prise en charge, consultez [ensembles de conditions requises IdentityAPI](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location Microsoft 365. Pour plus d’informations sur la procédure à suivre, consultez [Exchange Online : comment activer votre locataire pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="debugging-tools"></a>Outils de débogage

Lors du développement, nous vous recommandons vivement d’utiliser un outil capable d’intercepter et d’afficher les demandes HTTP du service web de votre complément, ainsi que les réponses. Les deux outils les plus populaires sont les suivants :

- [Fiddler](https://www.telerik.com/fiddler) : Gratuit ([documentation](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com) : Gratuit pendant 30 jours. ([Documentation](https://www.charlesproxy.com/documentation/))

## <a name="causes-and-handling-of-errors-from-getaccesstoken"></a>Causes et gestion des erreurs de getAccessToken

Pour consulter des exemples de la gestion des erreurs décrite dans cette section, reportez-vous aux articles suivants :
- [HomeES6.js dans Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js)
- [ssoAuthES6.js dans Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/public/javascripts/ssoAuthES6.js)

### <a name="13000"></a>13000

L’API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) n’est pas prise en charge par le complément ou la version d’Office.

- La version d’Office ne prend pas en charge la SSO. La version requise est l’abonnement Microsoft 365, dans n’importe quel canal mensuel.
- Le manifeste de complément n’inclut pas la section [WebApplicationInfo](/javascript/api/manifest/webapplicationinfo) appropriée.

Votre complément doit corriger cette erreur en basculant vers un autre système d’authentification des utilisateurs. Pour plus d’informations, voir [Meilleures Pratiques et Conditions Requises](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

### <a name="13001"></a>13001

L’utilisateur n’est pas connecté à Office. Dans la plupart des cas, vous devez éviter cette erreur en transmettant l’option `allowSignInPrompt: true` dans le paramètre `AuthOptions`.

Il peut toutefois y avoir des exceptions. Par exemple, vous souhaitez que le complément s’ouvre avec des fonctionnalités qui nécessitent un utilisateur connecté, mais *uniquement si* l’utilisateur a *déjà* ouvert une session dans Office. Si ce n’est pas le cas, vous voulez que le complément s’ouvre avec un autre groupe de fonctionnalités qui n’exigent pas que l’utilisateur soit connecté. Dans ce cas, la logique qui s’exécute lorsque le complément lance des appels `getAccessToken` sans `allowSignInPrompt: true`. Utilisez l’erreur 13001 comme indicateur pour indiquer au complément de présenter l’autre groupe de fonctionnalités.

Une autre option consiste à répondre au 13001 en basculant vers un autre système d’authentification des utilisateurs. Cette opération permet de connecter l’utilisateur à AAD, mais pas à Office.

Cette erreur n’est jamais apparue dans **Office sur le web**. Si le cookie de l’utilisateur a expiré, **Office sur le web** renvoie l’erreur 13006.

### <a name="13002"></a>13002

L’utilisateur a annulé la connexion ou l’autorisation, par exemple, en choisissant **Annuler** dans la boîte de dialogue d’autorisation.

- Si votre complément fournit des fonctions qui ne nécessitent pas la connexion (ou le consentement) de l’utilisateur, votre code doit intercepter cette erreur et autoriser l’exécution du complément.
- Si le complément requiert un utilisateur connecté qui a donné son accord, votre code doit inclure un bouton de connexion qui s’affiche.

### <a name="13003"></a>13003

Type d’utilisateur non pris en charge. L’utilisateur n’est pas connecté à Office avec un compte Microsoft valide ou un compte Microsoft 365 Éducation ou professionnel. Cela peut se produire si Office est exécuté avec un compte de domaine en local, par exemple. Votre code doit basculer vers un autre système d’authentification des utilisateurs. Dans Outlook, cette erreur peut également se produire si [l’authentification moderne est désactivée](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online) pour le locataire de l’utilisateur dans Exchange Online. Pour plus d’informations, voir [Meilleures Pratiques et Conditions Requises](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

### <a name="13004"></a>13004

Ressource non valide. (Cette erreur ne doit être visible qu’en développement.) Le manifeste du complément n’a pas été configuré correctement. Mettez à jour le manifeste. Pour en savoir plus, consultez [Valider le manifeste d’un complément Office](../testing/troubleshoot-manifest.md). Le problème le plus courant est que l’élément **\<Resource\>** (dans l’élément **\<WebApplicationInfo\>** ) a un domaine qui ne correspond pas au domaine du complément. Bien que la partie protocole de la valeur Resource devrait être « api » et non « https », toutes les autres parties du nom de domaine (dont le port éventuel) doivent être identiques à ceux du complément.

### <a name="13005"></a>13005

Octroi non valide. Cela signifie généralement qu’Office n’a pas été pré-autorisé sur le service web du complément. Pour plus d’informations, consultez la rubrique sur la [création de l’application de service](sso-in-office-add-ins.md#register-your-add-in-with-the-microsoft-identity-platform) et sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md). Cela peut également arriver si l’utilisateur n’a pas accordé à votre service les autorisations d’application pour son `profile` ou a révoqué l’accord. Votre code doit basculer vers un autre système d’authentification des utilisateurs.

Une autre cause possible, lors du développement, est que votre complément utilise Internet Explorer et que vous utilisez un certificat auto-signé. (Pour déterminer quel navigateur est utilisé par le complément, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).)

### <a name="13006"></a>13006

Erreur du client. Cette erreur apparaît uniquement dans **Office sur le web**. Votre code doit suggérer à l’utilisateur de se déconnecter et de redémarrer la session de navigateur Office.

### <a name="13007"></a>13007

L’application Office n’a pas pu obtenir de jeton d’accès au service web du complément.

- Si cette erreur se produit en cours de développement, assurez-vous que l’enregistrement de votre complément, ainsi que son manifeste, spécifient l’autorisation `profile` (et l’autorisation `openid` si vous utilisez MSAL.NET). Pour plus d’informations, voir [Inscrire votre complément avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).
- En production, plusieurs causes peuvent entraîner cette erreur. En voici certaines :
  - L’utilisateur a une identité de compte Microsoft.
  - Certaines situations qui entraîneraient l’une des autres erreurs 13xxx avec un Microsoft 365 Éducation ou un compte professionnel entraînent un 13007 lorsqu’un MSA est utilisé.

  Dans tous ces cas, votre code doit basculer vers un autre système d’authentification des utilisateurs.

### <a name="13008"></a>13008

L’utilisateur a déclenché une opération qui appelle `getAccessToken` avant d’avoir terminé une opération qui appelle `getAccessToken`. Cette erreur apparaît uniquement dans **Office sur le web**. Votre code doit demander à l’utilisateur de répéter l’opération une fois que l’opération précédente sera terminée.

### <a name="13010"></a>13010

L’utilisateur exécute le complément dans Office sur Microsoft Edge. Le domaine Microsoft 365 de l’utilisateur et le `login.microsoftonline.com` domaine se trouvent dans des zones de sécurité différentes dans les paramètres du navigateur. Cette erreur apparaît uniquement dans **Office sur le web**. Si cette erreur est renvoyée, l’utilisateur a déjà vu une erreur expliquant cela et menant vers une page sur la modification de la configuration de la zone. Si votre complément fournit des fonctions qui ne nécessitent pas que l’utilisateur soit connecté, votre code doit intercepter cette erreur et autoriser l’exécution du complément.

### <a name="13012"></a>13012

Il existe plusieurs causes possibles.

- Le complément est en cours d’exécution sur une plateforme qui ne prend pas en charge l’API `getAccessToken`. Par exemple, elle n’est pas compatible avec iPad. Consultez également [les ensembles de conditions requises de l’API d’identité](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
- L’option `forMSGraphAccess` a été transmise à l’appel à `getAccessToken` et l’utilisateur a obtenu le complément à partir d’AppSource. Dans ce scénario, l’administrateur du client n’a pas donné son accord au complément pour les étendues Microsoft Graph (autorisations) dont il a besoin. Le fait de rappeler `getAccessToken` avec le `allowConsentPrompt` ne résoudra pas le problème, car Office est autorisé à inviter l’utilisateur à donner l’autorisation uniquement à l’étendue de `profile` AAD.

Votre code doit basculer vers un autre système d’authentification des utilisateurs.

En développement, le complément est sideloaded dans Outlook et l’option `forMSGraphAccess` a été transmise dans l’appel à `getAccessToken`.

### <a name="13013"></a>13013

L’appel `getAccessToken` a été appelé trop de fois dans un court laps de temps, office a donc limité l’appel le plus récent. Cela est généralement dû à une boucle infinie d’appels à la méthode. Il existe des scénarios lorsque le rappel de la méthode est recommandé. Toutefois, votre code doit utiliser un compteur ou une variable d’indicateur pour vous assurer que la méthode n’est pas rappelée à plusieurs reprises. Si le même chemin d’accès de code « nouvelle tentative » s’exécute à nouveau, le code doit revenir à un autre système d’authentification utilisateur. Pour obtenir un exemple de code, voir comment la `retryGetAccessToken` variable est utilisée dans [HomeES6.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js) ou [ssoAuthES6.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO/Complete/public/javascripts/ssoAuthES6.js).

### <a name="50001"></a>50001

Cette erreur (qui n’est pas spécifique de `getAccessToken`) peut indiquer que le navigateur a mis en cache une ancienne copie des fichiers office.js. Quand vous développez, effacez le cache du navigateur. Une autre possibilité est que la version d’Office n’est pas assez récente pour prendre en charge l’authentification unique. Dans Windows, la version minimale est 16.0.12215.20006. Sur Mac, il s’agit de 16.32.19102902.

Dans un complément de production, celui-ci doit répondre à cette erreur en basculant vers un autre système d’authentification des utilisateurs. Pour plus d’informations, voir [Meilleures Pratiques et Conditions Requises](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erreurs d’Azure Active Directory côté serveur

Pour plus d’exemples de la gestion des erreurs décrite dans cette section, reportez-vous aux articles suivants :
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)

### <a name="conditional-access--multifactor-authentication-errors"></a>Erreurs d’accès conditionnel / authentification multifacteur

Dans certaines configurations d’identité dans AAD et Microsoft 365, il est possible que certaines ressources accessibles avec Microsoft Graph nécessitent une authentification multifacteur (MFA), même si la location Microsoft 365 de l’utilisateur ne le fait pas. Lorsqu’AAD reçoit une requête pour obtenir un jeton d’accès à la ressource protégée par AMF via le flux « de la part de », il renvoie au service web de votre complément un message JSON contenant une propriété `claims`. La propriété de revendication comporte des informations sur les facteurs d’authentification supplémentaires nécessaires.

Votre code doit tester cette propriété `claims`. En fonction de l’architecture de votre complément, vous pouvez le tester côté client, ou le tester sur le serveur et le relayer sur le client. Il vous faut ces informations dans le client, car Office gère l’authentification des compléments SSO. Si vous le relayez côté serveur, le message adressé au client peut être une erreur (telle que `500 Server Error` ou `401 Unauthorized`) ou se trouver dans le corps d’une réponse de succès (telle que `200 OK`). Dans les deux cas, le rappel (réussite ou échec) de l’appel AJAX de votre code côté client à l’API web de votre complément devra tester cette réponse.

Quelle que soit votre architecture, si la valeur des revendications a été envoyée à partir d’AAD, votre code doit rappeler `getAccessToken` et passer l’option `authChallenge: CLAIMS-STRING-HERE` dans le `options` paramètre. Lorsqu’AAD voit cette chaîne, il demande le(s) facteur(s) supplémentaire(s) à l’utilisateur, puis renvoie un nouveau jeton d’accès qui sera accepté dans le flux « de la part de ».

### <a name="consent-missing-errors"></a>Erreurs de consentement manquant

Si AAD ne détient aucune trace qu’un consentement (à la ressource Microsoft Graph) a été accordé au complément par l’utilisateur (ou administrateur client), AAD envoie un message d’erreur à votre service web. Votre code doit indiquer au client (dans le corps d’une réponse `403 Forbidden`, par exemple).

Si le complément a besoin d’étendues Microsoft Graph qui ne peuvent être envoyées qu’à un administrateur, votre code doit générer une erreur. Si les seules étendues requises peuvent être envoyées par l’utilisateur, votre code doit basculer vers un autre système d’authentification des utilisateurs.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erreurs d’étendue (autorisation) non valide ou manquante

Ce type d’erreur ne doit apparaître qu’en développement.

- Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client qui doit consigner l’erreur dans la console ou l’enregistrer dans un journal.
- Assurez-vous que la section [Scopes](/javascript/api/manifest/scopes) du manifeste de votre complément indique toutes les autorisations nécessaires. Vérifiez également que l’alignement du service web de votre complément spécifie les mêmes autorisations. Vérifiez les fautes d’orthographe. Pour plus d’informations, voir [Inscrire votre complément avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).

### <a name="invalid-audience-error-in-the-access-token-for-microsoft-graph"></a>Erreur d’audience non valide dans le jeton d’accès pour Microsoft Graph

Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client, qui doit présenter un message amical à l’utilisateur et éventuellement consigner l’erreur dans la console ou l’enregistrer dans un journal.
