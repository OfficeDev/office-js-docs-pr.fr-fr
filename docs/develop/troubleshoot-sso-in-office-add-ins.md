---
title: Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)
description: Conseils sur la résolution des problèmes liés à l’authentification unique (SSO) dans les compléments Office et la gestion des conditions ou des erreurs spéciales.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 49e967aa0d500df64828c66d9dee8574eb948cec
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093559"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>Résolution des messages d’erreur pour l’authentification unique (SSO) (aperçu)

Cet article fournit des conseils sur la résolution des problèmes liés à l’authentification unique (SSO) dans les compléments Office, et explique comment faire en sorte que votre complément gère correctement les conditions particulières ou les erreurs.

> [!NOTE]
> L’API de l’authentification unique est actuellement prise en charge en mode aperçu pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations sur l’endroit où l’API d’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](../reference/requirement-sets/identity-api-requirement-sets.md).
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location Microsoft 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="debugging-tools"></a>Outils de débogage

We strongly recommend that you use a tool that can intercept and display the HTTP Requests from, and Responses to, your add-in's web service when you are developing. Two of the most popular are:

- [Fiddler](https://www.telerik.com/fiddler) : Gratuit ([documentation](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com) : Gratuit pendant 30 jours. ([Documentation](https://www.charlesproxy.com/documentation/))

## <a name="causes-and-handling-of-errors-from-getaccesstoken"></a>Causes et gestion des erreurs de getAccessToken

Pour consulter des exemples de la gestion des erreurs décrite dans cette section, reportez-vous aux articles suivants :
- [HomeES6.js dans Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js)
- [ssoAuthES6.js dans Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js)

### <a name="13000"></a>13000

L’API [getAccessToken](../develop/sso-in-office-add-ins.md#sso-api-reference) n’est pas prise en charge par le complément ou la version d’Office.

- La version d’Office ne prend pas en charge la SSO. La version requise est un abonnement Microsoft 365, quel que soit le canal mensuel.
- Le manifeste de complément n’inclut pas la section [WebApplicationInfo](../reference/manifest/webapplicationinfo.md) appropriée.

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

Type d’utilisateur non pris en charge. L’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte professionnel ou scolaire Microsoft 365. Cela peut se produire si Office est exécuté avec un compte de domaine en local, par exemple. Votre code doit basculer vers un autre système d’authentification des utilisateurs. Pour plus d’informations, voir [Meilleures Pratiques et Conditions Requises](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

### <a name="13004"></a>13004

Ressource non valide. (Cette erreur doit uniquement être vue en développement.) Le manifeste du complément n’a pas été configuré correctement. Mettez à jour le manifeste. Pour en savoir plus, consultez [Valider le manifeste d’un complément Office](../testing/troubleshoot-manifest.md). Le problème le plus courant est que l’élément **Resource** (dans l’élément **WebApplicationInfo**) a un domaine qui ne correspond pas au domaine du complément. Bien que la partie protocole de la valeur Resource devrait être « api » et non « https », toutes les autres parties du nom de domaine (dont le port éventuel) doivent être identiques à ceux du complément.

### <a name="13005"></a>13005

Octroi non valide. Cela signifie généralement qu’Office n’a pas été pré-autorisé sur le service web du complément. Pour plus d’informations, consultez la rubrique sur la [création de l’application de service](sso-in-office-add-ins.md#create-the-service-application) et sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md). Cela peut également arriver si l’utilisateur n’a pas accordé à votre service les autorisations d’application pour son `profile` ou a révoqué l’accord. Votre code doit basculer vers un autre système d’authentification des utilisateurs.

Une autre cause possible, lors du développement, est que votre complément utilise Internet Explorer et que vous utilisez un certificat auto-signé. (Pour déterminer quel navigateur est utilisé par le complément, consultez [Navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).)

### <a name="13006"></a>13006

Erreur du client. Cette erreur apparaît uniquement dans **Office sur le web**. Votre code doit suggérer à l’utilisateur de se déconnecter et de redémarrer la session de navigateur Office.

### <a name="13007"></a>13007

L’hôte Office n’a pas pu obtenir de jeton d’accès au service web du complément.

- Si cette erreur se produit en cours de développement, assurez-vous que l’enregistrement de votre complément, ainsi que son manifeste, spécifient l’autorisation `profile` (et l’autorisation `openid` si vous utilisez MSAL.NET). Pour plus d’informations, voir [Inscrire votre complément avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).
- En production, plusieurs causes peuvent entraîner cette erreur. En voici certaines :
    - L’utilisateur dispose d’une identité de compte Microsoft (MSA).
    - Certaines situations entraînant l’ouverture d’une des autres erreurs 13xxx avec un compte d’éducation ou de travail Microsoft 365 entraînent une 13007 lors de l’utilisation d’un MSA.

  Dans tous ces cas, votre code doit basculer vers un autre système d’authentification des utilisateurs.

### <a name="13008"></a>13008

L’utilisateur a déclenché une opération qui appelle `getAccessToken` avant d’avoir terminé une opération qui appelle `getAccessToken`. Cette erreur apparaît uniquement dans **Office sur le web**. Votre code doit demander à l’utilisateur de répéter l’opération une fois que l’opération précédente sera terminée.

### <a name="13010"></a>13010

L’utilisateur exécute le complément dans Office sur Microsoft Edge ou Internet Explorer. Le domaine Microsoft 365 de l’utilisateur, ainsi que le `login.microsoftonline.com` domaine, se trouvent dans des zones de sécurité différentes dans les paramètres du navigateur. Cette erreur apparaît uniquement dans **Office sur le web**. Si cette erreur est renvoyée, l’utilisateur a déjà vu une erreur expliquant cela et menant vers une page sur la modification de la configuration de la zone. Si votre complément fournit des fonctions qui ne nécessitent pas que l’utilisateur soit connecté, votre code doit intercepter cette erreur et autoriser l’exécution du complément.

### <a name="13012"></a>13012

Il existe plusieurs causes possibles :

- Le complément est en cours d’exécution sur une plateforme qui ne prend pas en charge l’API `getAccessToken`. Par exemple, elle n’est pas compatible avec iPad. Voir également [Ensembles de conditions requises de l’API d’identité](../reference/requirement-sets/identity-api-requirement-sets.md).
- L’option `forMSGraphAccess` a été transmise à l’appel à `getAccessToken` et l’utilisateur a obtenu le complément à partir d’AppSource. Dans ce scénario, l’administrateur du client n’a pas donné son accord au complément pour les étendues Microsoft Graph (autorisations) dont il a besoin. Le fait de rappeler `getAccessToken` avec le `allowConsentPrompt` ne résoudra pas le problème, car Office est autorisé à inviter l’utilisateur à donner l’autorisation uniquement à l’étendue de `profile` AAD.

Votre code doit basculer vers un autre système d’authentification des utilisateurs.

En développement, le complément est sideloaded dans Outlook et l’option `forMSGraphAccess` a été transmise dans l’appel à `getAccessToken`.

### <a name="13013"></a>13013

Le `getAccessToken` a été appelé trop souvent en un peu de temps, donc Office a limité l’appel le plus récent. Cela est généralement dû à une boucle infinie d’appels à la méthode. Il existe des scénarios pour rappeler la méthode. Toutefois, votre code doit utiliser un compteur ou une variable d’indicateur pour s’assurer que la méthode n’est pas rappelée à plusieurs reprises. Si le chemin d’accès du code « nouvelle tentative » s’exécute à nouveau, le code doit revenir à un autre système d’authentification des utilisateurs. Pour obtenir un exemple de code, consultez la rubrique `retryGetAccessToken` utilisation de la variable dans [HomeES6.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js) ou [ssoAuthES6.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js).

### <a name="50001"></a>50001

Cette erreur (qui n’est pas spécifique de `getAccessToken`) peut indiquer que le navigateur a mis en cache une ancienne copie des fichiers office.js. Quand vous développez, effacez le cache du navigateur. Une autre possibilité est que la version d’Office n’est pas assez récente pour prendre en charge l’authentification unique. Dans Windows, la version minimale est 16.0.12215.20006. Sur Mac, il s’agit de 16.32.19102902.

Dans un complément de production, celui-ci doit répondre à cette erreur en basculant vers un autre système d’authentification des utilisateurs. Pour plus d’informations, voir [Meilleures Pratiques et Conditions Requises](../develop/sso-in-office-add-ins.md#requirements-and-best-practices).

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erreurs d’Azure Active Directory côté serveur

Pour plus d’exemples de la gestion des erreurs décrite dans cette section, reportez-vous aux articles suivants :
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)

### <a name="conditional-access--multifactor-authentication-errors"></a>Erreurs d’accès conditionnel / authentification multifacteur

Dans certaines configurations d’identités dans AAD et Microsoft 365, certaines ressources accessibles avec Microsoft Graph peuvent nécessiter une authentification multifacteur (MFA), même si le client Microsoft 365 n’est pas utilisé. Lorsqu’AAD reçoit une requête pour obtenir un jeton d’accès à la ressource protégée par AMF via le flux « de la part de », il renvoie au service web de votre complément un message JSON contenant une propriété `claims`. La propriété de revendication comporte des informations sur les facteurs d’authentification supplémentaires nécessaires.

Votre code doit tester cette propriété `claims`. En fonction de l’architecture de votre complément, vous pouvez le tester côté client, ou le tester sur le serveur et le relayer sur le client. Il vous faut ces informations dans le client, car Office gère l’authentification des compléments SSO. Si vous le relayez côté serveur, le message adressé au client peut être une erreur (telle que `500 Server Error` ou `401 Unauthorized`) ou se trouver dans le corps d’une réponse de succès (telle que `200 OK`). Dans les deux cas, le rappel (réussite ou échec) de l’appel AJAX de votre code côté client à l’API web de votre complément devra tester cette réponse. 

Quelle que soit votre architecture, si la valeur claims a été envoyée à partir d’AAD, votre code doit rappeler `getAccessToken` et transmettre l’option `authChallenge: CLAIMS-STRING-HERE` dans le `options` paramètre. Lorsqu’AAD voit cette chaîne, il demande le(s) facteur(s) supplémentaire(s) à l’utilisateur, puis renvoie un nouveau jeton d’accès qui sera accepté dans le flux « de la part de ».

### <a name="consent-missing-errors"></a>Erreurs de consentement manquant

Si AAD ne détient aucune trace qu’un consentement (à la ressource Microsoft Graph) a été accordé au complément par l’utilisateur (ou administrateur client), AAD envoie un message d’erreur à votre service web. Votre code doit indiquer au client (dans le corps d’une réponse `403 Forbidden`, par exemple).

Si le complément a besoin d’étendues Microsoft Graph qui ne peuvent être envoyées qu’à un administrateur, votre code doit générer une erreur. Si les seules étendues requises peuvent être envoyées par l’utilisateur, votre code doit basculer vers un autre système d’authentification des utilisateurs.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erreurs d’étendue (autorisation) non valide ou manquante

Ce type d’erreur ne doit apparaître qu’en développement.

- Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client qui doit consigner l’erreur dans la console ou l’enregistrer dans un journal.
- Assurez-vous que la section [Scopes](../reference/manifest/scopes.md) du manifeste de votre complément indique toutes les autorisations nécessaires. Vérifiez également que l’alignement du service web de votre complément spécifie les mêmes autorisations. Vérifiez les fautes d’orthographe. Pour plus d’informations, voir [Inscrire votre complément avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).

### <a name="invalid-audience-error-in-the-access-token-not-the-bootstrap-token"></a>Erreur d’audience non valide dans le jeton d’accès (pas le jeton bootstrap)

Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client, qui doit présenter un message amical à l’utilisateur et éventuellement consigner l’erreur dans la console ou l’enregistrer dans un journal.
