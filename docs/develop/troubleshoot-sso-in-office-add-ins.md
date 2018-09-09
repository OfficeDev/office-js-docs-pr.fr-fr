---
title: Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: ef4d7ed873121deec5fd235e0eace70a3a0c2f0e
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876598"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>Résolution des messages d’erreur pour l’authentification unique (SSO) (aperçu)

Cet article fournit des conseils sur la résolution des problèmes liés à l’authentification unique (SSO) dans les compléments Office, et explique comment faire en sorte que votre complément gère correctement les conditions particulières ou les erreurs.

> [!NOTE]
> L’API d'authentification unique est actuellement prise en charge en mode aperçu pour Word, Excel, Outlook et PowerPoint. Pour plus d’informations à propos de la prise en charge actuelle de l’API d’authentification unique, consultez la rubrique [Ensembles de conditions requises d'IdentityAPI]https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets).
> Pour utiliser l’authentification unique, vous devez charger la version bêta de la bibliothèque JavaScript Office à partir de https://appsforoffice.microsoft.com/lib/beta/hosted/office.js dans la page de démarrage HTML du complément.
> Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="debugging-tools"></a>Outils de débogage

Lors du développement, nous vous recommandons vivement d’utiliser un outil capable d’intercepter et d’afficher les demandes HTTP du service web de votre complément, ainsi que les réponses. Les deux outils les plus populaires sont les suivants : 

- [Fiddler](http://www.telerik.com/fiddler) : Gratuit ([documentation](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): Gratuit pendant 30 jours. ([Documentation](https://www.charlesproxy.com/documentation/))

Lorsque vous développez votre API de service, vous pouvez également essayer :

- [Postman](http://www.getpostman.com/postman) : Gratuit ([documentation](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>Causes et gestion des erreurs de getAccessTokenAsync

Pour consulter des exemples de la gestion des erreurs décrite dans cette section, reportez-vous aux articles suivants :
- [Home.js dans Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [program.js dans Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> Outre les suggestions faites dans cette section, un complément Outlook a un moyen supplémentaire pour répondre à n'importe quelle erreur 13*nnn*. Pour plus d’informations, reportez-vous à [Scénario : implémenter l’authentification unique sur votre service dans un complément Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) et [Exemple de complément AttachmentsDemo](https://github.com/OfficeDev/outlook-add-in-attachments-demo). 

### <a name="13000"></a>13000

L’API [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) n’est pas prise en charge par le complément ou la version d’Office. 

- La version d’office ne prend pas en charge SSO. La version requise est Office 2016, version 1710, build 8629.nnnn ou version ultérieure (la version par abonnement Office 365, parfois appelée « Démarrer en un clic »). Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à la rubrique [Participer au programme Office Insider](https://products.office.com/office-insider?tab=tab-1). 
- Le manifeste de complément n’inclut pas la section [WebApplicationInfo](https://dev.office.com/reference/add-ins/manifest/webapplicationinfo) appropriée.

### <a name="13001"></a>13001

L’utilisateur n’est pas connecté à Office. Votre code doit rappeler la `getAccessTokenAsync` méthode et passer l’option `forceAddAccount: true` dans le paramètre [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Mais ne le faites pas plus d'une fois. L'utilisateur peut avoir décidé de ne pas se connecter.

Cette erreur n’est jamais apparue dans Office Online. Si le cookie de l’utilisateur a expiré, Office Online renvoie l’erreur 13006. 

### <a name="13002"></a>13002

L'utilisateur a abandonné la connexion ou le consentement ; par exemple, en choisissant **Annuler** dans le dialogue de consentement. 
- Si votre complément fournit des fonctions qui ne nécessitent pas la connexion (ou le consentement) de l’utilisateur, votre code doit intercepter cette erreur et autoriser l’exécution du complément.
- Si le complément requiert un utilisateur connecté ayant accordé son consentement, votre code doit demander à l’utilisateur de répéter l’opération, mais pas plus d’une fois. 

### <a name="13003"></a>13003

Type d’utilisateur non pris en charge. L’utilisateur n’est pas connecté à Office avec un compte Microsoft, Professionnel ou Scolaire valide. Cela peut se produire si Office est exécuté avec un compte de domaine en local, par exemple. Votre code doit demander à l’utilisateur de se connecter à Office.

### <a name="13004"></a>13004

Ressource non valide. Le manifeste du complément n’a pas été configuré correctement. Mettez à jour le manifeste. Pour plus d’informations, consultez la rubrique [Validation et résolution des problèmes avec votre manifeste](../testing/troubleshoot-manifest.md). Le problème le plus courant est que l'élément **Resource** (dans l'élément **WebApplicationInfo**) a un domaine qui ne correspond pas au domaine du complément. Bien que la partie protocole de la valeur Resource doive être "api" et non "https", toutes les autres parties du nom de domaine (y compris le port, s'il existe) doivent être les mêmes que pour le complément.

### <a name="13005"></a>13005

Octroi non valide. Cela signifie généralement qu’Office n’a pas été pré-autorisé sur le service web du complément. Pour plus d’informations, consultez la rubrique sur la [création de l’application de service](sso-in-office-add-ins.md#create-the-service-application) et sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (nœud JS). Cela peut également arriver si l’utilisateur n’a pas accordé à votre service les autorisations d’application à leur `profile`.

### <a name="13006"></a>13006

Erreur client. Votre code doit suggérer à l’utilisateur de se déconnecter et de redémarrer Office, ou de redémarrer la session Office Online.

### <a name="13007"></a>13007

L’hôte Office n’a pas pu obtenir de jeton d’accès au service Web du complément.
- Si cette erreur se produit pendant le développement, assurez-vous que votre enregistrement de complément et votre manifeste de complément spécifient les autorisations `openid` et `profile`. Pour plus d’informations, consultez la rubrique sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (nœud JS), et sur la [configuration du complément](create-sso-office-add-ins-aspnet.md#configure-the-add-in)(ASP.NET) ou sur la [configuration du complément](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (nœud JS).
- En production, il y a plusieurs choses qui peuvent provoquer cette erreur. En voici certaines :
    - L'utilisateur a révoqué son consentement, après l'avoir préalablement accordé. Votre code doit rappeler la méthode `getAccessTokenAsync` avec l’option `forceConsent: true`, mais pas plus d’une fois.
    - L'utilisateur possède une identité de compte Microsoft (MSA). Certaines situations qui entraîneraient l’une des autres erreurs 13nnn avec un compte Work ou School, provoqueront une erreur 13007 lorsqu’un MSA est utilisé. 

  Pour tous ces cas, si vous avez déjà essayé l'option `forceConsent` une fois, votre code pourrait suggérer que l'utilisateur réessaye l'opération plus tard.

### <a name="13008"></a>13008

L’utilisateur a déclenché une opération qui appelle `getAccessTokenAsync` avant d’avoir terminé une opération qui appelle `getAccessTokenAsync`. Votre code doit demander à l’utilisateur de répéter l’opération une fois que l’opération précédente sera terminée.

### <a name="13009"></a>13009

Le complément a appelé la méthode `getAccessTokenAsync` avec l’option `forceConsent: true`, mais le manifeste du complément est déployé sur un type de catalogue qui ne prend pas en charge le consentement forcé. Votre code doit rappeler la `getAccessTokenAsync` méthode et passer l’option `forceConsent: false` dans le paramètre [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Toutefois, l’appel de `getAccessTokenAsync` avec `forceConsent: true` peut lui-même représenter une réponse automatique à un appel ayant échoué de `getAccessTokenAsync` avec `forceConsent: false`, donc votre code doit suivre si `getAccessTokenAsync` avec `forceConsent: false` a déjà été appelé. Si c’est le cas, votre code doit indiquer à l’utilisateur de se déconnecter d’Office et de se connecter à nouveau.

> [!NOTE]
> Microsoft n’imposera pas nécessairement cette restriction sur tous les types de catalogues de compléments. Si la restriction ne s’applique pas, l’erreur n’est jamais affichée.

### <a name="13010"></a>13010

L’utilisateur exécute le complément sur Office Online et utilise Edge ou Internet Explorer. Le domaine Office 365 de l’utilisateur et le domaine login.microsoftonline.com sont dans des zones de sécurité distinctes dans les paramètres de navigateur. Si cette erreur est renvoyée, l’utilisateur a déjà vu une erreur expliquant cela et menant vers une page sur la modification de la configuration de la zone. Si votre complément fournit des fonctions qui ne nécessitent pas que l’utilisateur soit connecté, votre code doit intercepter cette erreur et autoriser l’exécution du complément.

### <a name="50001"></a>50001

Cette erreur (qui n'est pas spécifique à `getAccessTokenAsync`) peut indiquer que le navigateur a en cache une ancienne copie des fichiers office.js. Effacez le cache du navigateur. Une autre possibilité est que la version d'Office n'est pas suffisamment récente pour prendre en charge SSO. Voir [Conditions préalables](create-sso-office-add-ins-aspnet.md#prerequisites).

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erreurs d’Azure Active Directory côté serveur

Pour plus d’exemples de la gestion des erreurs décrite dans cette section, reportez-vous aux articles suivants :
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>Erreurs d’accès conditionnel / authentification multifacteur
 
Dans certaines configurations d’identité sur AAD et Office 365, il est possible pour certaines ressources accessibles via Microsoft Graph d’exiger une authentification multifacteur (AMF), même lorsque ce n’est pas le cas de la location Office 365 de l’utilisateur. Lorsqu’AAD reçoit une requête pour obtenir un jeton d’accès à la ressource protégée par AMF via le flux « de la part de », il renvoie au service web de votre complément un message JSON contenant une propriété `claims`. La propriété de revendication comporte des informations sur les facteurs d’authentification supplémentaires nécessaires. 

Votre code côté serveur doit tester ce message et relayer la valeur de revendication à votre code côté client. Il vous faut ces informations dans le client, car Office gère l’authentification des compléments SSO. Le message adressé au client peut être une erreur (telle que `500 Server Error` ou `401 Unauthorized`) ou se trouver dans le corps d’une réponse de succès (telle que `200 OK`). Dans les deux cas, le rappel (réussite ou échec) de l’appel AJAX de votre code côté client à l’API web de votre complément devra tester cette réponse. Si la valeur de revendication a été relayée, votre code doit rappeler `getAccessTokenAsync` et transmettre l’option `authChallenge: CLAIMS-STRING-HERE` dans le paramètre [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Lorsqu’AAD voit cette chaîne, il demande le(s) facteur(s) supplémentaire(s) à l’utilisateur, puis renvoie un nouveau jeton d’accès qui sera accepté dans le flux « de la part de ».

### <a name="consent-missing-errors"></a>Erreurs de consentement manquant

Si AAD ne détient aucune trace qu’un consentement (à la ressource Microsoft Graph) a été accordé au complément par l’utilisateur (ou administrateur client), AAD envoie un message d’erreur à votre service web. Votre code doit indiquer au client (dans le corps d’une réponse `403 Forbidden`, par exemple) qu’il doit rappeler `getAccessTokenAsync` avec l’option `forceConsent: true`.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erreurs d’étendue (permission) non valide ou manquante

- Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client, qui doit présenter un message amical à l’utilisateur. Si possible, consignez l’erreur dans la console ou enregistrez-la dans un journal.
- Assurez-vous que la section [Scopes](https://dev.office.com/reference/add-ins/manifest/scopes) du manifeste de votre complément indique toutes les autorisations nécessaires. Vérifiez également que l’alignement du service web de votre complément spécifie les mêmes autorisations. Vérifiez les fautes d’orthographe. Pour plus d’informations, consultez la rubrique sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) ou sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (nœud JS), et sur la [configuration du complément](create-sso-office-add-ins-aspnet.md#configure-the-add-in)(ASP.NET) ou sur la [configuration du complément](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (nœud JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Erreurs de jetons expirés ou invalides lors de l’appel à Microsoft Graph

Certaines bibliothèques d’autorisation et d’authentification, y compris MSAL, évitent les erreurs de jetons expirés grâce à un jeton d’actualisation mis en cache. Vous pouvez également coder votre propre système de mise en cache de jeton. Pour un exemple, consultez la rubrique [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), notamment le fichier [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Cependant, si vous recevez un message d’erreur pour jeton expiré ou invalide, votre code doit demander au client (dans le corps d’une réponse `401 Unauthorized`, par exemple) de rappeler `getAccessTokenAsync` et répéter l’appel vers le point de terminaison de l’API web de votre complément, qui répétera le flux « de la part de » afin d’obtenir un nouveau jeton pour Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Erreur de jeton non valide lors de l’appel à Microsoft Graph

Gérez cette erreur de la même manière qu’une erreur de jeton expiré. Consultez la section précédente.

### <a name="invalid-audience-error"></a>Erreur de public non valide

Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client, qui doit présenter un message amical à l’utilisateur et éventuellement consigner l’erreur dans la console ou l’enregistrer dans un journal.

Pour plus d’informations sur l’ajout de prise en charge multi-locataire pour la validation de jeton, consultez la rubrique [Exemple de multi-locataire Azure](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
