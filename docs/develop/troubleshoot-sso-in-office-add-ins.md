# <a name="troubleshoot-error-messages-for-single-sign-on-sso"></a>Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)

Cet article fournit des conseils sur la résolution des problèmes liés à l’authentification unique (SSO) dans les compléments Office, et explique comment faire en sorte que votre complément gère correctement les conditions particulières ou les erreurs.

## <a name="debugging-tools"></a>Outils de débogage

Lorsque vous développez, nous vous recommandons vivement d’utiliser un outil capable d’intercepter et d’afficher les demandes HTTP du service web de votre complément, ainsi que les réponses. Deux des outils les plus appréciés sont : 

- [Fiddler](http://www.telerik.com/fiddler) : Gratuit ([documentation](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/) : Gratuit pendant 30 jours ([documentation](https://www.charlesproxy.com/documentation/))

Lorsque vous développez votre API de service, vous pouvez également essayer :

- [Postman](http://www.getpostman.com/postman) : Gratuit ([documentation](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>Causes et gestion des erreurs de getAccessTokenAsync

### <a name="13000"></a>13000

L’API [getAccessTokenAsync](http://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync) n’est pas prise en charge par le complément ou la version d’Office. 

- La version d’Office ne prend pas en charge la SSO. La version requise est Office 2016, version 1710, build 8629.nnnn ou ultérieure (la version par abonnement Office 365, parfois appelée « Démarrer en un clic »). Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1). 
- Le complément manifeste n’inclut pas la section [WebApplicationInfo](http://dev.office.com/reference/add-ins/manifest/webapplicationinfo) appropriée.

### <a name="13001"></a>13001

L’utilisateur n’est pas connecté à Office. Votre code doit rappeler la méthode `getAccessTokenAsync` et transmettre l’option `forceAddAccount: true` dans le paramètre [options](../../reference/shared/office.context.auth.getAccessTokenAsync.md#parameters). 

### <a name="13002"></a>13002

L’utilisateur a interrompu sa connexion ou son consentement. 
- Si votre complément fournit des fonctions qui ne nécessitent pas la connexion (ou le consentement) de l’utilisateur, votre code doit intercepter cette erreur et autoriser l’exécution du complément.
- Si le complément requiert un utilisateur connecté ayant accordé son consentement, votre code doit demander à l’utilisateur de répéter l’opération, mais pas plus d’une fois. 

### <a name="13003"></a>13003

Type d’utilisateur non pris en charge. L’utilisateur n’est pas connecté à Office avec un compte Microsoft, Professionnel ou Scolaire valide. Cela peut se produire si Office est exécuté avec un compte de domaine en local, par exemple. Votre code doit demander à l’utilisateur de se connecter à Office.

### <a name="13004"></a>13004

Ressource non valide. Le manifeste du complément n’a pas été configuré correctement. Mettez à jour le manifeste. Pour plus d’informations, consultez la rubrique [Validation et résolution des problèmes avec votre manifeste](troubleshoot-manifest.md).

### <a name="13005"></a>13005

Octroi non valide. Cela signifie généralement qu’Office n’a pas été pré-autorisé sur le service web du complément. Pour plus d’informations, consultez la rubrique sur la [création de l’application de service](../../docs/develop/sso-in-office-add-ins.md#create-the-service-application) et sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](../../docs/develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) ou sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](../../docs/develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (nœud JS). Cela peut également arriver si l’utilisateur n’a pas accordé à votre service les autorisations d’application pour son élément `profile`.

### <a name="13006"></a>13006

Erreur du client. Votre code doit suggérer à l’utilisateur de se déconnecter pour redémarrer Office.

### <a name="13007"></a>13007

L’hôte Office n’a pas pu obtenir de jeton d’accès au service web du complément.
- Assurez-vous que l’enregistrement de votre complément, ainsi que son manifeste, spécifient les autorisations `openid` et `profile`. Pour plus d’informations, consultez la rubrique sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](../../docs/develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) ou sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](../../docs/develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (nœud JS), et sur la [configuration du complément](../../docs/develop/create-sso-office-add-ins-aspnet.md#configure-the-add-in)(ASP.NET) ou sur la [configuration du complément](../../docs/develop/create-sso-office-add-ins-nodejs.md#configure-the-add-in) (nœud JS).
- Votre code peut suggérer à l’utilisateur de réessayer ultérieurement.

### <a name="13008"></a>13008

L’utilisateur a déclenché une opération qui appelle `getAccessTokenAsync` avant d’avoir terminé une opération qui appelle `getAccessTokenAsync`. Votre code doit demander à l’utilisateur de répéter l’opération une fois que l’opération précédente sera terminée.

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Erreurs d’Azure Active Directory côté serveur

### <a name="conditional-access--multifactor-authentication-errors"></a>Erreurs d’accès conditionnel / authentification multifacteur
 
Dans certaines configurations d’identité sur AAD et Office 365, il est possible pour certaines ressources accessibles via Microsoft Graph d’exiger une authentification multifacteur (AMF), même lorsque ce n’est pas le cas de la location Office 365 de l’utilisateur. Lorsqu’AAD reçoit une requête pour obtenir un jeton d’accès à la ressource protégée par AMF via le flux « de la part de », il renvoie au service web de votre complément un message JSON contenant une propriété `claims`. La propriété de revendication comporte des informations sur les facteurs d’authentification supplémentaires nécessaires. 

Votre code côté serveur doit tester ce message et relayer la valeur de revendication à votre code côté client. Il vous faut ces informations dans le client, car Office gère l’authentification des compléments SSO. Le message adressé au client peut être une erreur (telle que `500 Server Error` ou `401 Unauthorized`) ou se trouver dans le corps d’une réponse de succès (telle que `200 OK`). Dans les deux cas, le rappel (réussite ou échec) de l’appel AJAX de votre code côté client à l’API web de votre complément devra tester cette réponse. Si la valeur de revendication a été relayée, votre code doit rappeler `getAccessTokenAsync` et transmettre l’option `authChallenge: CLAIMS-STRING-HERE` dans le paramètre [options](../../reference/shared/office.context.auth.getAccessTokenAsync.md#parameters). Lorsqu’AAD voit cette chaîne, il demande le(s) facteur(s) supplémentaire(s) à l’utilisateur, puis renvoie un nouveau jeton d’accès qui sera accepté dans le flux « de la part de ».

Voici quelques exemples permettant d’illustrer cette gestion AMF : 

- [SSO ASPNET pour complément Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO) : La bibliothèque MSAL utilisée dans cet exemple expose le message AMF de AAD sous la forme d’une exception. Le code transmet ces informations au client sous la forme d’une réponse `500 Server Error`. Dans le script côté client, le rappel `fail` de l’appel AJAX rappelle `getAccessTokenAsync` avec l’option `authChallenge`. Reportez-vous en particulier aux fichiers [ValuesController.cs](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs) et [Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js).
- [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) : Le message AMF de AAD est envoyé au client sous la forme d’une réponse de succès. Dans le script côté client, le rappel `done` de l’appel AJAX rappelle `getAccessTokenAsync` avec l’option `authChallenge`. Reportez-vous en particulier aux fichiers [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) et [program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).

### <a name="consent-missing-errors"></a>Erreurs de consentement manquant

Si AAD ne détient aucune trace qu’un consentement (à la ressource Microsoft Graph) a été accordé au complément par l’utilisateur (ou administrateur client), AAD envoie un message d’erreur à votre service web. Votre code doit indiquer au client (dans le corps d’une réponse `403 Forbidden`, par exemple) qu’il doit rappeler `getAccessTokenAsync` avec l’option `forceConsent: true`.

### <a name="invalid-or-missing-scope-permission-errors"></a>Erreurs d’étendue (permission) non valide ou manquante

- Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client, qui doit présenter un message amical à l’utilisateur. Si possible, consignez l’erreur dans la console ou enregistrez-la dans un journal.
- Assurez-vous que la section [Scopes](http://dev.office.com/reference/add-ins/manifest/scopes) du manifeste de votre complément indique toutes les autorisations nécessaires. Vérifiez également que l’alignement du service web de votre complément spécifie les mêmes autorisations. Vérifiez les fautes d’orthographe. Pour plus d’informations, consultez la rubrique sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](../../docs/develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) ou sur l’[enregistrement du complément avec le point de terminaison Azure AD v2.0](../../docs/develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (nœud JS), et sur la [configuration du complément](../../docs/develop/create-sso-office-add-ins-aspnet.md#configure-the-add-in)(ASP.NET) ou sur la [configuration du complément](../../docs/develop/create-sso-office-add-ins-nodejs.md#configure-the-add-in) (nœud JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Erreurs de jetons expirés ou invalides lors de l’appel à Microsoft Graph

Certaines bibliothèques d’autorisation et d’authentification, y compris MSAL, évitent les erreurs de jetons expirés grâce à un jeton d’actualisation mis en cache. Vous pouvez également coder votre propre système de mise en cache de jeton. Pour un exemple, consultez la rubrique [SSO NodeJS pour complément Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), notamment le fichier [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Cependant, si vous recevez un message d’erreur pour jeton expiré ou invalide, votre code doit demander au client (dans le corps d’une réponse `401 Unauthorized`, par exemple) de rappeler `getAccessTokenAsync` et répéter l’appel vers le point de terminaison de l’API web de votre complément, qui répétera le flux « de la part de » afin d’obtenir un nouveau jeton pour Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Erreur de jeton non valide lors de l’appel à Microsoft Graph

Gérez cette erreur de la même manière qu’une erreur de jeton expiré. Reportez-vous à la section précédente.

### <a name="invalid-audience-error"></a>Erreur de public non valide

Votre code côté serveur doit envoyer une réponse `403 Forbidden` au client, qui doit présenter un message amical à l’utilisateur et éventuellement consigner l’erreur dans la console ou l’enregistrer dans un journal.

Pour plus d’informations sur l’ajout de prise en charge multi-locataire pour la validation de jeton, consultez la rubrique [Exemple de multi-locataire Azure](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
