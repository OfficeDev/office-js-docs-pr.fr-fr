# <a name="enable-single-sign-on-for-office-add-ins"></a>Activer l’authentification unique pour des compléments Office

Les utilisateurs se connectent à Office (plateformes en ligne, mobiles et de bureau) à l’aide de leur compte Microsoft personnel ou de leur compte professionnel ou scolaire (Office 365). Vous pouvez en profiter et vous servir de la SSO comme suit, sans que l’utilisateur ait besoin de se connecter une deuxième fois :

* Autorisez l’utilisateur à se connecter dans votre complément.
* Autorisez le complément à accéder à [Microsoft Graph](https://developer.microsoft.com/graph/docs).

![Image illustrant le processus de connexion pour un complément](../images/OfficeHostTitleBarLogin.png)

>**Remarque :** L’API de l’authentification unique est actuellement prise en charge pour Word, Excel et PowerPoint. Pour plus d’informations sur l’endroit où l’API de l’authentification unique est actuellement prise en charge, consultez la rubrique [Ensembles de conditions requises de l’API d’identité](http://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).
> L’authentification unique est actuellement en préversion pour Outlook. Si vous utilisez un complément Outlook, veillez à activer l’authentification moderne pour la location d’Office 365. Pour plus d’informations sur la manière de procéder, consultez la rubrique [Exchange Online : Activation de votre client pour l’authentification moderne](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Pour les utilisateurs, cela permet une exécution aisée de votre complément qui ne requiert qu’une seule connexion. Pour les développeurs, cela signifie que l’utilisation de votre complément permet d’authentifier les utilisateurs et d’obtenir un accès autorisé aux données de l’utilisateur via Microsoft Graph avec les informations d’identification que l’utilisateur a déjà fournies à l’application Office.

## <a name="sso-add-in-architecture"></a>Architecture des compléments d’authentification unique

Outre l’hébergement des pages et du JavaScript de l’application Web, le complément doit également héberger, dans le même [nom de domaine complet](https://msdn.microsoft.com/en-us/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly), une ou plusieurs API Web qui recevront un jeton d’accès à Microsoft Graph et effectueront des requêtes.

Le manifeste du complément contient un balisage qui spécifie comment le complément est enregistré dans le point de terminaison Azure Active Directory (Azure AD) v2.0 et il indique les autorisations à Microsoft Graph dont le complément a besoin.

### <a name="how-it-works-at-runtime"></a>Mode de fonctionnement en cours d’exécution

Le diagramme suivant illustre le mode de fonctionnement du processus d’authentification unique.
<!-- Minor fixes to the text in the diagram - change V2 to v2.0, and change "(e.g. Word, Excel, etc.)" to "(for example, Word, Excel)". -->
![Diagramme illustrant le processus d’authentification unique](../images/SSOOverviewDiagram.png)

1. Dans le complément, JavaScript appelle une nouvelle API Office.js `getAccessTokenAsync`. Cela indique à l’application hôte Office qu’elle doit obtenir un jeton d’accès au complément. (Ci-après, ce jeton est également appelé **« jeton de complément »**.)
1. Si l’utilisateur n’est pas connecté, l’application hôte Office ouvre une fenêtre contextuelle pour que l’utilisateur se connecte.
1.  Si c’est la première fois que l’utilisateur actuel utilise votre complément, il est invité à donner son consentement.
1. L’application hôte Office demande le **jeton de complément** au point de terminaison Azure AD v2.0 pour l’utilisateur actuel.
1. Azure AD envoie le jeton de complément à l’application hôte Office.
1. L’application hôte Office envoie le **jeton de complément** au complément dans le cadre de l’objet de résultat renvoyé par l’appel `getAccessTokenAsync`.
1. Un code JavaScript dans le complément effectue une requête HTTP à une API web qui est hébergée sur le même domaine complet que le complément et inclut le **jeton de complément** comme preuve d’autorisation.  
1. Le code côté serveur valide le **jeton de complément** entrant.
1. Le code côté serveur utilise le flux « de la part de » (défini dans [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) et l’application de [démon ou de serveur dans un scénario Azure avec une API web](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)) pour obtenir un jeton d’accès à Microsoft Graph (ci-après, le **« jeton MSG »**) en échange du jeton de complément.
1. Azure AD renvoie le **jeton MSG** (et un jeton d’actualisation si le complément demande l’autorisation *offline_access*) au complément.
1. Le code côté serveur met en cache le(s) **jeton(s) MSG**.
1. Le code côté serveur effectue des requêtes à Microsoft Graph et inclut le **jeton MSG**.
1. Microsoft Graph renvoie des données au complément, qui peut les transmettre à l’interface utilisateur du complément.
1. Lorsque le jeton MSG arrive à expiration, le code côté serveur peut utiliser son jeton d’actualisation pour obtenir un nouveau **jeton MSG**.

## <a name="develop-an-sso-add-in"></a>Développer un complément d’authentification unique

Cette section décrit les tâches impliquées dans la création d’un complément Office qui utilise l’authentification unique. Ces tâches sont décrites ici indépendamment du langage et de l’infrastructure. Pour obtenir des exemples de procédures pas à pas détaillées, consultez les rubriques suivantes :

* [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md)
* [Créer un complément Office ASP.NET qui utilise l’authentification unique](../develop/create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Créer l’application de service

Enregistrez le complément sur le portail d’inscription pour le point de terminaison Azure v2.0 : https://apps.dev.microsoft.com. Il s’agit d’un processus de 5 à 10 minutes qui inclut les tâches suivantes :

* Obtenez un ID de client et un code secret pour le complément.
* Spécifiez les autorisations dont votre complément a besoin pour Microsoft Graph.
* Accordez la confiance de l’application hôte Office au complément.
* Pré-autorisez l’application hôte Office pour le complément avec l’autorisation par défaut *access_as_user*.

### <a name="configure-the-add-in"></a>Configurer le complément

Ajoutez un nouveau balisage au manifeste du complément :

* **WebApplicationInfo** : parent des éléments suivants.
* **ID** : ID client du complément.
* **Resource** : URL du complément.
* **Scopes** : parent d’un ou plusieurs éléments **Scope**.
* **Scope** : spécifie une autorisation nécessaire pour le complément dans Microsoft Graph. Par exemple : `User.Read`, `Mail.Read` ou `offline_access`). Pour plus d’informations, reportez-vous à l’article relatif aux [Autorisations Microsoft Graph](https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference).

Pour les hôtes Office autres qu’Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

### <a name="add-client-side-code"></a>Ajouter du code côté client

Ajoutez un code JavaScript pour le complément à :

* Appel `Office.context.auth.getAccessTokenAsync(myTokenHandler)`.
* Créez un gestionnaire qui transmet le jeton de complément au code côté serveur du complément. Par exemple :

```js
function mytokenHandler(asyncResult) {
    // Passes asyncResult.value (which has the add-in access token)
    // to the add-in’s web API as an Authorization header.
}
```

### <a name="when-to-call-the-method"></a>Quand appeler la méthode

Si votre complément ne peut pas être utilisé lorsqu’un non-utilisateur est connecté à Office et qu’Office ne possède pas de jeton d’accès à votre complément, vous devez appeler `getAccessTokenAsync` * au lancement du complément*.

Si le complément possède certaines fonctionnalités qui ne nécessitent pas un accès à Microsoft Graph ni même un utilisateur connecté, appelez `getAccessTokenAsync` * lorsque l’utilisateur effectue une action qui requiert l’accès à Microsoft Graph ou, au moins, un utilisateur connecté*. Les appels répétés à `getAccessTokenAsync` ne causent aucune dégradation importante des performances, car Office met en cache le jeton d’accès et le réutilise jusqu'à ce qu’il arrive à expiration, sans effectuer un autre appel à l’AAD V. Point de terminaison 2.0 dès que `getAccessTokenAsync` est appelé. Ainsi, vous pouvez ajouter des appels de `getAccessTokenAsync` à l’ensemble des fonctions et gestionnaires qui lancent une action dans laquelle le jeton est nécessaire.

### <a name="add-server-side-code"></a>Ajouter du code côté serveur

Créez une ou plusieurs méthodes API Web qui obtiennent des données Microsoft Graph. Selon votre langue et votre infrastructure, des bibliothèques peuvent être disponibles pour simplifier le code à rédiger. Votre code côté serveur doit effectuer les opérations suivantes :

* Valider le jeton de complément reçu à partir du gestionnaire de jetons que vous avez créé précédemment.
* Démarrer le flux « de la part de » avec un appel du point de terminaison Azure AD v2.0 qui inclut le jeton d’accès du complément, certaines métadonnées relatives à l’utilisateur et les informations d’identification du complément (ID et code secret).
* Mettre en cache le jeton MSG renvoyé.
* Obtenir des données à partir de Microsoft Graph en utilisant le jeton MSG.
