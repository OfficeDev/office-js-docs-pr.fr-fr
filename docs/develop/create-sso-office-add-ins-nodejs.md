# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Création d’un complément Office Node.js qui utilise l’authentification unique

Les utilisateurs peuvent se connecter à Office et votre complément Web Office peut tirer parti de cette procédure de connexion pour autoriser les utilisateurs de votre complément et de Microsoft Graph sans obliger les utilisateurs à une deuxième authentification. Pour obtenir une vue d’ensemble, voir [Activer l’authentification unique dans un complément Office](../develop/sso-in-office-add-ins.md).

Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré avec Node.js et express. 

> **Remarque :** Pour un article similaire concernant un complément basé sur ASP.NET, voir [Créer un complément Office ASP.NET qui utilise l’authentification unique](../develop/create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Conditions préalables

* [Nœud et npm](https://nodejs.org/en/), version 6.9.4 ou ultérieure.
* [GIT Bash](https://git-scm.com/downloads) (ou un autre client Git)
* TypeScript version 2.2.2 ou ultérieure.
* Office 2016, Version 1708, build 8424.nnnn ou version ultérieure (la version par abonnement Office 365, parfois appelée « Démarrer en un clic »). Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, voir [Participez au programme Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

1. Clonez ou téléchargez le référentiel sur [Complément Office NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso). 


    > **Remarque :** Il existe deux versions de l’échantillon. 
    > 
    > * Le dossier **Before** est un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. Les sections suivantes de cet article vous guident tout au long de la procédure d’exécution de cette dernière. 
    > * La version **Finale** de l’échantillon s’apparente au complément que vous auriez si vous terminiez les procédures de cet article, sauf que le projet terminé comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, suivez simplement les instructions de cet article, mais remplacez « Avant » par « Finale » et ignorez les sections **Code côté client** et **Code côté serveur**.

1. Ouvrez une console Git Bash dans le dossier **Before**.

2. Saisissez `npm install` dans la console pour installer toutes les dépendances détaillées dans le fichier package.json.

3. Saisissez `npm run build ` dans la console pour générer le projet. 
     > Remarque : Vous pouvez voir certaines erreurs de construction indiquant que certaines variables sont déclarées mais pas utilisées. Ignorez ces erreurs. Elles représentent un effet secondaire du fait qu’il manque du code dans la version « Avant » de l’échantillon, qui sera ajouté ultérieurement.

## <a name="register-the-add-in-with-azure-ad-v2-endpoint"></a>Enregistrer le complément avec le point de terminaison Azure AD V2

1. Accédez à [https://apps.dev.microsoft.com](https://apps.dev.microsoft.com). 

1. Connectez-vous avec les informations d’identification d’administrateur à votre client Office 365. Par exemple, MonNom@contoso.onmicrosoft.com

1. Cliquez sur **Ajouter une application**.

1. Lorsque vous y êtes invité, utilisez « Office-Add-in-NodeJS-SSO » comme nom d’application et appuyez sur **Créer une application**.

1. Quand la page de configuration de l’application s’ouvre, copiez l’**ID de l’application** et enregistrez-le. Vous l’utiliserez dans une procédure ultérieure. 

    > Remarque : Cet ID est la valeur « audience » lorsque d’autres applications, telles que l’application hôte Office (par exemple, PowerPoint, Word, Excel) recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Dans la section **Secrets de l’application**, appuyez sur **Générer un nouveau mot de passe**. Une boîte de dialogue contextuelle s’ouvre avec un nouveau mot de passe (également appelé « secret de l’application »). *Copiez le mot de passe immédiatement et enregistrez-le avec l’ID de l’application.* Vous en aurez besoin dans une procédure ultérieure. Ensuite, fermez la boîte de dialogue.

1. Dans la section **Plateformes**, cliquez sur **Ajouter une plateforme**. 

1. Dans la boîte de dialogue qui s’ouvre, sélectionnez **API Web**.

1. Un **URI d’ID d’application** a été généré sous la forme « api://{GUID de l’ID d’application} ». Insérez la chaîne « localhost:3000 » entre les deux barres obliques et le GUID. L’ID entier doit se présenter sous la forme `api://localhost:3000/{App ID GUID}`. (La partie domaine du nom d’**étendue**, juste en dessous de l’**URI d’ID d’application** change automatiquement en conséquence. Il doit se présenter sous la forme `api://localhost:3000/{App ID GUID}/access_as_user`.)

1. Cette étape et la suivante permettent à l’application hôte Office d’accéder à votre complément. Dans la section **Applications pré-autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé. Chaque fois que vous en entrez un, une nouvelle zone de texte vide s’affiche. (Entrez uniquement le GUID.)

 * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
 * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
 * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online) 

1. Ouvrez le menu déroulant **Scope** à côté de chaque **ID d’application** et activez la case à cocher pour `api://localhost:44355/{App ID GUID}/access_as_user`.

1. En haut de la section **Plateformes**, cliquez sur **Ajouter une plateforme** à nouveau, puis sélectionnez **Web**.

1. Dans la nouvelle section **Web** sous **Plateformes**, entrez les informations suivantes en guise d’**URL de redirection** : `https://localhost:3000`. 

    > Remarque : À ce jour, la plateforme **API Web** disparaît parfois de la section **Plateformes**, tout particulièrement si la page est actualisée après l’ajout de la plateforme **Web** *et l’enregistrement de la page d’inscription*. Pour être sûr que votre plateforme **API Web** fait toujours partie de l’inscription, cliquez sur le bouton **Modifier le manifeste de l’application** près du bas de la page. Vous devriez voir la chaîne `api://localhost:3000/{App ID GUID}` dans la propriété **identifierUris** du manifeste. Il devrait également y avoir une propriété **oauth2Permissions** dont la propriété secondaire **value** a la valeur `access_as_user`.

1. Faites défiler jusqu’à la section **Autorisations pour Microsoft Graph** et à la sous-section **Autorisations déléguées**. Utilisez le bouton **Ajouter** pour ouvrir une boîte de dialogue **Sélectionner des autorisations**.

1. Dans la boîte de dialogue, cochez les cases pour les autorisations suivantes : 
    * Files.Read.All
    * profil

1. Cliquez sur **OK** au bas de la boîte de dialogue.

1. Cliquez sur **Enregistrer** au bas de la page d’inscription.

## <a name="grant-admin-consent-to-the-add-in"></a>Accorder le consentement de l’administrateur au complément

> **Remarque :** Cette procédure n’est nécessaire que quand vous êtes en train de développer le complément. Lorsque votre complément de production est déployé dans l’Office Store ou dans un catalogue de compléments, les utilisateurs l’approuvent individuellement à l’installation.

1. Dans la chaîne suivante, remplacez l’espace réservé « {application_ID} » par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément.

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Collez l’URL résultante dans la barre d’adresses d’un navigateur pour y accéder.

1. Lorsque vous y êtes invité, connectez-vous avec les informations d’identification d’administrateur à votre client Office 365.

1. Vous êtes ensuite invité à accorder des autorisations pour votre complément pour accéder à vos données Microsoft Graph. Cliquez sur **Accepter**. 

1. L’onglet ou la fenêtre du navigateur est alors redirigé vers l’**URL de redirection** que vous avez spécifiée lors de l’enregistrement du complément ; ainsi, si le complément fonctionne, la page d’accueil du complément s’ouvre dans le navigateur. Si le complément ne fonctionne pas, vous obtiendrez une erreur indiquant que la ressource au niveau de localhost:3000 ne peut pas être trouvée ou ouverte. *Mais le fait que la redirection ait été tentée signifie que le processus de consentement de l’administrateur a abouti*. Ainsi, que la page d’accueil soit ouverte ou que vous receviez un message d’erreur, vous pouvez passer à l’étape suivante.

2. Dans la barre d’adresses du navigateur, vous verrez un paramètre de requête « client » avec une valeur GUID. Il s’agit de l’ID de votre client Office 365. Copiez et enregistrez cette valeur. Vous l’utiliserez dans une étape ultérieure.

3. Fermez la fenêtre/l’onglet.

## <a name="configure-the-add-in"></a>Configurer le complément

1. Dans votre éditeur de code, ouvrez le fichier src\server.ts. Près de la partie supérieure se trouve un appel à un constructeur d’une classe `AuthModule`. Il existe certains paramètres de chaîne dans le constructeur auxquels vous devez affecter des valeurs.

2. Pour la propriété `client_id`, remplacez l’espace réservé `{client GUID}` par l’ID d’application que vous avez enregistré lorsque vous avez inscrit le complément. Lorsque vous avez terminé, vous obtenez simplement un GUID entre guillemets simples. Il ne doit pas exister de caractères « {} ».

3. Pour la propriété `client_secret`, remplacez l’espace réservé `{client secret}` par le secret de l’application que vous avez enregistré lorsque vous avez inscrit le complément.

4. Pour la propriété `audience`, remplacez l’espace réservé `{audience GUID}` par l’ID d’application que vous avez enregistré lorsque vous avez inscrit le complément. (La même valeur que celle affectée à la propriété `client_id`.)
  
3. Dans la chaîne affectée à la propriété `issuer`, vous verrez l’espace réservé *{O365 tenant GUID}*. Remplacez ceci par l’ID du client Office 365 que vous avez enregistré à la fin de la dernière procédure. Si pour une raison quelconque, vous n’avez pas obtenu l’ID antérieur, utilisez l’une des méthodes de la page [Trouver votre ID de client Office 365](https://support.office.com/fr-fr/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) pour l’obtenir. Lorsque vous avez terminé, la valeur de propriété `issuer` doit ressembler à ceci :

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. Conservez les autres paramètres du constructeur `AuthModule` inchangés. Enregistrez et fermez le fichier.

1. Dans la racine du projet, ouvrez le fichier manifeste du complément « Office-Add-in-NodeJS-SSO.xml ».

1. Faites défiler vers le bas du fichier.

1. Juste au-dessus de la balise de fin `</VersionOverrides>`, vous trouverez le balisage suivant :

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Remplacez l’espace réservé « {application_GUID here} » *aux deux endroits* du balisage par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément. (Les « {} » ne font pas partie de l’ID ; vous ne devez pas les inclure.) C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.

    >Remarque : 
    >
    >* La valeur **Resource** correspond à l’**URI d’ID d’application** défini lorsque vous avez ajouté la plateforme d’API web à l’enregistrement du complément.
    >* La section **Scopes** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via l’Office Store.

1. Enregistrez et fermez le fichier.

## <a name="code-the-client-side"></a>Code côté client

1. Ouvrez le fichier program.js dans le dossier **public**. Il contient déjà du code :

    * Une affectation à la méthode `Office.initialize` qui affecte elle-même un gestionnaire à l’événement ClickButton `getGraphAccessTokenButton`.
    * Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.

1. En dessous de l’affectation au `Office.initialize`, ajoutez le code ci-dessous. Tenez compte des informations suivantes : 

    * La fonction `getDataWithoutAuthChallenge` est appelée lors d’une première tentative d’utilisation du flux « de la part de ». Il est supposé que l’authentification par un facteur unique est suffisante. Vous ajouterez du code dans une étape ultérieure pour gérer le cas où plusieurs facteurs d’authentification seraient nécessaires.
    * `getAccessTokenAsync` est la nouvelle API d’Office.js qui permet à un complément de demander à l’application hôte Office (Excel, PowerPoint, Word, etc.) un jeton d’accès au complément (pour l’utilisateur connecté à Office). L’application hôte Office demande alors le jeton au point de terminaison Azure AD 2.0. Dans la mesure où vous avez préalablement autorisé l’hôte Office sur votre complément lors de son inscription, Azure AD enverra le jeton. 
     * Si aucun utilisateur n’est connecté à Office, l’hôte Office invite l’utilisateur à se connecter. 
     * Le paramètre options définit `forceConsent` sur false, donc l’utilisateur ne sera pas invité à accorder l’accès de l’hôte Office à votre complément.

    ```js
    function getOneDriveItems() {
        getDataWithoutAuthChallenge();
    }   
    
    function getDataWithoutAuthChallenge() {       
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    // TODO1: Use the access token to get Microsoft Graph data.
                }
                else {
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }
    ```

1. Remplacez TODO1 par les lignes suivantes. Vous créez la méthode `getData` et la route « /api/onedriveitems » côté serveur dans les étapes suivantes. Une URL relative est utilisée pour le point de terminaison car il doit être hébergé sur le même domaine que votre complément.

    ```
    accessToken = result.value;
    getData("/api/onedriveitems", accessToken);
    ```

1. En dessous de la méthode `getOneDriveFiles`, ajoutez le code suivant. Cette méthode utilitaire appelle un point de terminaison API Web spécifié et lui transmet le jeton d’accès que l’application hôte Office a utilisé pour accéder à votre complément. Sur le côté serveur, ce jeton d’accès est utilisé dans le flux « de la part de » pour obtenir un jeton d’accès à Microsoft Graph. 

    ```
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            TODO2: Display data and handle demand for multi-factor authentication.
        })
        .fail(function (result) {
            console.log(result.error);
       });
    }
    ```

1. Remplacez TODO2 par le code suivant. À propos de ce code, notez que :
    * Si la cible de Microsoft Graph demande un ou plusieurs facteurs d’authentification supplémentaires, le résultat ne contiendra pas de données. Le résultat sera un JSON de revendication indiquant à l’AAD quels facteurs supplémentaires l’utilisateur doit fournir. Dans ce cas, le client doit effectuer une nouvelle authentification qui transmettra cette chaîne de revendications à l’AAD afin que ce dernier fournisse les invites nécessaires.
    * Si le résultat est un JSON de revendication, il contiendra la chaîne « capolids ».
    * Vous allez créer la fonction `getDataUsingAuthChallenge` dans une étape ultérieure.

    ```
    if (result[0].indexOf('capolids') !== -1) {                
        result[0] = JSON.parse(result[0])
        getDataUsingAuthChallenge(result[0]);
    } else {  
        showResult(result);
    }
    ```

1. Ajoutez la fonction suivante au fichier juste en dessous de la fonction `getData`. À propos de cette fonction, vous remarquerez :
    * La fonction est utilisée lorsque l’AAD requiert un ou plusieurs facteurs d’authentification supplémentaires. 
    * La fonction déclenche une deuxième authentification, dans laquelle l’utilisateur est invité à fournir un ou plusieurs facteurs d’authentification supplémentaires. 
    * L’option `authChallenge` contient une chaîne qui indique à l’AAD quel(s) facteur(s) il doit demander. L’hôte Office transmet cette chaîne à l’AAD lorsqu’il demande à votre complément le jeton de complément.

    ```
    function getDataUsingAuthChallenge(authChallengeString) {       
        Office.context.auth.getAccessTokenAsync({authChallenge: authChallengeString},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/onedriveitems", accessToken);
                }
                else {
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }
    ```

1. Enregistrez et fermez le fichier.

## <a name="code-the-server-side"></a>Code côté serveur

Il existe deux fichiers côté serveur qui doivent être modifiés. 
- Le fichier src\auth.js fournit des fonctions d’assistance pour l’autorisation. Il dispose déjà des membres génériques qui sont utilisés dans une variété de flux d’autorisation. Nous devons ajouter des fonctions qui implémentent le flux « de la part de ».
- Le fichier src\server.js possède les membres de base requis pour exécuter un serveur et les intergiciels express. Nous devons y ajouter des fonctions qui servent la page d’accueil et une API Web pour obtenir des données Microsoft Graph.

### <a name="create-a-method-to-exchange-tokens"></a>Créer une méthode pour échanger des jetons

1. Ouvrez le fichier \src\auth.ts. Ajoutez la méthode ci-après à la classe `AuthModule`. Tenez compte des informations suivantes :
    * Le paramètre jwt est le jeton d’accès à l’application. Dans le flux « de la part de », il est échangé avec AAD pour un jeton d’accès à la ressource.
    * Le paramètre scopes a une valeur par défaut, mais dans cet exemple, elle sera remplacée par le code appelant.
    * Le paramètre de ressource est facultatif. Il ne doit pas être utilisé lorsque le STS est le point de terminaison AAD V2. Ce dernier déduit la ressource des étendues et renvoie une erreur si une ressource est envoyée dans la requête HTTP. 
    

    ```
    private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        try {
            // TODO3: Construct the parameters that will be sent in the body of the 
            //        HTTP Request to the STS that starts the "on behalf of" flow.
            // TODO4: Send the request to the STS.
            // TODO5: Process the response and persist the access token to resource.
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                        + JSON.stringify(exception), 
                                        exception);
        }
    }
    ```

2. Remplacez TODO3 par les lignes suivantes. À propos de ce code, notez que :
    * Un STS qui prend en charge le flux « de la part de » attend certaines paires de propriété/valeur dans le corps de la requête HTTP. Ce code construit un objet qui devient le corps de la requête. 
    * Une propriété de ressource est ajoutée au corps si, et uniquement si, une ressource a été transmise à la méthode.

    ```
    const v2Params = {
            client_id: this.clientId,
            client_secret: this.clientSecret,
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            assertion: jwt,
            requested_token_use: 'on_behalf_of',
            scope: scopes.join(' ')
        };
        let finalParams = {};
        if (resource) {
            // In JavaScript we could just add the resource property to the v2Params
            // object, but that won't compile in TypeScript.
            let v1Params  = { resource: resource };  
            for(var key in v2Params) { v1Params[key] = v2Params[key]; }
            finalParams = v1Params;
        } else {
            finalParams = v2Params;
        } 
    ```

3. Remplacez TODO4 par le code suivant, qui envoie la requête HTTP au point de terminaison de jeton du STS.

    ```
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. Remplacez TODO5 par le code suivant. Notez que le code prolonge le jeton d’accès à la ressource et son délai d’expiration, en plus de la renvoyer. Le code appelant permet d’éviter les appels inutiles au STS en réutilisant un jeton d’accès non expiré à la ressource. Vous verrez comment procéder dans la section suivante.

    ```
    if (res.status !== 200) {
        TODO6: Handle failure and the case where AAD asks for additional
               authentication factors.
    }
    const json = await res.json();
    // Persist the token and it's expiration time.
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

5. Remplacez TODO6 par le code suivant. À propos de ce code, notez que :

    * Il existe des configurations de Azure Active Directory où l’on demande à l’utilisateur de fournir un ou plusieurs facteurs d’authentification supplémentaires pour accéder à certaines cibles Microsoft Graph (par exemple, OneDrive), même si l’utilisateur peut se connecter à Office par un simple mot de passe. Dans ce cas, l’AAD envoie une réponse comportant une propriété `Claims`. 
    * Cette valeur `Claims` doit être transmise au client, qui doit démarrer une deuxième authentification de l’utilisateur et inclure la valeur `Claims` dans l’appel vers l’AAD. L’AAD invite l’utilisateur à fournir le(s) facteur(s) supplémentaire(s).
    * Par précaution, le code désactive le cache de tout jeton d’accès obtenu lorsque l’utilisateur est connecté avec un simple mot de passe.  

    ```
    const exception = await res.json();
    // Check if AAD is the STS.
    if (this.stsDomain === 'https://login.microsoftonline.com') {
        if (JSON.stringify(exception.claims)) {                       
            ServerStorage.clear();
            return JSON.stringify(exception.claims);    
        } else {                    
            throw exception;
        }
    }
    else {                    
        throw exception;
    }
    ```

5. Enregistrez le fichier, mais ne le fermez pas.

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>Créer une méthode pour accéder à la ressource à l’aide du flux « de la part de »

1. Toujours dans src/auth.ts, ajoutez la méthode ci-après à la classe `AuthModule`. Tenez compte des informations suivantes :
    * Les commentaires ci-dessus concernant les paramètres de la méthode `exchangeForToken` s’appliquent aussi aux paramètres de cette méthode.
    * La méthode recherche d’abord dans le stockage permanent un jeton d’accès à la ressource qui n’a pas expiré et qui ne va pas expirer dans la minute qui suit. Il appelle la méthode `exchangeForToken` que vous avez créée dans la dernière section uniquement si nécessaire.

    ```
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. Enregistrez et fermez le fichier.

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>Créer les points de terminaison que serviront la page d’accueil et les données du complément

1. Ouvrez le fichier src\server.ts. 

2. Ajoutez la méthode suivante au bas du fichier. Cette méthode servira la page d’accueil du complément. Le manifeste du complément spécifie l’URL de la page d’accueil.

    ```
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. Ajoutez la méthode suivante en bas du fichier. Cette méthode traite toutes les requêtes concernant l’API `onedriveitems`.
    ```
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Send to the client only the data that it actually needs.
    })); 
    ```

4. Remplacez TODO7 par le code suivant, qui valide le jeton d’accès reçu à partir de l’application hôte Office. La méthode `verifyJWT` est définie dans le fichier src\auth.ts. Elle valide toujours l’audience et l’émetteur. Nous utilisons le paramètre facultatif pour spécifier que nous souhaitons également vérifier que l’étendue du jeton d’accès est `access_as_user`. C’est la seule autorisation du complément dont l’utilisateur et l’hôte Office ont besoin pour obtenir un jeton d’accès à Microsoft Graph au moyen du flux « de la part de ». 

    ```
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

> **Remarque :** Vous ne pouvez utiliser l’étendue `access_as_user` que pour autoriser l’API qui gère le flux « de la part de » pour les compléments Office. D’autres API dans votre service peuvent avoir leurs propres exigences d’étendue. Cela permet de limiter ce à quoi donnent accès les jetons acquis par Office.

5. Remplacez l’élément TODO8 par le code suivant. Tenez compte des informations suivantes :

    * L’appel vers `acquireTokenOnBehalfOf` ne comprend pas de paramètre de ressource, étant donné que nous avons construit l’objet `AuthModule` (`auth`) avec le point de terminaison AAD V2.0 qui ne prend pas en charge une propriété de ressource.
    * Le deuxième paramètre de l’appel spécifie les autorisations dont le complément aura besoin pour obtenir une liste des fichiers et dossiers de l’utilisateur dans OneDrive. (L’autorisation `profile` n’est pas demandée, car elle n’est nécessaire qu’au moment où l’hôte Office obtient le jeton d’accès à votre complément, pas lorsque vous travaillez dans ce jeton pour un jeton d’accès à Microsoft Graph.)
    * Si la réponse est une chaîne contenant « capolids », il s’agit d’un message de revendication de l’AAD spécifiant qu’une authentification à facteurs multiples est requise. Le message est transmis au client, qui l’utilise pour démarrer une deuxième authentification. La chaîne indique à l’AAD quel(s) facteur(s) d’authentification supplémentaire(s) il doit inviter l’utilisateur à fournir.

    ```
    let graphToken = null;
    const tokenAcquisitionResponse = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    if (tokenAcquisitionResponse.includes('capolids')) {
        const claims: string[] = [];
        claims.push(tokenAcquisitionResponse);
        return res.json(claims);
    } else {
        // The response is the token to Microsoft Graph itself. Rename it so remaining code
        // is self-documenting.
        graphToken = tokenAcquisitionResponse;
    }
    ```

6. Remplacez TODO9 par la ligne suivante. Tenez compte des informations suivantes :

    * La classe MSGraphHelper est définie dans src\msgraph-helper.ts. 
    * Nous réduisons les données qui doivent être renvoyées en spécifiant que nous ne souhaitons que la propriété name et uniquement les 3 premiers éléments.

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. Remplacez TODO10 par le code suivant. Notez que Microsoft Graph renvoie des métadonnées OData et une propriété **eTag** pour chaque élément, même si `name` est la seule propriété demandée. Le code envoie uniquement les noms d’éléments au client.

    ```
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. Enregistrez et fermez le fichier.

## <a name="deploy-the-add-in"></a>Déploiement du complément

Vous devez maintenant indiquer à Office où trouver le complément.

1. Créez un partage réseau, ou [partagez un dossier sur le réseau](https://technet.microsoft.com/fr-fr/library/cc770880.aspx).

2. Placez une copie du fichier manifeste Office-Add-in-NodeJS-SSO.xml, depuis la racine du projet, dans le dossier partagé.

3. Lancez PowerPoint et ouvrez un document.

4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

5. Choisissez l’onglet **Fichier**, puis choisissez **Options**.

6. Choisissez **Catalogues de compléments approuvés**.

7. Dans le champ **URL du catalogue**, saisissez le chemin réseau permettant d’accéder au partage de dossier qui contient le fichier Office-Add-in-NodeJS-SSO.xml, puis sélectionnez **Ajouter un catalogue**.

8. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.

9. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage de Microsoft Office. Fermez PowerPoint.

## <a name="build-and-run-the-project"></a>Création et exécution du projet

Il existe deux manières de créer et d’exécuter le projet selon que vous utilisez Visual Studio Code. Pour les deux façons, le projet est généré et reconstruit automatiquement, puis ré-exécuté lorsque vous apportez des modifications au code.

1. Si vous n’utilisez pas Visual Studio Code : 
 1. Ouvrez un terminal de nœud et accédez au dossier racine du projet.
 2. Dans le terminal, entrez **npm run build**. 
 3. Ouvrez un second terminal de nœud et accédez au dossier racine du projet.
 4. Dans le terminal, entrez **npm run start**.

2. Si vous utilisez VS Code :
 1. Ouvrez le projet dans VS Code.
 2. Appuyez sur CTRL-MAJ-B pour générer le projet.
 3. Appuyez sur F5 pour exécuter le projet dans une session de débogage.


## <a name="add-the-add-in-to-an-office-document"></a>Ajouter le complément à un document Office

1. Redémarrez PowerPoint et ouvrez ou créez une présentation. 

2. Dans l’onglet **Développeur** de PowerPoint, choisissez **Mes compléments**.

3. Sélectionnez l’onglet **DOSSIER PARTAGÉ**.

4. Choisissez **Échantillon SSO NodeJS**, puis sélectionnez **OK**.

5. Dans le ruban **Accueil**, un nouveau groupe appelé **SSO NodeJS** apparaît avec un bouton intitulé **Afficher le complément** et une icône. 

## <a name="test-the-add-in"></a>Test du complément

1. Assurez-vous que vous disposez de fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.

2. Cliquez sur le bouton **Afficher le complément** pour ouvrir le complément.

2. Le complément s’ouvre avec une page d’accueil. Cliquez sur le bouton **Obtenir mes fichiers à partir de OneDrive**.

2. Si vous êtes connecté à Office, une liste de vos fichiers et dossiers sur OneDrive apparaîtront en dessous du bouton. La première fois, l’opération peut prendre plus de 15 secondes.

3. Si vous n’êtes pas connecté à Office, une fenêtre contextuelle s’ouvre et vous invite à vous connecter. Une fois que vous êtes connecté, la liste de vos fichiers et dossiers s’affiche après quelques secondes. *Vous n’appuyez pas sur le bouton une deuxième fois.*
> **Remarque :** Si vous étiez précédemment connecté à Office avec un ID différent, et si certaines applications Office sont toujours ouvertes, Office ne changera pas systématiquement votre identifiant même s’il semble l’avoir fait dans PowerPoint. Dans ce cas, l’appel vers Microsoft Graph peut échouer, ou des données de l’ID précédent peuvent être renvoyées. Afin d’éviter ce problème, veillez à *fermer toutes les autres applications Office* avant de cliquer sur **Obtenir mes fichiers à partir de OneDrive**.
