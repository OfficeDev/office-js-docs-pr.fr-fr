---
title: Utiliser l’authentification unique pour obtenir l’identité de l’utilisateur connecté
description: Appelez l’API getAccessToken pour obtenir le jeton d’ID avec le nom, l’e-mail et des informations supplémentaires sur l’utilisateur connecté.
ms.date: 02/16/2022
localization_priority: Normal
ms.openlocfilehash: 2e8cc0074f5b6f4f5598320f07c8bf5c0a7b301d
ms.sourcegitcommit: 3c5ede9c4f9782947cea07646764f76156504ff9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/06/2022
ms.locfileid: "64682237"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>Utiliser l’authentification unique pour obtenir l’identité de l’utilisateur connecté

Utilisez l’API `getAccessToken` pour obtenir un jeton d’accès qui contient l’identité de l’utilisateur actuel connecté à Office. Le jeton d’accès est également un jeton d’ID, car il contient des revendications d’identité relatives à l’utilisateur connecté, telles que son nom et son e-mail. Vous pouvez également utiliser le jeton d’ID pour identifier l’utilisateur lors de l’appel de vos propres services web. Pour appeler`getAccessToken`, vous devez configurer votre complément Office pour utiliser l’authentification unique avec Office.

Dans cet article, vous allez créer un complément Office qui obtient le jeton d’ID et affiche le nom, l’e-mail et l’ID unique de l’utilisateur dans le volet Office.

> [!NOTE]
> L’authentification unique avec Office et l’API `getAccessToken` ne fonctionne pas dans tous les scénarios. Implémentez toujours une boîte de dialogue de secours pour connecter l’utilisateur lorsque l’authentification unique n’est pas disponible. Pour plus d’informations, consultez [Authentifier et autoriser avec l’API de dialogue Office](auth-with-office-dialog-api.md).

## <a name="create-an-app-registration"></a>Créer une inscription d’application

Pour utiliser l’authentification unique avec Office, vous devez créer une inscription d’application dans le Portail Azure afin que le Plateforme d'identités Microsoft puisse fournir des services d’authentification et d’autorisation pour votre complément Office et ses utilisateurs.

1. Pour inscrire votre application, accédez à la page [Portail Azure - inscriptions d'applications](https://go.microsoft.com/fwlink/?linkid=2083908).

1. Connectez-vous avec les informations **_d’identification d’administrateur_** à votre Microsoft 365 location. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez **Nouvelle inscription**. Sur la page **Inscrire une application**, définissez les valeurs comme suit.

   - Définissez le **Nom** sur `Office-Add-in-SSO`.
   - Définissez les **Types de comptes pris en charge** à **Comptes dans un annuaire organisationnel et les comptes personnels Microsoft (par ex. Skype, Xbox et Outlook.com)**.
   - Définissez le type d’application sur **Web** , puis **définissez l’URI** de `https://localhost:[port]/dialog.html`redirection sur . Remplacez par `[port]` le numéro de port approprié pour votre application web. Si vous avez créé le complément à l’aide de yo office, le numéro de port est généralement 3000 et se trouve dans le fichier package.json. Si vous avez créé le complément avec Visual Studio 2019, le port se trouve dans la propriété **URL SSL** du projet web.
   - Choisissez **Inscrire**.

1. Dans la page **Office-Add-in-SSO**, copiez et enregistrez les valeurs de **l’ID d’application (client)** et de **l’ID d’annuaire (locataire**). Vous utiliserez les deux plus tard.

   > [!NOTE]
   > Cet **ID d’application (client)** est la valeur « audience » lorsque d’autres applications, telles que l’application cliente Office (par exemple, PowerPoint, Word Excel), recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Sous **Gérer**, sélectionnez **Authentification**. Dans la section **Octroi implicite** , activez les cases à cocher pour le **jeton d’accès** et le **jeton d’ID**.

1. En haut du formulaire, sélectionnez **Enregistrer**.

1. Sélectionnez **Exposer une API** sous **Gérer**. Sélectionnez le lien **Définir** . Cela génère l’URI d’ID d’application dans le formulaire `api://[app-id-guid]`, où `[app-id-guid]` se trouve **l’ID d’application (client**).

1. Dans l’ID généré, insérez `localhost:[port]/` (notez la barre oblique « / » ajoutée à la fin) entre les barres obliques doubles et le GUID. Remplacez par `[port]` le numéro de port approprié pour votre application web. Si vous avez créé le complément à l’aide de yo office, le numéro de port est généralement 3000 et se trouve dans le fichier package.json. Si vous avez créé le complément avec Visual Studio 2019, le port se trouve dans la propriété **URL SSL** du projet web.
   Lorsque vous avez terminé, l’ID entier doit avoir le formulaire `api://localhost:[port]/[app-id-guid]`, par exemple `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Sélectionnez le bouton **Ajouter une étendue**. Dans le volet qui s’ouvre, entrez `access_as_user` en tant que **nom de l’étendue**.

1. Donnez la valeur **Administrateurs et utilisateurs** à **Qui peut donner son consentement ?** .

1. Renseignez les champs permettant de configurer les invites de consentement de l’administrateur et de l’utilisateur avec des valeurs appropriées pour l’étendue `access_as_user` qui permet à l’application cliente Office d’utiliser les API web de votre complément avec les mêmes droits que l’utilisateur actuel. Suggestions :

   - **Nom d’affichage du consentement de** l’administrateur : Office pouvez agir en tant qu’utilisateur.
   - **Description consentement administrateur** : activez Office pour qu’il appelle les API de complément web avec les mêmes droits que l’utilisateur actuel.
   - **Nom d’affichage du consentement de** l’utilisateur : Office pouvez agir comme vous.
   - **Description du consentement de l’utilisateur** : activez Office pour appeler les API web du complément avec les mêmes droits que vous.

1. Vérifiez que **State** est défini comme **Activé**.

1. Sélectionnez **Ajouter une étendue**.

   > [!NOTE]
   > La partie domaine du **Nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’URI d’ID d’application définie à l’étape précédente avec `/access_as_user`ajouté à la fin, par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Dans la section **Applications clientes autorisées**, entrez l’ID suivant pour pré-autoriser tous les points de terminaison d’application Microsoft Office.

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`(Tous les points de terminaison d’application Microsoft Office)

    > [!NOTE]
    > L’ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` pré-autorise Office sur toutes les plateformes suivantes. Vous pouvez également entrer un sous-ensemble approprié des ID suivants si, pour une raison quelconque, vous souhaitez refuser l’autorisation de Office sur certaines plateformes. Il vous suffit d’exclure les ID des plateformes à partir desquelles vous souhaitez refuser l’autorisation. Les utilisateurs de votre complément sur ces plateformes ne pourront pas appeler vos API web, mais d’autres fonctionnalités de votre complément fonctionneront toujours.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office sur le web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook sur le web)

1. Sélectionnez le bouton **Ajouter une application cliente**, puis, dans le panneau qui s’ouvre, définissez l’ID `[app-id-guid]` d’application (client) et cochez la case .`api://localhost:44355/[app-id-guid]/access_as_user`

1. Sélectionnez **Ajouter une application**.

1. Sélectionnez **Autorisations API** sous **Gestion** et sélectionnez **Ajouter une autorisation**. Dans le volet qui s’ouvre, sélectionnez **Microsoft Graph**, puis **Autorisations déléguées**.

1. Utilisez la zone de recherche **Sélectionnez les autorisations** pour rechercher les autorisations dont votre complément a besoin. Recherchez et sélectionnez l’autorisation de **profil** . L’autorisation `profile` est requise pour que l’application Office obtienne un jeton pour votre application web de complément.

   - profil

   > [!NOTE]
   > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Une bonne pratique consiste à demander uniquement les autorisations dont vous avez besoin. Ainsi, nous vous recommandons de désactiver la case à cocher de cette autorisation si votre complément n’en a pas réellement besoin.

1. Sélectionner le bouton **Ajouter des autorisations** en bas du panneau.

1. Dans la même page, choisissez le **bouton Accorder le consentement \<tenant-name\>de l’administrateur**, puis sélectionnez **Oui** pour la confirmation qui s’affiche.

## <a name="create-the-office-add-in"></a>Créer le complément Office

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Démarrez Visual Studio 2019 et choisissez **de créer un projet**.
1. Recherchez et sélectionnez le **modèle de projet de complément web Excel**. Sélectionnez **Suivant**. Remarque : L’authentification unique fonctionne avec n’importe quelle application Office, mais pour cet article fonctionne avec Excel.
1. Entrez un nom de projet, tel que **sso-display-user-info** , puis choisissez **Créer**. Vous pouvez laisser les autres champs aux valeurs par défaut.
1. Dans la boîte **de dialogue Choisir le type de complément**, sélectionnez **Ajouter de nouvelles fonctionnalités pour Excel**, puis **choisissez Terminer**.

Le projet est créé et contiendra deux projets dans la solution.

- **sso-display-user-info** : contient le manifeste et les détails du chargement indépendant du complément pour Excel.
- **sso-display-user-infoWeb** : projet ASP.NET qui héberge les pages web du complément.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Assurez-vous que vous avez [configuré votre environnement de développement](../overview/set-up-your-dev-environment.md).

1. Pour créer le projet, entrez la commande suivante.

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

Le projet est créé dans un dossier nommé **sso-display-user-info**.

---

## <a name="configure-the-manifest"></a>Configurer le manifeste

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Dans **Explorateur de solutions** ouvrez **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Dans Visual Studio code, ouvrez le fichier **manifest.xml**.

---

1. Près du bas du manifeste se trouve un élément fermant `</Resources>` . Insérez le code XML suivant juste en dessous de l’élément `</Resources>` , mais avant l’élément fermant `</VersionOverrides>` . Pour Office applications autres que Outlook, ajoutez le balisage à la fin de la `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

   ```xml
   <WebApplicationInfo>
       <Id>[application-id]</Id>
       <Resource>api://localhost:[port]/[application-id]</Resource>
       <Scopes>
           <Scope>openid</Scope>
           <Scope>user.read</Scope>
           <Scope>profile</Scope>
       </Scopes>
   </WebApplicationInfo>
   ```

1. Remplacez par `[port]` le numéro de port approprié pour votre projet. Si vous avez créé le complément à l’aide de yo office, le numéro de port est généralement 3000 et se trouve dans le fichier package.json. Si vous avez créé le complément avec Visual Studio 2019, le port se trouve dans la propriété **URL SSL** du projet web.
1. Remplacez les deux `[application-id]` espaces réservés par l’ID d’application réel de l’inscription de votre application.
1. Enregistrez le fichier.

Le code XML que vous avez inséré contient les éléments et informations suivants.

- **WebApplicationInfo**: le parent des éléments suivants.
- **Id** - ID du client du compl?ment : il  s'agit d'un ID d'application que vous obtenez lors de l'enregistrement du compl?ment. Voir[Enregistrer un complément Office utilisant une SSO (authentification unique) avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).
- **Ressource**: l’URL du complément. Il s’agit du même URI (y compris le protocole`api:`) que vous avez utilisé lors de l’inscription du complément dans AAD. Le domaine et les sous-domaines doivent être les mêmes que ceux utilisés dans les URL dans la section`<Resources>` du manifeste du complément et l’URI doit se terminer par l’ID client dans le `<Id>`.
- **Scopes**: le parent d’un ou plusieurs éléments **Scope**.
- **Scope**: spécifie une autorisation nécessaire pour le complément dans l’AAD. Les autorisations `profile` et `openID` sont toujours nécessaires et peuvent être les seules autorisation nécessaires si votre complément n'accepte pas l’accès à Microsoft Graph. Si c'est le cas, vous avez ?galement besoin des ?l?ments d'une **?tendue** pour obtenir les autorisations Microsoft Graph requises; par exemple, `User.Read`, `Mail.Read`. Les biblioth?ques que vous utilisez dans votre code pour acc?der ? Microsoft Graph peuvent avoir des besoin d'autorisations suppl?mentaires. Par exemple, Microsoft Authentication Library (MSAL) pour .NET n?cessite `offline_access` une autorisation. Pour plus d'informations, voir [Autoriser Microsoft Graph ? partir d'un compl?ment Office](authorize-to-microsoft-graph.md).

## <a name="add-the-jwt-decode-package"></a>Ajouter le package jwt-decode

Vous pouvez appeler l’API `getAccessToken` pour obtenir le jeton d’ID à partir de Office. Tout d’abord, nous allons ajouter le package jwt-decode pour faciliter le décodage et l’affichage du jeton d’ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Ouvrez la solution Visual Studio.
1. Dans le menu, choisissez **Outils > NuGet Gestionnaire de package > Gestionnaire de package Console**.
1. Entrez la commande suivante dans la **console Gestionnaire de package**.

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. À partir d’une fenêtre de terminal/console, accédez au dossier racine de votre projet de complément.
1. Entrez la commande suivante

   `npm install jwt-decode`

---

## <a name="add-ui-to-the-task-pane"></a>Ajouter une interface utilisateur au volet Office

Nous devons modifier le volet Office afin qu’il puisse afficher les informations utilisateur que nous allons obtenir à partir du jeton d’ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Ouvrez le fichier Home.html.
1. Ajoutez la balise de script suivante à la `<head>` section de la page. Cela inclut le package jwt-decode que nous avons ajouté précédemment.

   ```html
   <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
   ```

1. Remplacez la `<body>` section par le code HTML suivant.

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Ouvrez le fichier **src/taskpane/taskpane.html** .
1. Remplacez la `<body>` section par le code HTML suivant.

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

---

## <a name="call-the-getaccesstoken-api"></a>Appeler l’API getAccessToken

La dernière étape consiste à obtenir le jeton d’ID en appelant `getAccessToken`.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Ouvrez le fichier **Home.js** .
1. Remplacez tout le contenu du fichier par le code suivant.

   ```javascript
   (function () {
     "use strict";

     // The initialize function must be run each time a new page is loaded.
     Office.initialize = function (reason) {
       $(document).ready(function () {
         $("#getIDToken").click(getIDToken);
       });
     };

     async function getIDToken() {
       try {
         let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
           allowSignInPrompt: true,
         });
         let userToken = jwt_decode(userTokenEncoded);
         document.getElementById("userInfo").innerHTML =
           "name: " +
           userToken.name +
           "<br>email: " +
           userToken.preferred_username +
           "<br>id: " +
           userToken.oid;
         console.log(userToken);
       } catch (error) {
         document.getElementById("userInfo").innerHTML =
           "An error occurred. <br>Name: " +
           error.name +
           "<br>Code: " +
           error.code +
           "<br>Message: " +
           error.message;
         console.log(error);
       }
     }
   })();
   ```

1. Enregistrez le fichier.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Ouvrez le fichier **src/taskpane/taskpane.js** .
1. Remplacez tout le contenu du fichier par le code suivant.

   ```javascript
   import jwt_decode from "jwt-decode";

   Office.onReady((info) => {
     if (info.host === Office.HostType.Excel) {
       document.getElementById("getIDToken").onclick = getIDToken;
     }
   });

   async function getIDToken() {
     try {
       let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
         allowSignInPrompt: true,
       });
       let userToken = jwt_decode(userTokenEncoded);
       document.getElementById("userInfo").innerHTML =
         "name: " +
         userToken.name +
         "<br>email: " +
         userToken.preferred_username +
         "<br>id: " +
         userToken.oid;
       console.log(userToken);
     } catch (error) {
       document.getElementById("userInfo").innerHTML =
         "An error occurred. <br>Name: " +
         error.name +
         "<br>Code: " +
         error.code +
         "<br>Message: " +
         error.message;
       console.log(error);
     }
   }
   ```

1. Enregistrez le fichier.

---

## <a name="run-the-add-in"></a>Exécuter du complément

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Choisissez **Déboguer > démarrer le débogage**, ou appuyez sur **F5**.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Exécutez à `npm start` partir de la ligne de commande.

---

1. Lorsque Excel démarre, connectez-vous à Office avec le compte de locataire que vous avez utilisé pour créer l’inscription de l’application.
1. Dans le ruban **Accueil** , choisissez **Afficher le volet Office** pour ouvrir le complément.
1. Dans le volet Office du complément, **choisissez Obtenir un jeton d’ID**.

Le complément affiche le nom, l’e-mail et l’ID du compte avec lequel vous vous êtes connecté.

> [!NOTE]
> Si vous rencontrez des erreurs, passez en revue les étapes d’inscription de cet article pour l’inscription de l’application. L’absence de détails lors de la configuration de l’inscription de l’application est une cause courante de problèmes liés à l’authentification unique. Si vous ne parvenez toujours pas à obtenir l’exécution du complément, consultez [Les messages d’erreur de résolution des problèmes pour l’authentification unique (SSO).](troubleshoot-sso-in-office-add-ins.md)

## <a name="see-also"></a>Voir aussi

[Utilisation de revendications pour identifier de manière fiable un utilisateur (objet et ID d’objet)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)

