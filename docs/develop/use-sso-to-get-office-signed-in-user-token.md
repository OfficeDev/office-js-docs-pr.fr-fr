---
title: Utiliser l’utilisateur unique pour obtenir l’identité de l’utilisateur qui est inscrit
description: Appelez l’API getAccessToken pour obtenir le jeton d’ID avec le nom, le courrier électronique et des informations supplémentaires sur l’utilisateur connexion.
ms.date: 01/25/2022
localization_priority: Normal
ms.openlocfilehash: 2c9b3c89a154d624f99e196014c7d8024286d927
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62322333"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>Utiliser l’utilisateur unique pour obtenir l’identité de l’utilisateur qui est inscrit

Utilisez l’API `getAccessToken` pour obtenir un jeton d’accès qui contient l’identité de l’utilisateur actuel qui s’est Office. Le jeton d’accès est également un jeton d’ID, car il contient des revendications d’identité sur l’utilisateur qui est signé, telles que son nom et son e-mail. Vous pouvez également utiliser le jeton d’ID pour identifier l’utilisateur lors de l’appel de vos propres services web. Pour appeler`getAccessToken`, vous devez configurer votre Office de manière à utiliser l’luiso avec Office.

Dans cet article, vous allez créer un Office qui obtient le jeton d’ID et affiche le nom de l’utilisateur, le courrier électronique et l’ID unique dans le volet Des tâches.

> [!NOTE]
> SSO avec Office et l’API `getAccessToken` ne fonctionnent pas dans tous les scénarios. Implémentez toujours une boîte de dialogue de base pour vous connectez à l’utilisateur lorsque l’ssO n’est pas disponible. Pour plus d’informations, voir [Authenticate and authorize with the Office dialog API](auth-with-office-dialog-api.md).

## <a name="create-an-app-registration"></a>Créer une inscription d’application

Pour utiliser l’authentification unique avec Office, vous devez créer une inscription d’application dans le portail Azure afin que le Plateforme d'identités Microsoft puisse fournir des services d’authentification et d’autorisation pour votre Office et ses utilisateurs.

1. Pour inscrire votre application, go to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page.

1. Connectez-vous avec **_les informations d’identification_** d’administrateur à Microsoft 365 location. Par exemple, MonNom@contoso.onmicrosoft.com.

1. Sélectionnez **Nouvelle inscription**. Sur la page **Inscrire une application**, définissez les valeurs comme suit.

   - Définissez le **Nom** sur `Office-Add-in-SSO`.
   - Définissez les **Types de comptes pris en charge** à **Comptes dans un annuaire organisationnel et les comptes personnels Microsoft (par ex. Skype, Xbox et Outlook.com)**.
   - Définissez le type d’application **sur Web** , puis définissez **l’URI de redirection** sur `https://localhost:[port]/dialog.html`. Remplacez `[port]` par le numéro de port correct pour votre application web. Si vous avez créé le add-in à l’aide de yo office, le numéro de port est généralement 3000 et se trouve dans le fichier package.json. Si vous avez créé le Visual Studio 2019, le port se trouve dans la propriété **URL SSL** du projet web.
   - Choisissez **Inscrire**.

1. Dans **la page Office-Add-in-SSO**, copiez et enregistrez les valeurs de l’ID d’application **(client) et de l’ID** d’annuaire **(client**). Vous utiliserez les deux plus tard.

   > [!NOTE]
   > Cet ID d’application **(client)** est la valeur « audience » lorsque d’autres applications, telles que l’application cliente Office (par exemple, PowerPoint, Word, Excel), recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Sous **Gérer**, sélectionnez **Authentification**. Dans la section **Octroi implicite** , activez les case à cocher pour le jeton **Access et** le **jeton d’ID**.

1. En haut du formulaire, sélectionnez **Enregistrer**.

1. Sélectionnez **Exposer une API** sous **Gérer**. Sélectionnez **le lien** Définir. Cela génère l’URI de l’ID d’application sous la `api://[app-id-guid]`forme , où `[app-id-guid]` se trouve **l’ID d’application (client**).

1. Dans l’ID généré, `localhost:[port]/` insérez (notez la barre oblique « / » à la fin) entre les doubles barres obliques et le GUID. Remplacez `[port]` par le numéro de port correct pour votre application web. Si vous avez créé le add-in à l’aide de yo office, le numéro de port est généralement 3000 et se trouve dans le fichier package.json. Si vous avez créé le Visual Studio 2019, le port se trouve dans la propriété **URL SSL** du projet web.
   Lorsque vous avez terminé, l’ID entier doit avoir le formulaire `api://localhost:[port]/[app-id-guid]`; par exemple `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Sélectionnez le bouton **Ajouter une étendue**. Dans le volet qui s’ouvre, entrez `access_as_user` en tant que **nom de l’étendue**.

1. Donnez la valeur **Administrateurs et utilisateurs** à **Qui peut donner son consentement ?** .

1. Remplissez les champs pour configurer les invites de consentement de l’administrateur et de l’utilisateur avec des valeurs `access_as_user` appropriées pour l’étendue, ce qui permet à l’application cliente Office d’utiliser les API web de votre add-in avec les mêmes droits que l’utilisateur actuel. Suggestions :

   - **Nom complet du consentement de** l’administrateur : Office peut agir en tant qu’utilisateur.
   - **Description consentement administrateur** : activez Office pour qu’il appelle les API de complément web avec les mêmes droits que l’utilisateur actuel.
   - **Nom complet du consentement de** l’utilisateur : Office peut agir en votre nom.
   - **Description du consentement de** l’utilisateur : Office pour appeler les API web du add-in avec les mêmes droits que vous.

1. Vérifiez que **State** est défini comme **Activé**.

1. Sélectionnez **Ajouter une étendue**.

   > [!NOTE]
   > La partie domaine du **Nom de l’étendue** affiché juste sous le champ de texte devrait automatiquement correspondre à l’URI d’ID d’application définie à l’étape précédente avec `/access_as_user`ajouté à la fin, par exemple, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Dans la section **Applications client autorisées**, vous identifiez les applications que vous souhaitez autoriser dans l’application web de votre complément. Chacun des ID suivants doit être pré-autorisé.

   - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
   - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office sur le web)
   - `08e18876-6177-487e-b8b5-cf950c1e598c` (Office sur le web)
   - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook sur le web)

   Pour chaque ID, procédez comme suit :

   a. **Sélectionnez Ajouter un bouton d’application cliente**, puis, dans le panneau qui s’ouvre, `[app-id-guid]` définissez l’ID d’application (client) et cochez la case .`api://localhost:44355/[app-id-guid]/access_as_user`

   b. Sélectionnez **Ajouter une application**.

1. Sélectionnez **Autorisations API** sous **Gestion** et sélectionnez **Ajouter une autorisation**. Dans le volet qui s’ouvre, sélectionnez **Microsoft Graph**, puis **Autorisations déléguées**.

1. Utilisez la zone de recherche **Sélectionnez les autorisations** pour rechercher les autorisations dont votre complément a besoin. Recherchez et sélectionnez **l’autorisation de** profil. L’autorisation `profile` est requise pour que l Office’application obtienne un jeton pour votre application web de add-in.

   - profil

   > [!NOTE]
   > L’autorisation `User.Read` est peut-être déjà répertoriée par défaut. Une bonne pratique consiste à demander uniquement les autorisations dont vous avez besoin. Ainsi, nous vous recommandons de désactiver la case à cocher de cette autorisation si votre complément n’en a pas réellement besoin.

1. Sélectionner le bouton **Ajouter des autorisations** en bas du panneau.

1. Sur la même page, choisissez le **\<tenant-name\>** bouton Accorder le consentement administrateur, puis sélectionnez **Oui** pour la confirmation qui s’affiche.

## <a name="create-the-office-add-in"></a>Créer le complément Office

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Démarrez Visual Studio 2019 et choisissez **de créer un projet**.
1. Recherchez et sélectionnez **le modèle Excel de projet de l’Application Web**. Sélectionnez **Suivant**. Remarque : SSO fonctionne avec n’importe quelle application Office, mais pour cet article fonctionne avec Excel.
1. Entrez un nom de projet, tel que **sso-display-user-info** , puis choisissez **Créer**. Vous pouvez laisser les valeurs par défaut des autres champs.
1. Dans la **boîte de dialogue Choisir le type de add-in**, sélectionnez Ajouter de nouvelles fonctionnalités **à Excel**, puis choisissez **Terminer**.

Le projet est créé et contiendra deux projets dans la solution.

- **sso-display-user-info** : contient le manifeste et les détails pour le chargement de version de version de chargement du Excel.
- **sso-display-user-infoWeb** : projet ASP.NET qui héberge les pages web du module complémentaire.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Assurez-vous que vous [avez bien installé votre environnement de développement](../overview/set-up-your-dev-environment.md).

1. Pour créer le projet, entrez la commande suivante.

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

Le projet est créé dans un nouveau dossier nommé **sso-display-user-info**.

---

## <a name="configure-the-manifest"></a>Configurer le manifeste

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Dans **l’Explorateur** de solutions, **ouvrez sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Dans Visual Studio code **, ouvrezmanifest.xml** fichier.

---

1. Dans la partie inférieure du manifeste se trouve un élément de `</Resources>` fermeture. Insérez le XML suivant juste en dessous de l’élément `</Resources>` , mais avant l’élément `</VersionOverrides>` de fermeture. Pour Office applications autres que Outlook, ajoutez le signet à la fin de la `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. Pour Outlook, ajoutez le balisage à la fin de la section `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

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

1. Remplacez `[port]` par le numéro de port correct pour votre projet. Si vous avez créé le add-in à l’aide de yo office, le numéro de port est généralement 3000 et se trouve dans le fichier package.json. Si vous avez créé le Visual Studio 2019, le port se trouve dans la propriété **URL SSL** du projet web.
1. Remplacez les deux `[application-id]` espaces réservé par l’ID d’application réel de l’inscription de votre application.
1. Enregistrez le fichier.

Le XML que vous avez inséré contient les éléments et informations suivants.

- **WebApplicationInfo**: le parent des éléments suivants.
- **Id** - ID du client du compl?ment : il  s'agit d'un ID d'application que vous obtenez lors de l'enregistrement du compl?ment. Voir[Enregistrer un complément Office utilisant une SSO (authentification unique) avec le point de terminaison Azure AD v2.0](register-sso-add-in-aad-v2.md).
- **Ressource**: l’URL du complément. Il s’agit du même URI (y compris le protocole`api:`) que vous avez utilisé lors de l’inscription du complément dans AAD. Le domaine et les sous-domaines doivent être les mêmes que ceux utilisés dans les URL dans la section`<Resources>` du manifeste du complément et l’URI doit se terminer par l’ID client dans le `<Id>`.
- **Scopes**: le parent d’un ou plusieurs éléments **Scope**.
- **Scope**: spécifie une autorisation nécessaire pour le complément dans l’AAD. Les autorisations `profile` et `openID` sont toujours nécessaires et peuvent être les seules autorisation nécessaires si votre complément n'accepte pas l’accès à Microsoft Graph. Si c'est le cas, vous avez ?galement besoin des ?l?ments d'une **?tendue** pour obtenir les autorisations Microsoft Graph requises; par exemple, `User.Read`, `Mail.Read`. Les biblioth?ques que vous utilisez dans votre code pour acc?der ? Microsoft Graph peuvent avoir des besoin d'autorisations suppl?mentaires. Par exemple, Microsoft Authentication Library (MSAL) pour .NET n?cessite `offline_access` une autorisation. Pour plus d'informations, voir [Autoriser Microsoft Graph ? partir d'un compl?ment Office](authorize-to-microsoft-graph.md).

## <a name="add-the-jwt-decode-package"></a>Ajouter le package de décodage jwt

Vous pouvez appeler l’API `getAccessToken` pour obtenir le jeton d’ID de Office. Permet d’abord d’ajouter le package de décodage jwt pour faciliter le décodage et l’affichage du jeton d’ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Ouvrez la Visual Studio solution.
1. Dans le menu, sélectionnez **Outils > NuGet Gestionnaire de package > Gestionnaire de package Console**.
1. Entrez la commande suivante dans la **console Gestionnaire de package.**

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. À partir d’une fenêtre terminal/console, allez dans le dossier racine de votre projet de add-in.
1. Entrez la commande suivante

   `npm install jwt-decode`

---

## <a name="add-ui-to-the-task-pane"></a>Ajouter une interface utilisateur au volet Des tâches

Nous devons modifier le volet Des tâches afin qu’il puisse afficher les informations utilisateur que nous allons obtenir à partir du jeton d’ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Ouvrez Home.html fichier.
1. Ajoutez la balise de script suivante à la `<head>` section de la page. Cela inclut le package de décodage jwt que nous avons ajouté précédemment.

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

1. Ouvrez **le fichier src/taskpane/taskpane.html** .
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

1. Ouvrez **Home.js** fichier.
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

1. Ouvrez **le fichier src/taskpane/taskpane.js** .
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

1. Choisissez **Déboguer > démarrer le débogage**, ou appuyez **sur F5**.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Exécuter à `npm start` partir de la ligne de commande.

---

1. Lorsque Excel démarre, connectez-vous Office avec le même compte client que celui que vous avez utilisé pour créer l’inscription de l’application.
1. Dans le **ruban Accueil** , **sélectionnez Afficher lepane des tâches** pour ouvrir le module.
1. Dans le volet Des tâches du module, sélectionnez **Obtenir un jeton d’ID**.

Le add-in affiche le nom, le courrier électronique et l’ID du compte avec qui vous vous êtes inscrit.

> [!NOTE]
> Si vous rencontrez des erreurs, examinez les étapes d’inscription de cet article pour l’inscription de l’application. L’absence d’un détail lors de la configuration de l’inscription de l’application est une cause courante de problèmes d’utilisation de l’unique. Si vous ne par arrivez toujours pas à obtenir le add-in pour qu’il s’exécute correctement, voir Résolution des problèmes de messages d’erreur pour [l’sign-on unique (SSO).](troubleshoot-sso-in-office-add-ins.md)

## <a name="see-also"></a>Voir aussi

[Utilisation de revendications pour identifier de manière fiable un utilisateur (objet et ID d’objet)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)
