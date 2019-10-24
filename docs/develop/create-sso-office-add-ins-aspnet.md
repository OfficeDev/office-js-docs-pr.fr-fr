---
title: Créer un complément Office ASP.NET qui utilise l’authentification unique
description: ''
ms.date: 10/11/2019
localization_priority: Priority
ms.openlocfilehash: 9844b8f9b9b966c0a5348f02f5797e7a07eb67b6
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626795"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>Créer un complément Office ASP.NET qui utilise l’authentification unique (aperçu)

Lorsque les utilisateurs sont connectés à Office, votre complément peut utiliser les mêmes informations d’identification pour permettre aux utilisateurs d’accéder à plusieurs applications sans avoir à se connecter une deuxième fois. Pour en savoir plus, consultez [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md).

Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré avec ASP.NET, OWIN et la bibliothèque d’authentification Microsoft (MSAL) pour .NET.

> [!NOTE]
> Pour un article similaire concernant un complément basé sur Node.js, consultez [Création d’un complément Office Node.js qui utilise l’authentification unique](create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Conditions préalables

* Version la plus récente disponible de Visual Studio 2019.

* Office 365 (version d’Office par abonnement). Dernière version mensuelle et build du canal du programme Insider. Vous devez participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1). Veuillez noter que lorsqu’un build passe au canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu, y compris l’authentification unique, est désactivée pour ce build.

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

1. Clonez ou téléchargez le référentiel sur [Complément Office ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).

1. Ouvrez le dossier **Before** et ouvrez le fichier .sln dans Visual Studio. Il s’agit d’un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés.

    > [!NOTE]
    > Il existe également une version finale de l’échantillon dans le même référentiel. Elle est équivalente au complément que vous obtiendriez si vous terminiez les procédures de cet article, sauf que le projet terminé comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, ouvrez simplement le fichier `sln` et suivez les instructions de cet article, mais ignorez les sections **Code côté client** et **Code côté serveur**.

1. Une fois le projet ouvert, générez-le dans Visual Studio, qui installera les packages répertoriés dans le fichier packages.config. L’opération peut prendre de quelques secondes à plusieurs minutes selon le nombre de packages présents dans le cache de packages de l’ordinateur local.

    > [!NOTE]
    > Vous obtiendrez une erreur relative à l’espace de noms Identity. Il s’agit d’un effet indésirable dû à un problème de configuration qui sera corrigé à la prochaine étape. Le plus important est que les packages soient bien installés.

1. Pour l’instant, la version de la bibliothèque MSAL (Microsoft.Identity.Client) dont vous avez besoin pour l’authentification unique (version `1.1.4-preview0002`) ne fait pas partie du catalogue NuGet standard, elle n’est donc pas répertoriée dans package.config et doit être installée séparément.

   > 1. Dans le menu **Outils**, accédez à **Gestionnaire de package NuGet** > **Console du Gestionnaire de package**.
   > 2. Dans la console, exécutez la commande suivante. L’opération peut prendre une minute ou plus, même avec une bonne connexion Internet. Une fois l’opération terminée, le message **Successfully installed ’Microsoft.Identity.Client 1.1.4-preview0002’ ...** doit être affiché vers la fin de la sortie de la console.
   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`
   > 3. Dans l’**Explorateur de solutions**, développez les **Références** du projet **Office-Add-in-ASPNET-SSO-WebAPI**. Vérifiez que **Microsoft.Identity.Client** est répertorié. S’il n’y est pas ou qu’une icône d’avertissement figure sur son entrée, supprimez l’entrée, puis utilisez l’Assistant Ajouter une référence Visual Studio pour ajouter une référence à l’assembly dans **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**

1. Créez le projet une deuxième fois.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Enregistrez le complément avec le point de terminaison Azure AD v2.0

Les instructions suivantes présentant un manière générique, vous pouvez les utiliser dans plusieurs emplacements. En lien avec ce article, procédez comme suit :

- Remplacez l’espace réservé **$ADD-IN-NAME$** par `Office-Add-in-ASPNET-SSO`.
- Remplacez l’espace réservé **$FQDN-WITHOUT-PROTOCOL$** par `localhost:44355`.
- Lorsque vous spécifiez des autorisations dans la boîte de dialogue **Sélectionner les autorisations**, cochez les cases correspondant aux autorisations suivantes. Seule la première est réellement nécessaire pour votre complément proprement dit, mais la bibliothèque MSAL utilisée par le code côté serveur requiert `offline_access` et `openid`. L’autorisation `profile` est requise pour l’hôte Office afin d’obtenir un jeton pour l’application web de votre complément.
  * Files.Read.All
  * offline_access
  * openid
  * profil


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a>Octroi du consentement administrateur pour le complément

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Configurer le complément

1. Dans la chaîne suivante, remplacez l’espace réservé “{tenant_ID}” par votre ID de client Office 365. Si vous n’avez pas copié l’ID de client lorsque vous avez inscrit le complément auprès d’AAD, utilisez une des méthodes dans [Trouver votre ID de client Office 365](/onedrive/find-your-office-365-tenant-id) pour l’obtenir.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. Dans Visual Studio, ouvrez le fichier web.config. Il existe certaines clés dans la section **appSettings** à laquelle vous devez affecter des valeurs.

1. Utilisez la chaîne que vous avez créée à l’étape 1 en tant que valeur pour la clé nommée « ida:Issuer ». Assurez-vous que la valeur ne comporte aucun espace vide.

1. Affectez les valeurs suivantes aux clés correspondantes :

    |Clé|Valeur|
    |:-----|:-----|
    |ida:ClientID|L’ID d’application que vous avez obtenu lorsque vous avez enregistré le complément.|
    |ida:Audience|L’ID d’application que vous avez obtenu lorsque vous avez enregistré le complément.|
    |ida:Password|Mot de passe que vous avez obtenu lorsque vous avez inscrit le complément.|

   Voici un exemple de ce à quoi doivent ressembler les quatre clés que vous avez modifiées. *Vous remarquerez que les clés ClientID et Audience sont identiques*. Vous pouvez également utiliser une seule clé pour les deux fonctions, mais votre balisage web.config sera mieux réutilisable si vous les séparez, car elles ne sont pas toujours identiques. En outre, des clés séparées renforcent l’idée que votre complément est à la fois une ressource OAuth, par rapport à l’hôte Office, et un client OAuth, par rapport à Microsoft Graph.

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />

    ```

   > [!NOTE]
   > Conservez tels quels les autres paramètres de la section **appSettings**.

1. Enregistrez et fermez le fichier.

1. Dans le projet de complément, ouvrez le fichier manifeste du complément « Office-Add-in-ASPNET-SSO.xml ».

1. Faites défiler vers le bas du fichier.

1. Juste au-dessus de la balise de fin `</VersionOverrides>`, vous trouverez le balisage suivant :

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Remplacez l’espace réservé « {application_GUID here} » *aux deux endroits* du balisage par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément. Les crochets « {} » ne font pas partie de l’ID, ne les incluez pas. C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.

    > [!NOTE]
    > * La valeur **Resource** correspond à l’**URI d’ID d’application** défini lorsque vous avez ajouté la plateforme d’API web à l’enregistrement du complément.
    > * La section **Scopes** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.

1. Ouvrez l’onglet **Avertissements** de la **liste d’erreurs** dans Visual Studio. Si un message d’avertissement indique que `<WebApplicationInfo>` n’est pas un enfant valide de `<VersionOverrides>`, votre version de Visual Studio ne reconnaît pas le balisage d’authentification unique. Solution de contournement : procédez comme suit pour un complément Word, Excel ou PowerPoint. (Si vous utilisez un complément Outlook, consultez la solution de contournement ci-dessous.)

   - **Solution de contournement pour Word, Excel et PowerPoint**

        1. Commentez la section `<WebApplicationInfo>` du manifeste juste au-dessus de la fin de `</VersionOverrides>`.

        2. Appuyez sur **F5** pour démarrer une session de débogage. Cette opération entraîne la création d’une copie du manifeste dans le dossier suivant (auquel il est plus facile d’accéder dans l’**Explorateur de fichiers** que dans Visual Studio) : `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

        3. Dans la copie du manifeste, supprimez la syntaxe de commentaire autour de la section `<WebApplicationInfo>`.

        4. Enregistrez la copie du manifeste.

        5. À présent, vous devez empêcher Visual Studio de remplacer la copie du manifeste la prochaine fois que vous appuyez sur F5. Cliquez avec le bouton droit de la souris sur le nœud de solution en haut de l’**explorateur de solutions** (et non sur l’un des nœuds de projet).

        6. Sélectionnez **Propriétés** dans le menu contextuel, puis une boîte de dialogue **Pages de propriétés de la solution** s’ouvre.

        7. Développez **Propriétés de configuration** et sélectionnez **Configuration**.

        8. Désélectionnez **Créer** et **Déployer** dans la ligne pour le projet **Office-Add-in-ASPNET-SSO** (et *pas* le projet **Office-Add-in-ASPNET-SSO-WebAPI**).

        9. Cliquez sur **OK** pour fermer la boîte de dialogue.

   - **Solution de contournement pour Outlook**

        1. Sur votre ordinateur de développement, recherchez l’élément `MailAppVersionOverridesV1_1.xsd` existant. Il doit se trouver dans le répertoire d’installation Visual Studio sous `./Xml/Schemas/{lcid}`. Par exemple, sur une installation standard de VS 2017 32 bits sur un système anglais (États-Unis), le chemin d’accès complet serait `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.

        2. Renommez le fichier existant comme suit : `MailAppVersionOverridesV1_1.old`.

        3. Copiez la version modifiée du fichier dans le dossier : [Schéma MailAppVersionOverrides modifié](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Enregistrez et fermez le fichier manifeste principal dans Visual Studio.

## <a name="code-the-client-side"></a>Code côté client

1. Ouvrez le fichier Home.js dans le dossier **Scripts**. Il contient déjà du code :
    * Une affectation à la méthode `Office.initialize` qui affecte elle-même un gestionnaire à l’événement ClickButton `getGraphAccessTokenButton`.
    * Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.
    * Une méthode `logErrors` qui consigne dans la console les erreurs qui ne sont pas destinées à l’utilisateur final.

1. En dessous de l’affectation au `Office.initialize`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :

    * La gestion des erreurs dans le complément tente parfois automatiquement d’obtenir un jeton d’accès une deuxième fois, à l’aide d’un autre jeu d’options. La variable de compteur `timesGetOneDriveFilesHasRun` et la variable d’indicateur `triedWithoutForceConsent` permettent de s’assurer que l’utilisateur ne tente pas de manière répétée d’obtenir un jeton sans y parvenir.
    * Vous allez créer la méthode `getDataWithToken` à l’étape suivante, mais rappelez-vous qu’elle définit une option appelée `forceConsent` sur `false`. Vous en saurez plus à la prochaine étape.

    ```js
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :

    * [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) est la nouvelle API d’Office.js qui permet à un complément de demander à l’application hôte Office (Excel, PowerPoint, Word, etc.) un jeton d’accès au complément (pour l’utilisateur connecté à Office). L’application hôte Office demande alors le jeton au point de terminaison Azure AD 2.0. Dans la mesure où vous avez préalablement autorisé l’hôte Office sur votre complément lors de son inscription, Azure AD enverra le jeton.
    * Si aucun utilisateur n’est connecté à Office, l’hôte Office invite l’utilisateur à se connecter.
    * Le paramètre d’options définit `forceConsent` sur `false`, donc l’utilisateur ne sera pas invité à accorder à l’hôte Office l’accès à votre complément chaque fois qu’il utilisera le complément. La première fois que l’utilisateur exécutera le complément, l’appel à `getAccessTokenAsync` échouera, mais la logique de gestion des erreurs que vous ajouterez dans une étape ultérieure effectuera automatiquement un autre appel avec le jeu d’options `forceConsent` défini sur `true`, et l’utilisateur sera invité à donner son consentement, mais uniquement la première fois.
    * Vous créerez la méthode `handleClientSideErrors` à une étape ultérieure.

    ```js
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. Remplacez TODO1 par les lignes suivantes. Vous créez la méthode `getData` et la route « /api/values » côté serveur dans les étapes suivantes. Une URL relative est utilisée pour le point de terminaison car il doit être hébergé sur le même domaine que votre complément.

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :

    * Cette méthode appelle un point de terminaison d’API Web spécifié et lui transmet le même jeton d’accès que l’application hôte Office a utilisé pour accéder à votre complément. Côté serveur, ce jeton d’accès est utilisé dans le flux « de la part de » pour obtenir un jeton d’accès à Microsoft Graph.
    * Vous créerez la méthode `handleServerSideErrors` à une étape ultérieure.

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        });
    }
    ```

### <a name="create-the-error-handling-methods"></a>Création des méthodes de gestion des erreurs

1. En dessous de la méthode `getData`, ajoutez la méthode suivante. Cette méthode gérera les erreurs dans le client du complément lorsque l’hôte Office ne parviendra pas à obtenir un jeton d’accès pour le service web du complément. Ces erreurs sont signalées avec un code d’erreur, donc la méthode utilise une instruction `switch` pour les distinguer.

    ```js
    function handleClientSideErrors(result) {

        switch (result.error.code) {

            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor.

            // TODO3: Handle the case where the user's sign-in or consent was aborted.

            // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
            //        nor Microsoft Account.

            // TODO5: Handle the case where the Office host has not been authorized to the add-in's web service or
            //        the user has not granted the service permission to their `profile`.

            // TODO6: Handle an unspecified error from the Office host.

            // TODO7: Handle the case where the Office host cannot get an access token to the add-ins
            //        web service/application.

            // TODO8: Handle the case where the user triggered an operation that calls `getAccessTokenAsync`
            //        before a previous call of it completed.

            // TODO9: Handle the case where the add-in does not support forcing consent.

            // TODO10: Log all other client errors.
        }
    }
    ```

1. Remplacez `TODO2` par le code suivant. L’erreur 13001 se produit si l’utilisateur n’est pas connecté, ou s’il a annulé, sans y répondre, une invite lui demandant d’indiquer un deuxième facteur d’authentification. Dans les deux cas, le code réexécute la méthode `getDataWithToken` et définit une option pour forcer une invite de connexion.

    ```js
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. Remplacez `TODO3` par le code suivant. L’erreur 13002 se produit lorsque la connexion ou l’octroi du consentement de l’utilisateur a été abandonné. Demandez à l’utilisateur de réessayer, mais seulement une fois.

    ```js
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. Remplacez `TODO4` par le code suivant. L’erreur 13003 se produit si l’utilisateur est connecté avec un compte qui n’est ni un compte professionnel ni un compte scolaire, ni un compte Microsoft. Demandez à l’utilisateur de se déconnecter, puis de se reconnecter avec un type de compte pris en charge.

    ```js
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > L’erreur 13004 n’est pas gérée dans cette méthode, car elle ne devrait se produire qu’en développement. Elle ne peut pas être résolue par du code d’exécution et il ne serait d’aucune utilité de la signaler à un utilisateur final.

1. Remplacez `TODO5` par le code suivant. L’erreur 13005 se produit si Office n’a pas été autorisé à accéder au service web du complément ou si l’utilisateur n’a pas accordé l’autorisation de service à son `profile`.

    ```js
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. Remplacez `TODO6` par le code suivant. L’erreur 13006 se produit lorsqu’une erreur non spécifiée indiquant que l’hôte est dans un état instable est survenue dans l’hôte Office. Demandez à l’utilisateur de redémarrer Office.

    ```js
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. Remplacez `TODO7` par le code suivant. L’erreur 13007 se produit lorsqu’un problème est survenu au niveau de l’interaction de l’hôte Office avec AAD de telle sorte que l’hôte ne peut pas obtenir de jeton d’accès pour accéder à l’application/au service Web des compléments. Il peut s’agir d’un problème temporaire de réseau. Demandez à l’utilisateur de réessayer plus tard.

    ```js
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. Remplacez `TODO8` par le code suivant. L’erreur 13008 se produit lorsque l’utilisateur a déclenché une opération qui appelle `getAccessTokenAsync` avant que la fin de l’appel précédent.

    ```js
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. Remplacez `TODO9` par le code suivant. L’erreur 13009 se produit lorsque le complément ne prend pas en charge l’obligation d’afficher une invite de consentement, mais que `getAccessTokenAsync` a été appelé avec l’option `forceConsent` définie sur `true`. Dans le cas habituel, lorsque cela se produit, le code doit automatiquement réexécuter `getAccessTokenAsync` avec l’option de consentement définie sur `false`. Toutefois, dans certains cas, l’appel de la méthode avec `forceConsent` défini sur `true` était lui-même une réponse automatique à une erreur dans un appel à la méthode avec l’option définie sur `false`. Dans ce cas, le code ne doit pas réessayer, mais il doit à la place conseiller à l’utilisateur de se déconnecter et de se reconnecter.

    ```js
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```

1. Remplacez `TODO10` par le code suivant.

    ```js
    default:
        logError(result);
        break;
    ```  


1. En dessous de la méthode `handleClientSideErrors`, ajoutez la méthode suivante. Cette méthode gérera les erreurs du service web du complément en cas de problème d’exécution du flux « de la part de » ou de problème d’obtention de données à partir de Microsoft Graph.

    ```js
    function handleServerSideErrors(result) {

        // TODO11: Parse the JSON response.

        // TODO12: Handle the case where AAD asks for an additional form of authentication.

        // TODO13: Handle missing consent and scope (permission) related issues.

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. Remplacez `TODO11` par le code suivant. Pour la plupart des erreurs `4xx` que le service web du complément transmettra du côté client du complément, une propriété **ExceptionMessage** se trouvera dans la réponse contenant le numéro d’erreur AADSTS (Azure Active Directory Secure Token Service), ainsi que d’autres données. Toutefois, lorsqu’AAD enverra un message au service web du complément pour demander un facteur d’authentification supplémentaire, le message contiendra une propriété **Claims** spéciale spécifiant (avec un numéro de code) le facteur supplémentaire nécessaire. Les API ASP.NET qui créent et envoient des réponses HTTP aux clients ne connaissent pas cette propriété **Claims**, donc ils ne l’incluent pas dans l’objet de la réponse. Le code côté serveur que vous allez créer dans une étape ultérieure y remédiera en ajoutant manuellement la valeur **Claims** à l’objet de réponse. Cette valeur sera dans la propriété **Message**, donc le code doit également analyser cette propriété.

    ```js
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. Remplacez `TODO12` par le code suivant. Tenez compte des informations suivantes :

    * L’erreur 50076 se produit lorsque Microsoft Graph exige un formulaire d’authentification supplémentaire.
    * L’hôte Office dois obtenir un nouveau jeton avec la valeur **Claims** pour l’option `authChallenge`. Cela demande à AAD d’inviter l’utilisateur à accepter tous les formulaires d’authentification requis.

    ```js
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }
    ```

1. Remplacez `TODO13` par le code suivant. Dans les prochaines étapes, vous allez remplacer les trois `TODO` dans ce code par un bloc conditionnel *interne*.

    ```js
    else if (exceptionMessage) {

        // TODO13A: Handle the case where consent has not been granted, or has been revoked.

        // TODO13B: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO13C: Handle the case where the token that the add-in's client-side sends to it's
        //          server-side is not valid because it is missing `access_as_user` scope (permission).
    }
  
    ```


1. Remplacez `TODO13A` par le code suivant. (Cette opération crée la première partie d’un bloc conditionnel *interne*.) Note sur ce code :

    * L’erreur 65001 signifie que l’utilisateur a refusé de donner l’accès à Microsoft Graph (ou que l’accès a été révoqué) pour une ou plusieurs autorisations.
    * Le complément doit obtenir un nouveau jeton avec l’option `forceConsent` définie sur `true`.

    ```js
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
       getDataWithToken({ forceConsent: true });
    }
    ```

1. Remplacez `TODO13B` par le code suivant. Tenez compte des informations suivantes :

    * L’erreur 70011 a plusieurs sens. Le problème qui importe pour ce complément est lorsque cette erreur indique qu’une étendue (autorisation) non valide a été demandée ; le code vérifie alors la description complète de l’erreur, pas seulement le numéro.
    * Le complément doit signaler l’erreur.

    ```js
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. Remplacez `TODO13C` par le code suivant. Tenez compte des informations suivantes :

    * Le code côté serveur que vous allez créer dans une étape ultérieure enverra le message `Missing access_as_user` si l’étendue (autorisation) `access_as_user` ne se trouve pas dans le jeton d’accès que le client du complément envoie à AAD pour qu’il l’utilise dans flux « de la part de ».
    * Le complément doit signaler l’erreur.

    ```js
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. Remplacez `TODO14` par le code suivant. (Cela fait partie du bloc conditionnel *externe* et doit figurer immédiatement après le crochet fermant de la structure commençant par `else if (exceptionMessage) {` et au même niveau de retrait.) Note sur ce code :

    * La bibliothèque d’identité que vous allez utiliser dans le code côté serveur (Microsoft Authentication Library, MSAL) doit garantir qu’aucun jeton expiré ou non valide n’est envoyé à Microsoft Graph. Cependant, si cela se produit, l’erreur renvoyée par Microsoft Graph au service web du complément a le code `InvalidAuthenticationToken`. Le code côté serveur que vous allez créer dans une étape ultérieure envoie ce message au client du complément.
    * Dans ce cas, le complément doit recommencer l’intégralité du processus d’authentification en réinitialisant les variables de compteur et d’indicateur, puis en appelant à nouveau la méthode de gestionnaire de boutons.

    ```js
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }
    ```

1. Remplacez `TODO15` par le code suivant.

    ```js
    else {
        logError(result);
    }
    ```

1. Enregistrez et fermez le fichier.

## <a name="code-the-server-side"></a>Code côté serveur

### <a name="configure-the-owin-middleware"></a>Configurer les intergiciels OWIN

1. Ouvrez le fichier Startup.cs à la racine du projet.

1. Ajoutez le mot clé `partial` à la déclaration de la classe de démarrage, si ce n’est pas déjà fait. Elle doit ressembler à ceci :

    `public partial class Startup`

1. Ajoutez la ligne suivante dans le corps de la méthode `Configuration`. Vous créez la méthode `ConfigureAuth` dans une étape ultérieure.

    `ConfigureAuth(app);`

1. Enregistrez et fermez le fichier.

1. Cliquez avec le bouton droit de la souris sur le dossier **App_Start**, puis sélectionnez **Ajouter > Classe**.

1. Dans la boîte de dialogue **Ajouter un nouvel élément** nommez le fichier **Startup.Auth.cs**, puis cliquez sur **Ajouter**.

1. Raccourcissez le nom de l’espace de noms dans le nouveau fichier `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Vérifiez que toutes les instructions `using` suivantes se trouvent en haut du fichier.

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Ajoutez le mot clé `partial` à la déclaration de la classe `Startup`, si ce n’est pas déjà fait. Elle doit ressembler à ceci :

    `public partial class Startup`

1. Ajoutez la méthode suivante à la classe `Startup`. Cette méthode spécifie comment l’intergiciel OWIN valide les jetons d’accès qui lui sont transmis à partir de la méthode `getData` dans le fichier Home.js côté client. Le processus d’autorisation est déclenché chaque fois qu’un point de terminaison Web API décoré avec l’attribut `[Authorize]` est appelé.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Remplacez TODO3 par les lignes suivantes. Tenez compte des informations suivantes :

    * Le code demande à OWIN de s’assurer que l’audience et l’émetteur du jeton spécifiés dans le jeton d’accès qui provient de l’hôte Office (et est transmis par l’appel côté client de `getData`) doivent correspondre aux valeurs spécifiées dans le fichier web.config.
    * Le réglage de `SaveSigninToken` sur `true` fait qu’OWIN enregistre le jeton brut à partir de l’hôte Office. Le complément en a besoin pour obtenir un jeton d’accès à Microsoft Graph avec le flux « de la part de ».
    * Les étendues ne sont pas validées par l’intergiciel OWIN. Les étendues du jeton d’accès, qui doivent inclure `access_as_user`, sont validées dans le contrôleur.

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Remplacez TODO4 par le code suivant. Tenez compte des informations suivantes :

    * La méthode `UseOAuthBearerAuthentication` est appelée au lieu de la méthode `UseWindowsAzureActiveDirectoryBearerAuthentication` plus courante, car cette dernière n’est pas compatible avec le point de terminaison Azure AD V2.
    * L’URL de découverte transmise à la méthode correspond à l’endroit où l’intergiciel OWIN obtient les instructions permettant d’obtenir la clé requise pour vérifier la signature sur le jeton d’accès reçu de l’hôte Office.

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. Enregistrez et fermez le fichier.

### <a name="create-the-apivalues-controller"></a>Créer le contrôleur /api/values

1. Ouvrez le fichier **Controllers\ValueController.cs**.

1. Vérifiez que les instructions `using` suivantes se trouvent en haut du fichier.

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

1. Juste au-dessus de la ligne qui déclare `ValuesController`, ajoutez l’attribut `[Authorize]`. Cela permet de s’assurer que votre complément exécutera le processus d’autorisation que vous avez configuré dans la dernière procédure chaque fois qu’une méthode de contrôleur est appelée. Seuls les appelants avec un jeton d’accès valide à votre complément peuvent ainsi appeler les méthodes du contrôleur.

    > [!NOTE]
    > Un service d’API Web MVC ASP.NET en production doit avoir une logique personnalisée pour le flux « de la part de » dans une ou plusieurs classes **FilterAttribute** personnalisées. Cet exemple pédagogique place la logique dans le contrôleur principal afin que l’intégralité du flux de la logique d’extraction de données et d’autorisation puisse être facilement suivie. De plus, l’exemple est cohérent avec les exemples de modèle d’autorisation dans [Exemples Azure](https://github.com/Azure-Samples/).

1. Ajoutez la méthode suivante à `ValuesController`. Vous remarquerez que la valeur renvoyée est `Task<HttpResponseMessage>` et non `Task<IEnumerable<string>>`, laquelle serait plus courante pour une méthode `GET api/values`. Il s’agit d’un effet secondaire du fait que notre logique d’autorisation personnalisée se trouvera dans le contrôleur : certaines conditions d’erreur de cette logique nécessitent qu’un objet Réponse HTTP soit envoyé au client du complément.

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

1. Remplacez `TODO1` par le code suivant pour confirmer que les étendues spécifiées dans le jeton incluent `access_as_user`.

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > Vous ne pouvez utiliser l’étendue `access_as_user` que pour autoriser l’API qui gère le flux « de la part de » pour les compléments Office. D’autres API dans votre service peuvent avoir leurs propres exigences d’étendue. Cela permet de limiter ce à quoi donnent accès les jetons acquis par Office.

1. Remplacez `TODO2` par le code suivant. Tenez compte des informations suivantes :
    * Ce code transforme le jeton d’accès brut reçu de l’hôte Office en objet `UserAssertion` qui sera transmis à une autre méthode.
    * Votre complément ne joue plus le rôle d’une ressource (ou audience) à laquelle l’hôte Office et l’utilisateur doivent accéder. Désormais, il est lui-même un client qui a besoin d’accéder à Microsoft Graph. `ConfidentialClientApplication` est l’objet de « contexte client » MSAL.
    * Le troisième paramètre du constructeur `ConfidentialClientApplication` est une URL de redirection qui n’est pas utilisée dans le flux « de la part de », mais il est recommandé d’utiliser l’URL correcte. Les quatrième et cinquième paramètres peuvent être utilisés pour définir un magasin permanent qui permettrait la réutilisation des jetons non expirés entre différentes sessions avec le complément. Cet exemple n’implémente pas un stockage permanent.
    * MSAL requiert les étendues `openid` et `offline_access` pour fonctionner, mais il génère une erreur si votre code les demande de façon redondante. Il génère également une erreur si votre code demande `profile`, qui est utilisé uniquement lorsque l’application Office hôte obtient le jeton pour l’application web de votre complément. Seul `Files.Read.All` est demandé explicitement.

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

1. Remplacez `TODO3` par le code suivant. Tenez compte des informations suivantes :

    * La méthode `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` recherchera tout d’abord dans le cache MSAL, c’est-à-dire en mémoire, un jeton d’accès correspondant. Uniquement s’il n’existe pas, elle lance le flux « de la part de » avec le point de terminaison Azure AD V2.
    * Si une authentification multifacteur est requise par la ressource MS Graph et si l’utilisateur ne l'a pas encore fournie, AAD lève une exception qui contient une propriété de revendication.
    * La valeur de la propriété Claims doit être transmise au client qui la transmettra à son tour à l’hôte Office, qui l’inclura alors dans une demande de nouveau jeton. AAD demandera à l’utilisateur d’accepter tous les formulaires d’authentification requis.
    * Les exceptions qui ne sont pas de type `MsalServiceException` ne sont intentionnellement pas capturées afin d’être propagées au client sous la forme de messages `500 Server Error`.

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

1. Remplacez `TODO3a` par le code suivant. Tenez compte des informations suivantes :

    * Si l’authentification multifacteur est requise par la ressource MS Graph et que l’utilisateur ne l'a pas encore fournie, AAD renvoie « 400 - Demande incorrecte » avec l’erreur AADSTS50076 et une propriété **Claims**. MSAL génère une exception **MsalUiRequiredException** (qui hérite de **MsalServiceException**) avec ces informations. 
    * La valeur de la propriété **Claims** doit être transmise au client qui doit la transmettre à son tour à l’hôte Office, qui l’inclut alors dans une demande de nouveau jeton. AAD demandera à l’utilisateur d’accepter tous les formulaires d’authentification requis.
    * Les API qui créent des réponses HTTP à partir d’exceptions ne connaissent pas la propriété **Claims**, donc ils ne l’incluent pas dans l’objet de la réponse. Nous devons créer manuellement un message qui l’inclut. Une propriété **Message** personnalisé, cependant, bloque la création d’une propriété **ExceptionMessage**, afin que la seule façon de communiquer l’ID d’erreur `AADSTS50076` au client est de l’ajouter à la propriété **Message** personnalisée. JavaScript dans le client devra découvrir si une réponse a une propriété **Message** ou **ExceptionMessage**, afin qu’il sache laquelle lire.
    * Le message personnalisé est au format JSON pour que le code JavaScript côté client puisse l’analyser avec des méthodes d’objet `JSON` connues.
    * Vous créerez la méthode `SendErrorToClient` à une étape ultérieure. Son deuxième paramètre est un objet **Exception**. Dans ce cas, le code transmet `null` car même l’objet **Exception** bloque l’inclusion de la propriété **Message** dans la réponse HTTP qui est générée.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. Remplacez `TODO3b` et `TODO3c` par le code suivant. Tenez compte des informations suivantes :

    * Si l’appel à AAD contenait au moins une étendue (autorisation) pour laquelle ni l’utilisateur, ni un administrateur client a consenti (ou pour laquelle le consentement a été révoqué) : AAD renverra « 400 Demande incorrecte » avec l’erreur `AADSTS65001`. MSAL génère une exception **MsalUiRequiredException** avec ces informations. Le client doit de nouveau appeler `getAccessTokenAsync` avec l’option `{ forceConsent: true }`.
    *  Si l’appel à AAD contenait au moins une étendue non reconnue par AAD, AAD renvoie « 400 Demande incorrecte » avec l’erreur `AADSTS70011`. MSAL génère une exception **MsalUiRequiredException** avec ces informations. Le client doit informer l’utilisateur.
    *  La description entière est incluse, car l’erreur 70011 est renvoyée dans d’autres conditions et elle doit être gérée dans ce complément uniquement lorsqu’elle indique une étendue non valide.
    *  L’objet **MsalUiRequiredException** est transmis à `SendErrorToClient`. Cela permet de garantir qu’une propriété **ExceptionMessage** qui contient les informations d’erreur est incluse dans la réponse HTTP.
    *  Il n’y a aucun message personnalisé, donc `null` est transmis en tant que troisième paramètre.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. Remplacez `TODO3d` par le code suivant. Vous remarquerez que le code génère de nouveau l’exception au lieu de la relayer dans une réponse HTTP personnalisée avec **HttpStatusCode.Forbidden** (401). L’effet de cette opération est l’envoi par ASP.NET de sa propre réponse HTTP avec le statut « Erreur serveur 500 ».

    ```csharp
    else
    {
        throw e;
    }  
    ```

1. Remplacez `TODO4` par le code suivant. Tenez compte des informations suivantes :

    * Les classes `GraphApiHelper` et `ODataHelper` sont définies dans les fichiers du dossier **Helpers**. La classe `OneDriveItem` est définie dans un fichier du dossier **Models**. La description détaillée de ces classes n’est pas pertinente pour l’autorisation ou l’authentification unique, elle est donc hors de portée de cet article.
    * Vous pouvez améliorer les performances en ne demandant à Microsoft Graph que les données réellement requises. Ainsi, le code utilise le paramètre de requête `$select` pour spécifier que nous ne souhaitons que la propriété de nom, et le paramètre `$top` pour spécifier que nous ne voulons que les trois premiers noms de fichier ou de dossier.
    * Si le jeton envoyé à Microsoft Graph n’est pas valide, Microsoft Graph envoie l’erreur « 401 accès non autorisé » avec le code « InvalidAuthenticationToken ». ASP.NET génère ensuite une exception **RuntimeBinderException**. C’est également ce qu’il se passe lorsque le jeton a expiré, bien que MSAL doive l’empêcher. 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);
    }
    ```

1. Remplacez `TODO5` par le code suivant. Tenez compte des informations suivantes :

    * Bien que le code ci-dessus demande uniquement la propriété *name* des éléments OneDrive, Microsoft Graph comporte toujours la propriété *eTag* pour les éléments OneDrive. Pour réduire la charge utile envoyée au client, le code ci-dessous reconstruit les résultats avec uniquement les noms d’élément.
    * La liste des trois fichiers et dossiers OneDrive est envoyée au client en tant que réponse HTTP « 200 OK ».

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames);
    return response;
    ```

1. Au-dessous de la méthode Get, ajoutez la méthode suivante. Tenez compte des informations suivantes sur ce code :  

    * La méthode communique au client les informations sur une exception côté serveur.
    * Si l’exception d’origine est transmise à la méthode, le constructeur HttpError inclura les informations de l’objet d’exception dans une propriété **ExceptionMessage**.  
    * Si `null` est transmis pour l’exception, le constructeur HttpError inclura le paramètre de message dans une propriété **Message** et aucune propriété **ExceptionMessage** ne sera présente.

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }
    ```

## <a name="run-the-add-in"></a>Exécution du complément

1. Assurez-vous que vous disposez de fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.

1. Dans Visual Studio, appuyez sur F5. PowerPoint s’ouvre et un groupe **SSO ASP.NET** se trouve sur le ruban **Accueil**.

1. Appuyez sur le bouton **Afficher le complément** dans ce groupe pour voir l’interface utilisateur du complément dans le volet de tâches.

1. Appuyez sur le bouton **Obtenir mes fichiers à partir de OneDrive**. Si vous n’êtes pas connecté à Office, vous serez invité à vous connecter.

    > [!NOTE]
    > Si vous étiez précédemment connecté à Office avec un ID différent, et si certaines applications Office sont toujours ouvertes, Office ne changera pas systématiquement votre identifiant même s’il semble l’avoir fait dans PowerPoint. Dans ce cas, l’appel vers Microsoft Graph peut échouer, ou des données de l’ID précédent peuvent être renvoyées. Afin d’éviter ce problème, veillez à *fermer toutes les autres applications Office* avant de cliquer sur **Obtenir mes fichiers à partir de OneDrive**.

1. Une fois que vous êtes connecté, la liste de vos fichiers et dossiers dans OneDrive s’affiche sous le bouton. Cette opération peut prendre plus de 15 secondes, surtout la première fois.
