---
title: Création d’un complément Office Node.js qui utilise l’authentification unique
description: 23/01/2018
ms.openlocfilehash: a6e91b84de69e4b2da5cc10277f0ca3579287b96
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533762"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="b25f8-103">Créer un complément Office Node.js qui utilise l’authentification unique (aperçu)</span><span class="sxs-lookup"><span data-stu-id="b25f8-103">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="b25f8-p101">Les utilisateurs peuvent se connecter à Office et votre complément Web Office peut tirer parti de cette procédure de connexion pour autoriser les utilisateurs à accéder à votre complément et à Microsoft Graph sans obliger les utilisateurs à se connecter une deuxième fois. Pour obtenir une vue d’ensemble, consultez [Activer l’authentification unique pour des compléments Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b25f8-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="b25f8-106">Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré à Node.js et Express.</span><span class="sxs-lookup"><span data-stu-id="b25f8-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span> 

> [!NOTE]
> <span data-ttu-id="b25f8-107">Pour voir un article similaire sur un complément basé sur ASP.NET, reportez-vous à [Créer un complément Office ASP.NET qui utilise l’authentification unique](create-sso-office-add-ins-aspnet.md).</span><span class="sxs-lookup"><span data-stu-id="b25f8-107">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b25f8-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="b25f8-108">Prerequisites</span></span>

* <span data-ttu-id="b25f8-109">[Nœud et npm](https://nodejs.org/en/), version 6.9.4 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b25f8-109">[Node and npm](https://nodejs.org/en/), version 6.9.4 or later</span></span>

* <span data-ttu-id="b25f8-110">[Git Bash](https://git-scm.com/downloads) (ou un autre client Git)</span><span class="sxs-lookup"><span data-stu-id="b25f8-110">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="b25f8-111">TypeScript version 2.2.2 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b25f8-111">TypeScript version 2.2.2 or later</span></span>

* <span data-ttu-id="b25f8-112">Office 2016, version 1708, build 8424.nnnn ou version ultérieure (la version par abonnement Office 365, parfois appelée « Démarrer en un clic »).</span><span class="sxs-lookup"><span data-stu-id="b25f8-112">Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”)</span></span>

  <span data-ttu-id="b25f8-p102">Il vous sera peut-être demandé de participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, consultez la page [Participez au programme Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="b25f8-p102">You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="b25f8-115">Configurer le projet de démarrage</span><span class="sxs-lookup"><span data-stu-id="b25f8-115">Set up the starter project</span></span>

1. <span data-ttu-id="b25f8-116">Clonez ou téléchargez le référentiel sur [Complément Office NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span><span class="sxs-lookup"><span data-stu-id="b25f8-116">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="b25f8-117">Il existe trois versions de l’échantillon :</span><span class="sxs-lookup"><span data-stu-id="b25f8-117">There are two versions of the sample:</span></span>  
    > * <span data-ttu-id="b25f8-p103">Le dossier **Before** est un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. Les sections suivantes de cet article vous guident tout au long de la procédure d’exécution de cette dernière.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span> 
    > * <span data-ttu-id="b25f8-p104">La version **Finale** de l’échantillon s’apparente au complément que vous auriez si vous terminiez les procédures de cet article, sauf que le projet terminé comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, suivez simplement les instructions de cet article, mais remplacez « Avant » par « Finale » et ignorez les sections **Code côté client** et **Code côté serveur**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p104">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="b25f8-123">La version **mutualisée finale** est un échantillon final qui prend en charge l’architecture mutualisée.</span><span class="sxs-lookup"><span data-stu-id="b25f8-123">The **Completed Multitenant** version is a completed sample that supports multitenancy.</span></span> <span data-ttu-id="b25f8-124">Si vous avez l’intention de prendre en charge des comptes Microsoft de différents domaines avec l’authentification unique, explorez cet exemple.</span><span class="sxs-lookup"><span data-stu-id="b25f8-124">Explore this sample if you intend to support Microsoft accounts from different domains with SSO.</span></span>
    >
    > <span data-ttu-id="b25f8-125">_Quelle que soit la version que vous utilisez, vous devrez approuver un certificat pour l’hôte local. Consultez la note « IMPORTANT » dans le fichier Lisez-moi du référentiel._</span><span class="sxs-lookup"><span data-stu-id="b25f8-125">_Regardless of which version you use, you will need to trust a certificate for the localhost. See the "IMPORTANT" note in the Readme of the repo._</span></span>

2. <span data-ttu-id="b25f8-126">Ouvrez une console Git Bash dans le dossier **Before**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-126">Open a Git bash console in the **Before** folder.</span></span>

3. <span data-ttu-id="b25f8-127">Saisissez `npm install` dans la console pour installer toutes les dépendances détaillées dans le fichier package.json.</span><span class="sxs-lookup"><span data-stu-id="b25f8-127">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

4. <span data-ttu-id="b25f8-128">Saisissez `npm run build ` dans la console pour générer le projet.</span><span class="sxs-lookup"><span data-stu-id="b25f8-128">Enter `npm run build ` in the console to build the project.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="b25f8-p106">Il se peut que vous voyiez certaines erreurs de construction indiquant que certaines variables sont déclarées mais pas utilisées. Ignorez ces erreurs. Elles représentent un effet secondaire du fait qu’il manque du code dans la version « Avant » de l’échantillon. Ce code sera ajouté ultérieurement.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p106">You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="b25f8-132">Enregistrez le complément avec le point de terminaison Azure AD v2.0</span><span class="sxs-lookup"><span data-stu-id="b25f8-132">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="b25f8-133">Les instructions suivantes présentant un manière générique, vous pouvez les utiliser dans plusieurs emplacements.</span><span class="sxs-lookup"><span data-stu-id="b25f8-133">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="b25f8-134">En lien avec ce article, procédez comme suit :</span><span class="sxs-lookup"><span data-stu-id="b25f8-134">For this article do the following:</span></span>
- <span data-ttu-id="b25f8-135">Remplacez l’espace réservé **$ADD-IN-NAME$** par `“Office-Add-in-NodeJS-SSO`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-135">Replace the placeholder **$ADD-IN-NAME$** with `“Office-Add-in-NodeJS-SSO`.</span></span>
- <span data-ttu-id="b25f8-136">Remplacez l’espace réservé **$FQDN-WITHOUT-PROTOCOL$** par `localhost:3000`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-136">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:3000`.</span></span>
- <span data-ttu-id="b25f8-137">Lorsque vous spécifiez des autorisations dans la boîte de dialogue **Sélectionner les autorisations**, cochez les cases correspondant aux autorisations suivantes.</span><span class="sxs-lookup"><span data-stu-id="b25f8-137">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="b25f8-138">Votre complément proprement dit ne requiert que la première. Mais l’autorisation `profile` est également requise pour que l’hôte Office puisse obtenir un jeton pour l’application web de votre complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-138">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="b25f8-139">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="b25f8-139">Files.Read.All</span></span>
    * <span data-ttu-id="b25f8-140">profil</span><span class="sxs-lookup"><span data-stu-id="b25f8-140">profile</span></span>

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="b25f8-141">Octroi du consentement administrateur pour le complément</span><span class="sxs-lookup"><span data-stu-id="b25f8-141">Details are at: Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="b25f8-142">Configurer le complément</span><span class="sxs-lookup"><span data-stu-id="b25f8-142">Configure the add-in</span></span>

1. <span data-ttu-id="b25f8-p109">Dans votre éditeur de code, ouvrez le fichier src\server.ts. Près de la partie supérieure se trouve un appel à un constructeur d’une classe `AuthModule`. Il existe certains paramètres de chaîne dans le constructeur auxquels vous devez affecter des valeurs.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p109">In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.</span></span>

2. <span data-ttu-id="b25f8-146">Pour la propriété `client_id`, remplacez l’espace réservé `{client GUID}` par l’ID d’application que vous avez enregistré lorsque vous avez inscrit le complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-146">For the `client_id` property, replace the placeholder `{client GUID}` with the application secret that you saved when you registered the add-in.</span></span> <span data-ttu-id="b25f8-147">Lorsque vous avez terminé, vous obtenez simplement un GUID entre guillemets simples.</span><span class="sxs-lookup"><span data-stu-id="b25f8-147">When you are done, there should just be a GUID in single quotation marks.</span></span> <span data-ttu-id="b25f8-148">Il ne doit pas y avoir de caractère «{}».</span><span class="sxs-lookup"><span data-stu-id="b25f8-148">There should not be any "{}" characters.</span></span>

3. <span data-ttu-id="b25f8-149">Pour la propriété `client_secret`, remplacez l’espace réservé `{client secret}` par le secret de l’application que vous avez enregistré lorsque vous avez inscrit le complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-149">For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.</span></span>

4. <span data-ttu-id="b25f8-p111">Pour la propriété `audience`, remplacez l’espace réservé `{audience GUID}` par l’ID d’application que vous avez enregistré lorsque vous avez inscrit le complément. (La même valeur que celle affectée à la propriété `client_id`.)</span><span class="sxs-lookup"><span data-stu-id="b25f8-p111">For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)</span></span>
  
3. <span data-ttu-id="b25f8-152">Dans la chaîne affectée à la propriété `issuer`, vous verrez l’espace réservé *{O365 tenant GUID}*.</span><span class="sxs-lookup"><span data-stu-id="b25f8-152">In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*.</span></span> <span data-ttu-id="b25f8-153">Remplacez-le par l’ID de client Office 365.</span><span class="sxs-lookup"><span data-stu-id="b25f8-153">Replace this with the Office 365 tenancy ID.</span></span> <span data-ttu-id="b25f8-154">Pour obtenir de celui-ci, utilisez l’une des méthodes décrites dans [Trouver votre ID de client Office 365](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id).</span><span class="sxs-lookup"><span data-stu-id="b25f8-154">Use one of the methods in [Find your Office 365 tenant ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span> <span data-ttu-id="b25f8-155">Lorsque vous avez terminé, la valeur de la propriété `issuer` doit ressembler à ceci :</span><span class="sxs-lookup"><span data-stu-id="b25f8-155">When you are done, the `issuer` property value should look something like this:</span></span>

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. <span data-ttu-id="b25f8-156">Conservez les autres paramètres du constructeur `AuthModule` inchangés.</span><span class="sxs-lookup"><span data-stu-id="b25f8-156">Leave the other parameters in the `AuthModule` constructor unchanged.</span></span> <span data-ttu-id="b25f8-157">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b25f8-157">Save and close  the file.</span></span>

1. <span data-ttu-id="b25f8-158">Dans la racine du projet, ouvrez le fichier manifeste du complément « Office-Add-in-NodeJS-SSO.xml ».</span><span class="sxs-lookup"><span data-stu-id="b25f8-158">In the root of the project, open the add-in manifest file “Office-Add-in-NodeJS-SSO.xml”.</span></span>

1. <span data-ttu-id="b25f8-159">Faites défiler vers le bas du fichier.</span><span class="sxs-lookup"><span data-stu-id="b25f8-159">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="b25f8-160">Juste au-dessus de la balise de fin `</VersionOverrides>`, vous trouverez le balisage suivant :</span><span class="sxs-lookup"><span data-stu-id="b25f8-160">Just above the end `</VersionOverrides>` tag, you will find the following markup:</span></span>

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

1. <span data-ttu-id="b25f8-161">Remplacez l’espace réservé « {application_GUID here} » *aux deux endroits* du balisage par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-161">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="b25f8-162">(Les « {} » ne font pas partie de l’ID ; vous ne devez pas les inclure.) C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.</span><span class="sxs-lookup"><span data-stu-id="b25f8-162">(The "{}" are not part of the ID, so don't include them.) This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="b25f8-163">La valeur **Resource** correspond à l’**URI d’ID d’application** défini lorsque vous avez ajouté la plateforme d’API web à l’enregistrement du complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-163">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="b25f8-164">La section **Scopes** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via AppSource.</span><span class="sxs-lookup"><span data-stu-id="b25f8-164">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="b25f8-165">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b25f8-165">Save and close  the file.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="b25f8-166">Code côté client</span><span class="sxs-lookup"><span data-stu-id="b25f8-166">Code the client side</span></span>

1. <span data-ttu-id="b25f8-p115">Ouvrez le fichier program.js dans le dossier **public**. Il contient déjà du code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p115">Open the program.js file in the **public** folder. It already has some code in it:</span></span>

    * <span data-ttu-id="b25f8-169">Une affectation à la méthode `Office.initialize` qui affecte elle-même un gestionnaire à l’événement ClickButton `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-169">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="b25f8-170">Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.</span><span class="sxs-lookup"><span data-stu-id="b25f8-170">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="b25f8-171">Une méthode `logErrors` qui consigne dans la console les erreurs qui ne sont pas destinées à l’utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="b25f8-171">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

11. <span data-ttu-id="b25f8-p116">En dessous de l’affectation au `Office.initialize`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p116">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="b25f8-174">La gestion des erreurs dans le complément tente parfois automatiquement d’obtenir un jeton d’accès une deuxième fois, à l’aide d’un autre jeu d’options.</span><span class="sxs-lookup"><span data-stu-id="b25f8-174">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="b25f8-175">La variable de compteur `timesGetOneDriveFilesHasRun` et la variable d’indicateur `triedWithoutForceConsent` et `timesMSGraphErrorReceived` permettent de s’assurer que l’utilisateur ne tente pas de manière répétée d’obtenir un jeton sans y parvenir.</span><span class="sxs-lookup"><span data-stu-id="b25f8-175">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variables `triedWithoutForceConsent` and `timesMSGraphErrorReceived` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="b25f8-p118">Vous allez créer la méthode `getDataWithToken` à l’étape suivante, mais rappelez-vous qu’elle définit une option appelée `forceConsent` sur `false`. Vous en saurez plus à la prochaine étape.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p118">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="b25f8-p119">En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p119">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="b25f8-180">[getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) est la nouvelle API d’Office.js qui permet à un complément de demander à l’application hôte Office (Excel, PowerPoint, Word, etc.) un jeton d’accès au complément (pour l’utilisateur connecté à Office).</span><span class="sxs-lookup"><span data-stu-id="b25f8-180">The [](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="b25f8-181">L’application hôte Office demande alors le jeton au point de terminaison Azure AD 2.0.</span><span class="sxs-lookup"><span data-stu-id="b25f8-181">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="b25f8-182">Dans la mesure où vous avez préalablement autorisé l’hôte Office sur votre complément lors de son inscription, Azure AD enverra le jeton.</span><span class="sxs-lookup"><span data-stu-id="b25f8-182">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="b25f8-183">Si aucun utilisateur n’est connecté à Office, l’hôte Office invite l’utilisateur à se connecter.</span><span class="sxs-lookup"><span data-stu-id="b25f8-183">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="b25f8-184">Le paramètre d’options définit `forceConsent` sur `false`, donc l’utilisateur ne sera pas invité à accorder à l’hôte Office l’accès à votre complément chaque fois qu’il utilisera le complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-184">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="b25f8-185">La première fois que l’utilisateur exécutera le complément, l’appel à `getAccessTokenAsync` échouera, mais la logique de gestion des erreurs que vous ajouterez dans une étape ultérieure effectuera automatiquement un autre appel avec le jeu d’options `forceConsent` défini sur `true`, et l’utilisateur sera invité à donner son consentement, mais uniquement la première fois.</span><span class="sxs-lookup"><span data-stu-id="b25f8-185">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="b25f8-186">Vous créerez la méthode `handleClientSideErrors` à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b25f8-186">You will create the `handleClientSideErrors` method in a later step.</span></span>

    ```javascript
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

1. <span data-ttu-id="b25f8-p122">Remplacez TODO1 par les lignes suivantes. Vous créez la méthode `getData` et la route « /api/values » côté serveur dans les étapes suivantes. Une URL relative est utilisée pour le point de terminaison car il doit être hébergé sur le même domaine que votre complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p122">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="b25f8-p123">En dessous de la méthode `getOneDriveFiles`, ajoutez le code ci-dessous. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p123">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="b25f8-p124">Cette méthode appelle un point de terminaison d’API Web spécifié et lui transmet le même jeton d’accès que l’application hôte Office a utilisé pour accéder à votre complément. Côté serveur, ce jeton d’accès est utilisé dans le flux « de la part de » pour obtenir un jeton d’accès à Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p124">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="b25f8-194">Vous créerez la méthode `handleServerSideErrors` à une étape ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b25f8-194">You will create the `handleServerSideErrors` method in a later step.</span></span>

    ```javascript
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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="b25f8-195">Création des méthodes de gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="b25f8-195">Create the error-handling methods</span></span>

1. <span data-ttu-id="b25f8-196">En dessous de la méthode `getData`, ajoutez la méthode suivante.</span><span class="sxs-lookup"><span data-stu-id="b25f8-196">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="b25f8-197">Cette méthode gérera les erreurs dans le client du complément lorsque l’hôte Office ne parviendra pas à obtenir un jeton d’accès pour le service web du complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-197">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="b25f8-198">Ces erreurs sont signalées avec un code d’erreur, donc la méthode utilise une instruction `switch` pour les distinguer.</span><span class="sxs-lookup"><span data-stu-id="b25f8-198">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="b25f8-199">Remplacez `TODO2` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-199">Replace `TODO2` with the following code:</span></span> <span data-ttu-id="b25f8-200">L’erreur 13001 se produit si l’utilisateur n’est pas connecté, ou s’il a annulé, sans y répondre, une invite lui demandant d’indiquer un deuxième facteur d’authentification.</span><span class="sxs-lookup"><span data-stu-id="b25f8-200">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="b25f8-201">Dans les deux cas, le code réexécute la méthode `getDataWithToken` et définit une option pour forcer une invite de connexion.</span><span class="sxs-lookup"><span data-stu-id="b25f8-201">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="b25f8-202">Remplacez `TODO3` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-202">Replace `TODO3` with the following code:</span></span> <span data-ttu-id="b25f8-203">L’erreur 13002 se produit lorsque la connexion ou l’octroi du consentement de l’utilisateur a été abandonné.</span><span class="sxs-lookup"><span data-stu-id="b25f8-203">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="b25f8-204">Demandez à l’utilisateur de réessayer, mais seulement une fois.</span><span class="sxs-lookup"><span data-stu-id="b25f8-204">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="b25f8-205">Remplacez `TODO4` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-205">Replace `TODO4` with the following code:</span></span> <span data-ttu-id="b25f8-206">L’erreur 13003 se produit si l’utilisateur est connecté avec un compte qui n’est ni un compte professionnel ni un compte scolaire, ni un compte Microsoft.</span><span class="sxs-lookup"><span data-stu-id="b25f8-206">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Micrososoft Account.</span></span> <span data-ttu-id="b25f8-207">Demandez à l’utilisateur de se déconnecter, puis de se reconnecter avec un type de compte pris en charge.</span><span class="sxs-lookup"><span data-stu-id="b25f8-207">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="b25f8-208">Les erreurs 13004 et 13005 ne sont pas gérées dans cette méthode, car elles ne devraient se produire qu’en développement.</span><span class="sxs-lookup"><span data-stu-id="b25f8-208">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="b25f8-209">Elles ne peuvent pas être résolues par du code d’exécution et il ne serait d’aucune utilité de les signaler à un utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="b25f8-209">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="b25f8-p130">Remplacez `TODO5` par le code suivant. L’erreur 13006 se produit lorsqu’une erreur non spécifiée indiquant que l’hôte est dans un état instable est survenue dans l’hôte Office. Demandez à l’utilisateur de redémarrer Office.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p130">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="b25f8-213">Remplacez `TODO6` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-213">Replace `TODO6` with the following code:</span></span> <span data-ttu-id="b25f8-214">L’erreur 13007 se produit lorsqu’un problème est survenu au niveau de l’interaction de l’hôte Office avec AAD de telle sorte que l’hôte ne peut pas obtenir de jeton d’accès pour accéder à l’application/au service Web des compléments.</span><span class="sxs-lookup"><span data-stu-id="b25f8-214">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="b25f8-215">Il peut s’agir d’un problème temporaire de réseau.</span><span class="sxs-lookup"><span data-stu-id="b25f8-215">This may be a temporary network issue.</span></span> <span data-ttu-id="b25f8-216">Demandez à l’utilisateur de réessayer plus tard.</span><span class="sxs-lookup"><span data-stu-id="b25f8-216">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="b25f8-p132">Remplacez `getAccessTokenAsync` par le code suivant. L’erreur 13008 se produit lorsque l’utilisateur a déclenché une opération qui appelle `TODO7` avant la fin de l’appel précédent.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p132">Replace `TODO7` with the following code. Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="b25f8-219">Remplacez `TODO8` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-219">Replace `TODO8` with the following code:</span></span> <span data-ttu-id="b25f8-220">L’erreur 13009 se produit lorsque le complément ne prend pas en charge l’obligation d’afficher une invite de consentement, mais que `getAccessTokenAsync` a été appelé avec l’option `forceConsent` définie sur `true`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-220">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="b25f8-221">Dans le cas habituel, lorsque cela se produit, le code doit automatiquement réexécuter `getAccessTokenAsync` avec l’option de consentement définie sur `false`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-221">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="b25f8-222">Toutefois, dans certains cas, l’appel de la méthode avec `forceConsent` défini sur `true` était lui-même une réponse automatique à une erreur dans un appel à la méthode avec l’option définie sur `false`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-222">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="b25f8-223">Dans ce cas, le code ne doit pas réessayer, mais il doit à la place conseiller à l’utilisateur de se déconnecter et de se reconnecter.</span><span class="sxs-lookup"><span data-stu-id="b25f8-223">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="b25f8-224">Remplacez `TODO9` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-224">Replace `TODO9` with the following code:</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. <span data-ttu-id="b25f8-p134">En dessous de la méthode `handleClientSideErrors`, ajoutez la méthode suivante. Cette méthode gérera les erreurs du service web du complément en cas de problème d’exécution du flux « de la part de » ou de problème d’obtention de données à partir de Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p134">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. <span data-ttu-id="b25f8-p135">Remplacez `TODO10` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p135">Replace `TODO10` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="b25f8-p136">Il existe des configurations d’Azure Active Directory où l’on demande à l’utilisateur de fournir un ou plusieurs facteurs d’authentification supplémentaires pour accéder à certaines cibles Microsoft Graph (par exemple, OneDrive), même si l’utilisateur peut se connecter à Office par un simple mot de passe. Dans ce cas, AAD enverra, avec l’erreur 50076, une réponse comportant la propriété `Claims`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p136">There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response, with error 50076, that has a `Claims` property.</span></span> 
    * <span data-ttu-id="b25f8-231">L’hôte Office dois obtenir un nouveau jeton avec la valeur **Claims** pour l’option `authChallenge`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-231">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="b25f8-232">Cela demande à AAD d’inviter l’utilisateur à accepter tous les formulaires d’authentification requis.</span><span class="sxs-lookup"><span data-stu-id="b25f8-232">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. <span data-ttu-id="b25f8-p138">Remplacez `TODO11` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p138">Replace `TODO11` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="b25f8-235">L’erreur 65001 signifie que l’utilisateur a refusé de donner l’accès à Microsoft Graph (ou que l’accès a été révoqué) pour une ou plusieurs autorisations.</span><span class="sxs-lookup"><span data-stu-id="b25f8-235">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="b25f8-236">Le complément doit obtenir un nouveau jeton avec l’option `forceConsent` définie sur `true`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-236">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
        // getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="b25f8-p139">Remplacez `TODO12` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p139">Replace `TODO12` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="b25f8-239">L’erreur 70011 signifie qu’une portée (autorisation) non valide a été demandée.</span><span class="sxs-lookup"><span data-stu-id="b25f8-239">Error 70011 means that an invalid scope (permission) has been requested.</span></span> <span data-ttu-id="b25f8-240">Le complément doit signaler l’erreur.</span><span class="sxs-lookup"><span data-stu-id="b25f8-240">The add-in should report the error.</span></span>
    * <span data-ttu-id="b25f8-241">Le code consigne les autres erreurs avec un numéro d’erreur AAD.</span><span class="sxs-lookup"><span data-stu-id="b25f8-241">The code logs any other error with an AAD error number.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="b25f8-p141">Remplacez `TODO13` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p141">Replace `TODO13` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="b25f8-244">Le code côté serveur que vous créerez à une étape ultérieure enverra le message qui se termine par `... expected access_as_user` si l’étendue (autorisation) `access_as_user` ne se trouve pas dans le jeton d’accès que le client du complément envoie à AAD, afin qu’il soit utilisé dans le flux « de la part de ».</span><span class="sxs-lookup"><span data-stu-id="b25f8-244">Server-side code that you create in a later step will send the message that ends with `... expected access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="b25f8-245">Le complément doit signaler l’erreur.</span><span class="sxs-lookup"><span data-stu-id="b25f8-245">The add-in should report the error.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="b25f8-p142">Remplacez `TODO14` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*. Tenez compte des remarques suivantes à propos de ce code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p142">Replace `TODO14` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="b25f8-248">Il est peu probable qu’un jeton expiré ou non valide soit envoyé à Microsoft Graph. Cependant, si cela se produit, le code côté serveur que vous créerez dans une étape ultérieure se terminera par la chaîne `Microsoft Graph error`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-248">It is unlikely that an expired or invalid token will be sent to Microsoft Graph; but if it does happen, the server-side code that you will create in a later step will end with the string `Microsoft Graph error`.</span></span>
    * <span data-ttu-id="b25f8-249">Dans ce cas, le complément doit recommencer l’intégralité du processus d’authentification en réinitialisant les variables de compteur `timesGetOneDriveFilesHasRun` et d’indicateur `timesGetOneDriveFilesHasRun`, puis en appelant à nouveau la méthode de gestionnaire de boutons.</span><span class="sxs-lookup"><span data-stu-id="b25f8-249">In this case, the add-in should start the entire authentication process over by resetting the `timesGetOneDriveFilesHasRun` counter and `timesGetOneDriveFilesHasRun` flag variables, and then re-calling the button handler method.</span></span> <span data-ttu-id="b25f8-250">Toutefois, il ne doit faire cela qu’une seule fois.</span><span class="sxs-lookup"><span data-stu-id="b25f8-250">But it should do this only once.</span></span> <span data-ttu-id="b25f8-251">Si l’erreur se produit à nouveau, il doit simplement la consigner.</span><span class="sxs-lookup"><span data-stu-id="b25f8-251">If it happens again, it should just log the error.</span></span>
    * <span data-ttu-id="b25f8-252">Le code consigne l’erreur si elle se produit deux fois de suite.</span><span class="sxs-lookup"><span data-stu-id="b25f8-252">The code logs the error if it happens twice in succession.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }        
    }
    ```

1. <span data-ttu-id="b25f8-253">Remplacez `TODO15` par le code suivant *juste en dessous de la dernière accolade fermante du code que vous avez ajouté à l’étape précédente*.</span><span class="sxs-lookup"><span data-stu-id="b25f8-253">Replace `TODO15` with the following code *just below the last closing brace of the code you added in the previous step*.</span></span>

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a><span data-ttu-id="b25f8-254">Code côté serveur</span><span class="sxs-lookup"><span data-stu-id="b25f8-254">Code the server side</span></span>

<span data-ttu-id="b25f8-255">Il existe deux fichiers côté serveur qui doivent être modifiés.</span><span class="sxs-lookup"><span data-stu-id="b25f8-255">There are two server-side files that need to be modified.</span></span> 
- <span data-ttu-id="b25f8-p144">Le fichier src\auth.js fournit des fonctions d’assistance pour l’autorisation. Il dispose déjà des membres génériques qui sont utilisés dans une variété de flux d’autorisation. Nous devons ajouter des fonctions qui implémentent le flux « de la part de ».</span><span class="sxs-lookup"><span data-stu-id="b25f8-p144">The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.</span></span>
- <span data-ttu-id="b25f8-p145">Le fichier src\server.js possède les membres de base requis pour exécuter un serveur et les intergiciels express. Nous devons y ajouter des fonctions qui servent la page d’accueil et une API Web pour obtenir des données Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p145">The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.</span></span>

### <a name="create-a-method-to-exchange-tokens"></a><span data-ttu-id="b25f8-261">Créer une méthode pour échanger des jetons</span><span class="sxs-lookup"><span data-stu-id="b25f8-261">Create a method to exchange tokens</span></span>

1. <span data-ttu-id="b25f8-p146">Ouvrez le fichier \src\auth.ts. Ajoutez la méthode ci-après à la classe `AuthModule`. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p146">Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="b25f8-p147">Le paramètre `jwt` est le jeton d’accès à l’application. Dans le flux « de la part de », il est échangé avec AAD contre un jeton d’accès à la ressource.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p147">The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.</span></span>
    * <span data-ttu-id="b25f8-267">Le paramètre scopes a une valeur par défaut, mais dans cet exemple, elle sera remplacée par le code appelant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-267">The scopes parameter has a default value, but in this sample it will be overridden by the calling code.</span></span>
    * <span data-ttu-id="b25f8-p148">Le paramètre de ressource est facultatif. Il ne doit pas être utilisé lorsque le STS est le point de terminaison AAD V 2.0. Le point de terminaison AAD V 2.0 déduit la ressource des étendues et renvoie une erreur si une ressource est envoyée dans la requête HTTP.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p148">The resource parameter is optional. It should not be used when the STS is the AAD V 2.0 endpoint. The V 2.0 endpoint infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request.</span></span> 
    * <span data-ttu-id="b25f8-271">La génération d’une exception dans le bloc `catch` ne provoquera *pas* l’envoi immédiat du message « 500 Erreur interne du serveur » au client.</span><span class="sxs-lookup"><span data-stu-id="b25f8-271">Throwing an exception in the `catch` block will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="b25f8-272">L’appel de code dans le fichier server.js interceptera cette exception et la convertira en un message d’erreur qui sera envoyé au client.</span><span class="sxs-lookup"><span data-stu-id="b25f8-272">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the 
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                            + JSON.stringify(exception), 
                                            exception);
            }
        }
        ```

2. <span data-ttu-id="b25f8-p150">Remplacez `TODO3` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p150">Replace `TODO3` with the following code. About this code, note:</span></span>
    * <span data-ttu-id="b25f8-p151">Un STS qui prend en charge le flux « de la part de » attend certaines paires de propriété/valeur dans le corps de la requête HTTP. Ce code construit un objet qui devient le corps de la requête.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p151">An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request.</span></span> 
    * <span data-ttu-id="b25f8-277">Une propriété de ressource est ajoutée au corps si, et uniquement si, une ressource a été transmise à la méthode.</span><span class="sxs-lookup"><span data-stu-id="b25f8-277">A resource property is added to the body if, and only if, a resource was passed to the method.</span></span>

        ```typescript
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

3. <span data-ttu-id="b25f8-278">Remplacez `TODO4` par le code suivant, qui envoie la requête HTTP au point de terminaison de jeton du STS.</span><span class="sxs-lookup"><span data-stu-id="b25f8-278">Replace `TODO4` with the following code which sends the HTTP request to the token endpoint of the STS.</span></span>

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. <span data-ttu-id="b25f8-279">Remplacez `TODO5` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-279">Replace `TODO5` with the following code:</span></span> <span data-ttu-id="b25f8-280">Vous remarquerez que la génération d’une exception ne provoquera *pas* l’envoi immédiat d’un message « 500 Erreur interne du serveur » au client.</span><span class="sxs-lookup"><span data-stu-id="b25f8-280">Note that throwing an exception will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="b25f8-281">L’appel de code dans le fichier server.js interceptera cette exception et la convertira en un message d’erreur qui sera envoyé au client.</span><span class="sxs-lookup"><span data-stu-id="b25f8-281">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. <span data-ttu-id="b25f8-p153">Remplacez `TODO6` par le code suivant. Vous remarquerez que le code prolonge le jeton d’accès à la ressource et son délai d’expiration, en plus de le renvoyer. Le code d’appel permet d’éviter les appels inutiles au STS en réutilisant un jeton d’accès non expiré à la ressource. Vous verrez comment procéder dans la section suivante.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p153">Replace `TODO6` with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.</span></span>

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. <span data-ttu-id="b25f8-286">Enregistrez le fichier, mais ne le fermez pas.</span><span class="sxs-lookup"><span data-stu-id="b25f8-286">Save the file but don't close it yet.</span></span>

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a><span data-ttu-id="b25f8-287">Créer une méthode pour accéder à la ressource à l’aide du flux « de la part de »</span><span class="sxs-lookup"><span data-stu-id="b25f8-287">Create a method to get access to the resource using the "on behalf of" flow</span></span>

1. <span data-ttu-id="b25f8-p154">Toujours dans src/auth.ts, ajoutez la méthode ci-après à la classe `AuthModule`. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p154">Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="b25f8-290">Les commentaires ci-dessus concernant les paramètres de la méthode `exchangeForToken` s’appliquent aussi aux paramètres de cette méthode.</span><span class="sxs-lookup"><span data-stu-id="b25f8-290">The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.</span></span>
    * <span data-ttu-id="b25f8-p155">La méthode recherche d’abord dans le stockage permanent un jeton d’accès à la ressource qui n’a pas expiré et qui ne va pas expirer dans la minute qui suit. Il appelle la méthode `exchangeForToken` que vous avez créée dans la dernière section uniquement si nécessaire.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p155">The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.</span></span>

    ```typescript
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

2. <span data-ttu-id="b25f8-293">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b25f8-293">Save and close  the file.</span></span>

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a><span data-ttu-id="b25f8-294">Créer les points de terminaison que serviront la page d’accueil et les données du complément</span><span class="sxs-lookup"><span data-stu-id="b25f8-294">Create the endpoints that will serve the add-in's home page and data</span></span>

1. <span data-ttu-id="b25f8-295">Ouvrez le fichier src\server.ts.</span><span class="sxs-lookup"><span data-stu-id="b25f8-295">Open the src\server.ts file.</span></span> 

2. <span data-ttu-id="b25f8-p156">Ajoutez la méthode suivante au bas du fichier. Cette méthode servira la page d’accueil du complément. Le manifeste du complément spécifie l’URL de la page d’accueil.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p156">Add the following method to the bottom of the file. This method will serve the add-in's home page. The add-in manifest specifies the home page URL.</span></span>

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. <span data-ttu-id="b25f8-p157">Ajoutez la méthode suivante en bas du fichier. Cette méthode traite toutes les requêtes concernant l’API `onedriveitems`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p157">Add the following method to bottom of the file. This method will handle any requests for the `onedriveitems` API.</span></span>
    ```typescript
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. <span data-ttu-id="b25f8-301">Remplacez `TODO7` par le code suivant, qui valide le jeton d’accès reçu à partir de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="b25f8-301">Replace TODO7 with the following code which validates the access token received from the Office host application.</span></span> <span data-ttu-id="b25f8-302">La méthode `verifyJWT` est définie dans le fichier src\auth.ts.</span><span class="sxs-lookup"><span data-stu-id="b25f8-302">The `verifyJWT` method is defined in the src\auth.ts file.</span></span> <span data-ttu-id="b25f8-303">Elle valide toujours l’audience et l’émetteur.</span><span class="sxs-lookup"><span data-stu-id="b25f8-303">It always validates the audience and the issuer.</span></span> <span data-ttu-id="b25f8-304">Nous utilisons le paramètre facultatif pour spécifier que nous souhaitons également vérifier que l’étendue du jeton d’accès est `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="b25f8-304">We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`.</span></span> <span data-ttu-id="b25f8-305">C’est la seule autorisation d’accès au complément dont l’utilisateur et l’hôte Office ont besoin pour obtenir un jeton d’accès à Microsoft Graph au moyen du flux « de la part de ».</span><span class="sxs-lookup"><span data-stu-id="b25f8-305">This is the only permisison to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf flow".</span></span> 

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > <span data-ttu-id="b25f8-306">Vous ne pouvez utiliser l’étendue `access_as_user` que pour autoriser l’API qui gère le flux « de la part de » pour les compléments Office. D’autres API dans votre service peuvent avoir leurs propres exigences d’étendue.</span><span class="sxs-lookup"><span data-stu-id="b25f8-306">Note: You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office add-ins. Other APIs in your service should have their own scope requirements.</span></span> <span data-ttu-id="b25f8-307">Cela permet de limiter ce à quoi donnent accès les jetons acquis par Office.</span><span class="sxs-lookup"><span data-stu-id="b25f8-307">This limits what can be accessed with the tokens that Office acquires.</span></span>

5. <span data-ttu-id="b25f8-p160">Remplacez `TODO8` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p160">Replace `TODO8` with the following code. Note the following about this code:</span></span>

    * <span data-ttu-id="b25f8-310">L’appel vers `acquireTokenOnBehalfOf` ne comprend pas de paramètre de ressource, étant donné que nous avons construit l’objet `AuthModule` (`auth`) avec le point de terminaison AAD V2.0 qui ne prend pas en charge une propriété de ressource.</span><span class="sxs-lookup"><span data-stu-id="b25f8-310">The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.</span></span>
    * <span data-ttu-id="b25f8-311">Le deuxième paramètre de l’appel spécifie les autorisations dont le complément aura besoin pour obtenir une liste des fichiers et dossiers de l’utilisateur dans OneDrive.</span><span class="sxs-lookup"><span data-stu-id="b25f8-311">The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive.</span></span> <span data-ttu-id="b25f8-312">(L’autorisation `profile` n’est pas demandée, car elle n’est nécessaire qu’au moment où l’hôte Office obtient le jeton d’accès à votre complément, pas lorsque vous travaillez dans ce jeton pour un jeton d’accès à Microsoft Graph.)</span><span class="sxs-lookup"><span data-stu-id="b25f8-312">(The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)</span></span>

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. <span data-ttu-id="b25f8-p162">Remplacez `TODO9` par la ligne suivante. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p162">Replace `TODO9` with the following line. Note the following about this code:</span></span>

    * <span data-ttu-id="b25f8-315">La classe MSGraphHelper est définie dans src\msgraph-helper.ts.</span><span class="sxs-lookup"><span data-stu-id="b25f8-315">The MSGraphHelper class is defined in src\msgraph-helper.ts.</span></span> 
    * <span data-ttu-id="b25f8-316">Nous réduisons les données qui doivent être renvoyées en spécifiant que nous ne souhaitons que la propriété name et uniquement les 3 premiers éléments.</span><span class="sxs-lookup"><span data-stu-id="b25f8-316">We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.</span></span>

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

7. <span data-ttu-id="b25f8-317">Remplacez `TODO10` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-317">Replace `TODO10` with the following code:</span></span> <span data-ttu-id="b25f8-318">Notez que ce code gère les erreurs « 401 Non autorisé » de Microsoft Graph qui signalent un jeton expiré ou non valide.</span><span class="sxs-lookup"><span data-stu-id="b25f8-318">Note that this code handles '401 Unauthorized" errors from Microsoft Graph which would indicate an expired or invalid token.</span></span> <span data-ttu-id="b25f8-319">Il est très peu probable que cela se produise, car la logique persistante du jeton doit empêcher ces erreurs.</span><span class="sxs-lookup"><span data-stu-id="b25f8-319">It is very unlikely that this would ever happen since the token persisting logic should prevent it.</span></span> <span data-ttu-id="b25f8-320">(Reportez-vous à la section **Créer une méthode pour accéder à la ressource à l’aide du flux « de la part de »** ci-dessus.) Si cela se produit, ce code communiquera l’erreur au client avec, dans le nom de l’erreur, « Microsoft Graph error ».</span><span class="sxs-lookup"><span data-stu-id="b25f8-320">(See the section **Create a method to get access to the resource using the "on behalf of" flow** above.) If it does happen, this code will relay the error to the client with "Microsoft Graph error" in the error name.</span></span> <span data-ttu-id="b25f8-321">(Reportez-vous à la méthode `handleClientSideErrors` que vous avez créée dans le fichier program.js dans une étape précédente.) Le code que vous ajouterez au fichier ODataHelper.js à une étape ultérieure vous permet de traiter les erreurs provenant de Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b25f8-321">(See the `handleClientSideErrors` method that you created in the program.js file in an earlier step.) Code that you add to the ODataHelper.js file in a later step helps process errors from Microsoft Graph.</span></span>

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. <span data-ttu-id="b25f8-p164">Remplacez `TODO11` par le code suivant. Notez que Microsoft Graph renvoie des métadonnées OData et une propriété **eTag** pour chaque élément, même si `name` est la seule propriété demandée. Le code envoie uniquement les noms d’éléments au client.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p164">Replace `TODO11` with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.</span></span>

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. <span data-ttu-id="b25f8-325">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b25f8-325">Save and close  the file.</span></span>

### <a name="add-response-handling-to-the-odatahelper"></a><span data-ttu-id="b25f8-326">Ajouter une gestion des réponses à ODataHelper</span><span class="sxs-lookup"><span data-stu-id="b25f8-326">Add response handling to the ODataHelper</span></span>

1. <span data-ttu-id="b25f8-327">Ouvrez le fichier src\odata-helper.ts.</span><span class="sxs-lookup"><span data-stu-id="b25f8-327">Open the file src\odata-helper.ts.</span></span> <span data-ttu-id="b25f8-328">Le fichier est presque complet.</span><span class="sxs-lookup"><span data-stu-id="b25f8-328">The file is almost complete.</span></span> <span data-ttu-id="b25f8-329">Il manquant le corps du rappel au gestionnaire pour l’événement de « fin » de demande.</span><span class="sxs-lookup"><span data-stu-id="b25f8-329">What's missing is the body of the callback to the handler for the request "end" event.</span></span> <span data-ttu-id="b25f8-330">Remplacez `TODO` par le code suivant.</span><span class="sxs-lookup"><span data-stu-id="b25f8-330">Replace the `TODO` with the following code.</span></span> <span data-ttu-id="b25f8-331">Tenez compte des informations suivantes sur ce code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-331">About this code note:</span></span>

    * <span data-ttu-id="b25f8-332">La réponse du point de terminaison OData peut-être une erreur, supposons une erreur 401 si le point de terminaison nécessite un jeton d’accès et que celui-ci n’est pas valide ou a expiré.</span><span class="sxs-lookup"><span data-stu-id="b25f8-332">The response from the OData endpoint might be an error, say a 401 if the endpoint requires an access token and it was invalid or expired.</span></span> <span data-ttu-id="b25f8-333">Cependant, un message d’erreur reste un *message*, pas une erreur dans l’appel de `https.get`, donc la ligne `on('error', reject)` à la fin de `https.get` n’est pas déclenchée.</span><span class="sxs-lookup"><span data-stu-id="b25f8-333">But an error message is still a *message*, not an error in the call of `https.get`, so the `on('error', reject)` line at the end of `https.get` isn't triggered.</span></span> <span data-ttu-id="b25f8-334">Par conséquent, le code distingue les messages de réussite (200) des messages d’erreur, et envoie un objet JSON à l’appelant soit les informations d’erreur, soit avec les informations demandées.</span><span class="sxs-lookup"><span data-stu-id="b25f8-334">So, the code distinguishes success (200) messages from error messages and sends a JSON object to the caller with either the requested OData or error information.</span></span>

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  <span data-ttu-id="b25f8-p167">Remplacez `TODO1` par le code suivant. Notez que le code suppose que les données sont renvoyées au format JSON.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p167">Replace `TODO1` with the following code. Note that the code assumes the data is returned as JSON.</span></span>

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  <span data-ttu-id="b25f8-p168">Remplacez `TODO2` par le code suivant. Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="b25f8-p168">Replace `TODO2` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="b25f8-339">Une réponse d’erreur d’une source OData aura toujours un code d’état (statusCode) et généralement un message d’état (statusMessage).</span><span class="sxs-lookup"><span data-stu-id="b25f8-339">An error response from an OData source will always have a statusCode and usually a statusMessage.</span></span> <span data-ttu-id="b25f8-340">Certaines sources OData ajoutent également une propriété d’erreur au corps avec des informations supplémentaires, telles qu’un message et un code internes, ou plus spécifiques.</span><span class="sxs-lookup"><span data-stu-id="b25f8-340">Some OData sources also add an error property to the body with further information, such as an inner, or more specific, code and message.</span></span>
    * <span data-ttu-id="b25f8-341">L’objet de promesse est résolu, pas rejeté.</span><span class="sxs-lookup"><span data-stu-id="b25f8-341">The Promise object is resolved, not rejected.</span></span> <span data-ttu-id="b25f8-342">`https.get` s’exécute quand un service web appelle un point de terminaison OData de serveur à serveur.</span><span class="sxs-lookup"><span data-stu-id="b25f8-342">The `https.get` runs when a web service calls an OData endpoint server-to-server.</span></span> <span data-ttu-id="b25f8-343">Cependant, cet appel s’inscrit dans le contexte d’un appel d’un client à une API web dans le service web.</span><span class="sxs-lookup"><span data-stu-id="b25f8-343">But that call comes in the context of a call from a client to a web API in the web service.</span></span> <span data-ttu-id="b25f8-344">La demande « externe » du client au service web n’aboutit jamais si cette demande « interne » est rejetée.</span><span class="sxs-lookup"><span data-stu-id="b25f8-344">The "outer" request from the client to the web service never completes if this "inner" request is rejected.</span></span> <span data-ttu-id="b25f8-345">De plus, la résolution de la requête avec l’objet `Error` personnalisé est obligatoire si l’émetteur de l’appel `http.get` doit communiquer les erreurs du point de terminaison OData au client.</span><span class="sxs-lookup"><span data-stu-id="b25f8-345">Also, resolving the request with the custom `Error` object is required if the caller of `http.get` needs to relay errors from the OData endpoint to the client.</span></span>

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;
    
    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. <span data-ttu-id="b25f8-346">Enregistrez et fermez le fichier.</span><span class="sxs-lookup"><span data-stu-id="b25f8-346">Save and close  the file.</span></span>

## <a name="deploy-the-add-in"></a><span data-ttu-id="b25f8-347">Déploiement du complément</span><span class="sxs-lookup"><span data-stu-id="b25f8-347">Deploy the add-in</span></span>

<span data-ttu-id="b25f8-348">Vous devez maintenant indiquer à Office où trouver le complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-348">Now you need to let Office know where to find the add-in.</span></span>

1. <span data-ttu-id="b25f8-349">Créez un partage réseau, ou [partagez un dossier sur le réseau](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span><span class="sxs-lookup"><span data-stu-id="b25f8-349">Create a network share, or [share a folder to the network](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span></span>

2. <span data-ttu-id="b25f8-350">Placez une copie du fichier manifeste Office-Add-in-NodeJS-SSO.xml, depuis la racine du projet, dans le dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="b25f8-350">Place a copy of the Office-Add-in-NodeJS-SSO.xml manifest file, from the root of the project, into the shared folder.</span></span>

3. <span data-ttu-id="b25f8-351">Lancez PowerPoint et ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="b25f8-351">Launch Word and open a document.</span></span>

4. <span data-ttu-id="b25f8-352">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-352">Choose the **File** tab, and then choose **Options**.</span></span>

5. <span data-ttu-id="b25f8-353">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-353">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

6. <span data-ttu-id="b25f8-354">Choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-354">Choose **Trusted Add-ins Catalogs**.</span></span>

7. <span data-ttu-id="b25f8-355">Dans le champ **URL du catalogue**, saisissez le chemin réseau permettant d’accéder au partage de dossier qui contient le fichier Office-Add-in-NodeJS-SSO.xml, puis sélectionnez **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-355">In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-NodeJS-SSO.xml, and then choose **Add Catalog**.</span></span>

8. <span data-ttu-id="b25f8-356">Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-356">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

9. <span data-ttu-id="b25f8-p171">Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage de Microsoft Office. Fermez PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p171">A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close PowerPoint.</span></span>

## <a name="build-and-run-the-project"></a><span data-ttu-id="b25f8-359">Création et exécution du projet</span><span class="sxs-lookup"><span data-stu-id="b25f8-359">Build and run the project</span></span>

<span data-ttu-id="b25f8-p172">Il existe deux manières de créer et d’exécuter le projet selon que vous utilisez Visual Studio Code. Pour les deux façons, le projet est généré et reconstruit automatiquement, puis ré-exécuté lorsque vous apportez des modifications au code.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p172">There are two ways to build and run the project depending on whether you are using Visual Studio Code. For both ways, the project builds and automatically rebuilds and reruns when you make changes to the code.</span></span>

1. <span data-ttu-id="b25f8-362">Si vous n’utilisez pas Visual Studio Code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-362">If you are not using Visual Studio Code:</span></span> 
 1. <span data-ttu-id="b25f8-363">Ouvrez un terminal de nœud et accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="b25f8-363">Open a node terminal and navigate to the root folder of the project.</span></span>
 2. <span data-ttu-id="b25f8-364">Dans le terminal, entrez **npm run build**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-364">In the terminal, enter **npm run build**.</span></span> 
 3. <span data-ttu-id="b25f8-365">Ouvrez un second terminal de nœud et accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="b25f8-365">Open a second node terminal and navigate to the root folder of the project.</span></span>
 4. <span data-ttu-id="b25f8-366">Dans le terminal, entrez **npm run start**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-366">In the terminal, enter **npm run start**.</span></span>

2. <span data-ttu-id="b25f8-367">Si vous utilisez VS Code :</span><span class="sxs-lookup"><span data-stu-id="b25f8-367">If you are using VS Code:</span></span>
 1. <span data-ttu-id="b25f8-368">Ouvrez le projet dans VS Code.</span><span class="sxs-lookup"><span data-stu-id="b25f8-368">Open the project in VS Code.</span></span>
 2. <span data-ttu-id="b25f8-369">Appuyez sur CTRL-MAJ-B pour générer le projet.</span><span class="sxs-lookup"><span data-stu-id="b25f8-369">Press CTRL-SHIFT-B to build the project.</span></span>
 3. <span data-ttu-id="b25f8-370">Appuyez sur F5 pour exécuter le projet dans une session de débogage.</span><span class="sxs-lookup"><span data-stu-id="b25f8-370">Press F5 to run the project in a debugging session.</span></span>


## <a name="add-the-add-in-to-an-office-document"></a><span data-ttu-id="b25f8-371">Ajouter le complément à un document Office</span><span class="sxs-lookup"><span data-stu-id="b25f8-371">Add the add-in to an Office document</span></span>

1. <span data-ttu-id="b25f8-372">Redémarrez PowerPoint et ouvrez ou créez une présentation.</span><span class="sxs-lookup"><span data-stu-id="b25f8-372">Restart PowerPoint and open or create a presentation.</span></span>

1. <span data-ttu-id="b25f8-373">Si l’onglet **Développeur** n’est pas visible sur le ruban, activez-le en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="b25f8-373">If the **Developer** tab is not visible on the ribbon, enable it with the following steps:</span></span>
 1. <span data-ttu-id="b25f8-374">Accédez à **Fichier** | **Options** | **Personnaliser le ruban**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-374">Navigate to **File** | **Options** | **Customize Ribbon**.</span></span>
 2. <span data-ttu-id="b25f8-375">Cliquez sur la case à cocher pour activer **Développeur** dans l’arborescence des noms de contrôle dans la partie droite de la page **Personnaliser le ruban**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-375">Click the check box to enable **Developer** in the tree of control names on the right of the **Customize Ribbon** page.</span></span>
 3. <span data-ttu-id="b25f8-376">Appuyez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-376">Press **OK**.</span></span>

2. <span data-ttu-id="b25f8-377">Sous l’onglet **Développeur** de PowerPoint, choisissez **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-377">On the **Developer** tab in PowerPoint, choose **My Add-ins**.</span></span>

3. <span data-ttu-id="b25f8-378">Sélectionnez l’onglet **DOSSIER PARTAGÉ**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-378">Select the **SHARED FOLDER** tab.</span></span>

4. <span data-ttu-id="b25f8-379">Choisissez **Échantillon SSO NodeJS**, puis sélectionnez **OK**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-379">Choose **SSO NodeJS Sample**, and then select **OK**.</span></span>

5. <span data-ttu-id="b25f8-380">Dans le ruban **Accueil**, un nouveau groupe appelé **SSO NodeJS** apparaît avec un bouton intitulé **Afficher le complément** et une icône.</span><span class="sxs-lookup"><span data-stu-id="b25f8-380">On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="b25f8-381">Test du complément</span><span class="sxs-lookup"><span data-stu-id="b25f8-381">Test the add-in</span></span>

1. <span data-ttu-id="b25f8-382">Assurez-vous que vous disposez de fichiers dans votre espace OneDrive afin de pouvoir vérifier les résultats.</span><span class="sxs-lookup"><span data-stu-id="b25f8-382">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

2. <span data-ttu-id="b25f8-383">Cliquez sur le bouton **Afficher le complément** pour ouvrir le complément.</span><span class="sxs-lookup"><span data-stu-id="b25f8-383">Click **Show Add-in** button to open the add-in.</span></span>

2. <span data-ttu-id="b25f8-p173">Le complément s’ouvre avec une page d’accueil. Cliquez sur le bouton **Obtenir mes fichiers à partir de OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p173">The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.</span></span>

2. <span data-ttu-id="b25f8-p174">Si vous êtes connecté à Office, une liste de vos fichiers et dossiers sur OneDrive apparaîtront en dessous du bouton. La première fois, l’opération peut prendre plus de 15 secondes.</span><span class="sxs-lookup"><span data-stu-id="b25f8-p174">If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.</span></span>

3. <span data-ttu-id="b25f8-388">Si vous n’êtes pas connecté à Office, une fenêtre contextuelle s’ouvre et vous invite à vous connecter.</span><span class="sxs-lookup"><span data-stu-id="b25f8-388">If you are not signed into Office, a popup will open and prompt you to sign in.</span></span> <span data-ttu-id="b25f8-389">Une fois que vous êtes connecté, la liste de vos fichiers et dossiers s’affiche après quelques secondes.</span><span class="sxs-lookup"><span data-stu-id="b25f8-389">If you are not signed into Office, a popup will open and prompt you to sign in. After you have completed the sign-in, the list of your files and folders will appear after a few seconds. You do not press the button a second time.</span></span> <span data-ttu-id="b25f8-390">*N’appuyez pas sur le bouton une deuxième fois.*</span><span class="sxs-lookup"><span data-stu-id="b25f8-390">*You should not press the button a second time.*</span></span>

> [!NOTE]
> <span data-ttu-id="b25f8-391">Si vous étiez précédemment connecté à Office avec un ID différent, et si certaines applications Office sont toujours ouvertes, Office ne changera pas systématiquement votre identifiant même s’il semble l’avoir fait dans PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b25f8-391">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="b25f8-392">Dans ce cas, l’appel vers Microsoft Graph peut échouer, ou des données de l’ID précédent peuvent être renvoyées.</span><span class="sxs-lookup"><span data-stu-id="b25f8-392">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="b25f8-393">Afin d’éviter ce problème, veillez à *fermer toutes les autres applications Office* avant de cliquer sur **Obtenir mes fichiers à partir de OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="b25f8-393">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
