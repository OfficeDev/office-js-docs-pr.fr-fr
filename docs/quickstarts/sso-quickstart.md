---
title: Utiliser le générateur Yeoman pour créer un complément Office qui utilise l’authentification unique (aperçu)
description: Utiliser le générateur Yeoman pour créer un complément Office Node.js qui utilise l’authentification unique (aperçu)
ms.date: 01/27/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: d3a78a99574c92d0066003f0e39e835563f473cd
ms.sourcegitcommit: 413f163729183994de61a8281685184b377ef76c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/28/2020
ms.locfileid: "41571394"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="f1c75-103">Utiliser le générateur Yeoman pour créer un complément Office qui utilise l’authentification unique (aperçu)</span><span class="sxs-lookup"><span data-stu-id="f1c75-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="f1c75-104">Dans cet article, vous allez découvrir le processus d’utilisation du générateur Yeoman pour créer un complément Office pour Excel, Outlook, Word ou PowerPoint qui utilise l’authentification unique (SSO) lorsque c’est possible, et utilise une autre méthode d’authentification utilisateur lorsque l’authentification unique n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="f1c75-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Outlook, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="f1c75-105">Avant d'essayer de terminer ce démarrage rapide, consultez la section [Activer l'authentification unique pour les compléments Office](../develop/sso-in-office-add-ins.md) pour apprendre les concepts de base de l'authentification unique dans les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="f1c75-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="f1c75-106">Le générateur Yeoman simplifie le processus de création d’un complément d’authentification unique en automatisant les étapes nécessaires pour configurer l’authentification unique dans Azure et la génération du code nécessaire pour qu’un complément utilise l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="f1c75-107">Si vous souhaitez avoir une description détaillée de la procédure à suivre pour effectuer manuellement les étapes que le générateur Yeoman automatise, veuillez consulter le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="f1c75-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f1c75-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="f1c75-108">Prerequisites</span></span>

* <span data-ttu-id="f1c75-109">[Node.js](https://nodejs.org) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="f1c75-109">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

* <span data-ttu-id="f1c75-110">La dernière version de[Yeoman](https://github.com/yeoman/yo) et de [Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :</span><span class="sxs-lookup"><span data-stu-id="f1c75-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="f1c75-111">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="f1c75-111">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="f1c75-112">Le générateur Yeoman peut créer un complément Office avec authentification unique pour Excel, Outlook, Word ou PowerPoint, et peut être créé avec des scripts de type JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="f1c75-112">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Outlook, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="f1c75-113">Les instructions suivantes indiquent `JavaScript` et `Excel`, mais vous devez choisir le type de script et l’application client Office les mieux adaptées à votre scénario.</span><span class="sxs-lookup"><span data-stu-id="f1c75-113">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="f1c75-114">**Sélectionnez un type de projet :** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="f1c75-114">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="f1c75-115">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="f1c75-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="f1c75-116">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="f1c75-116">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="f1c75-117">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="f1c75-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-sso-excel.png)

<span data-ttu-id="f1c75-119">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="f1c75-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="f1c75-120">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="f1c75-120">Explore the project</span></span>

<span data-ttu-id="f1c75-121">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un code pour un complément de volet Office avec authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-121">The add-in project that you've created with the Yeoman generator contains code for an SSO-enabled task pane add-in.</span></span>

- <span data-ttu-id="f1c75-122">Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="f1c75-122">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="f1c75-123">Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="f1c75-123">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="f1c75-124">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="f1c75-124">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="f1c75-125">Le fichier **./src/taskpane/taskpane.js** contient le code de l’API JavaScript pour Office qui facilite l’interaction entre le volet Office et l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="f1c75-125">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

- <span data-ttu-id="f1c75-126">Le fichier **./src/helpers/documentHelper.js** utilise la bibliothèque JavaScript Office pour ajouter les données de Microsoft Graph au document Office.</span><span class="sxs-lookup"><span data-stu-id="f1c75-126">The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>
- <span data-ttu-id="f1c75-127">Le fichier **./src/helpers/fallbackauthdialog.html** est la page sans interface utilisateur qui charge le code JavaScript de la méthode d’authentification de secours.</span><span class="sxs-lookup"><span data-stu-id="f1c75-127">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.</span></span>
- <span data-ttu-id="f1c75-128">Le fichier **./src/helpers/fallbackauthdialog.html** contient le code JavaScript de la méthode d’authentification de secours qui se connecte l'utilisateur avec msal.js.</span><span class="sxs-lookup"><span data-stu-id="f1c75-128">The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.</span></span>
- <span data-ttu-id="f1c75-129">Le fichier **./SRC/helpers/fallbackauthhelper.js** contient le volet Office JavaScript qui appelle la méthode d’authentification de secours dans les scénarios lorsque l’authentification unique n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="f1c75-129">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.</span></span>
- <span data-ttu-id="f1c75-130">Le fichier **./src/helpers/ssoauthhelper.js** contient l’appel JavaScript à l’API de l’authentification unique, `getAccessToken`, reçoit le jeton d’amorçage, initialise le remplacement du jeton d’amorçage pour un jeton d’accès à Microsoft Graph et appelle Microsoft Graph pour les données.</span><span class="sxs-lookup"><span data-stu-id="f1c75-130">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>

- <span data-ttu-id="f1c75-131">Le fichier **./ENV** dans le répertoire racine du projet définit les constantes utilisées par le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="f1c75-131">The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>
    > [!NOTE]
    > <span data-ttu-id="f1c75-132">Certaines des constantes définies dans ce fichier sont utilisées pour simplifier le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-132">Some of the constants defined in this file are used to facilitate the SSO process.</span></span> <span data-ttu-id="f1c75-133">Vous pouvez mettre à jour les valeurs de ce fichier pour qu'elles correspondent à votre scénario spécifique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-133">You may want to update values in this file to match your specific scenario.</span></span> <span data-ttu-id="f1c75-134">Par exemple, vous pouvez mettre à jour ce fichier pour spécifier une autre étendue, si votre complément nécessite une autre valeur que `User.Read`.</span><span class="sxs-lookup"><span data-stu-id="f1c75-134">For example, you can update this file to specify a different scope, if your add-in requires something other than `User.Read`.</span></span>

## <a name="configure-sso"></a><span data-ttu-id="f1c75-135">Configurer l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="f1c75-135">Configure SSO</span></span>

<span data-ttu-id="f1c75-136">À ce stade, votre projet de complément a été créé et contient le code nécessaire pour simplifier le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-136">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="f1c75-137">Ensuite, procédez comme suit pour configurer l’authentification unique pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="f1c75-137">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="f1c75-138">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="f1c75-138">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="f1c75-139">Exécutez la commande suivante pour configurer l’authentification unique pour le complément.</span><span class="sxs-lookup"><span data-stu-id="f1c75-139">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="f1c75-140">Cette commande échouera si votre locataire est configuré pour nécessiter une authentification à deux facteurs.</span><span class="sxs-lookup"><span data-stu-id="f1c75-140">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="f1c75-141">Dans ce scénario, vous devez effectuer manuellement les étapes d’inscription et de configuration de l’authentification unique de l’application Azure, comme décrit dans le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="f1c75-141">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="f1c75-142">Une fenêtre de navigateur web s’ouvre et vous invite à vous connecter à Azure.</span><span class="sxs-lookup"><span data-stu-id="f1c75-142">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="f1c75-143">Connectez-vous à Azura à l’aide de vos informations d’identification d’administrateur Office 365.</span><span class="sxs-lookup"><span data-stu-id="f1c75-143">Sign in to Azure using your Office 365 administrator credentials.</span></span> <span data-ttu-id="f1c75-144">Ces informations d’identification sont utilisées pour inscrire une nouvelle application dans Azure et configurer les paramètres requis par l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-144">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f1c75-145">Si vous vous connectez à Azure à l’aide d’informations d’identification non-administrateur au cours de cette étape, le script `configure-sso` ne peut pas fournir d’accord d’administrateur pour le complément aux utilisateurs au sein de votre organisation.</span><span class="sxs-lookup"><span data-stu-id="f1c75-145">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="f1c75-146">Par conséquent, l’authentification unique ne sera pas disponible pour les utilisateurs du complément. vous serez invité à vous connecter.</span><span class="sxs-lookup"><span data-stu-id="f1c75-146">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="f1c75-147">Une fois vos informations d'identification saisies, fermez la fenêtre du navigateur et revenez à l'invite de commande.</span><span class="sxs-lookup"><span data-stu-id="f1c75-147">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="f1c75-148">Au fur et à mesure du processus de configuration de l’authentification unique, les messages d’État s’affichent sur la console.</span><span class="sxs-lookup"><span data-stu-id="f1c75-148">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="f1c75-149">Comme décrit dans la section messages de la console, les fichiers du projet de complément que le générateur Yeoman a créé sont automatiquement mis à jour avec les données requises par le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-149">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="f1c75-150">Try it out</span><span class="sxs-lookup"><span data-stu-id="f1c75-150">Try it out</span></span>

<span data-ttu-id="f1c75-151">Si vous avez créé un complément Excel, Word ou PowerPoint, suivez les étapes décrites dans la section suivante pour le tester. Si vous avez créé un complément Outlook, suivez les étapes décrites dans la section [d'Outlook](#outlook) à la place.</span><span class="sxs-lookup"><span data-stu-id="f1c75-151">If you've created an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If you've created an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="f1c75-152">Excel, Word et PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f1c75-152">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="f1c75-153">Pour tester un complément Excel, Word ou PowerPoint, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="f1c75-153">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="f1c75-154">Une fois le processus de configuration de l’authentification unique terminé, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et mettez votre complément en sideload dans l’application client Office précédemment sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="f1c75-154">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f1c75-155">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="f1c75-155">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f1c75-156">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f1c75-156">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="f1c75-157">Dans l’application client Office qui s’ouvre lorsque vous exécutez la commande précédente (par exemple, Excel, Word ou PowerPoint), assurez-vous que vous êtes connecté avec un utilisateur membre de la même organisation Office 365 que le compte d’administrateur Office 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="f1c75-157">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="f1c75-158">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-158">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="f1c75-159">Dans l’application client Office, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="f1c75-159">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="f1c75-160">L’image ci-après illustre ce bouton dans Excel.</span><span class="sxs-lookup"><span data-stu-id="f1c75-160">The following image shows this button in Excel.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="f1c75-162">Au bas du volet Office, sélectionnez le bouton **Obtenir mes informations de profil utilisateur** pour lancer le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-162">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="f1c75-163">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="f1c75-163">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="f1c75-164">Cela peut se produire lorsque l’administrateur du locataire n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Office 365 (« professionnel ou scolaire »).</span><span class="sxs-lookup"><span data-stu-id="f1c75-164">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="f1c75-165">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="f1c75-165">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Boîte de dialogue demande d’autorisation](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="f1c75-167">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="f1c75-167">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="f1c75-168">Le complément récupère les informations de profil de l’utilisateur connecté et écrit celui-ci dans le document.</span><span class="sxs-lookup"><span data-stu-id="f1c75-168">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="f1c75-169">L’image suivante montre un exemple d’informations de profil écrites dans une feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="f1c75-169">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Informations de profil utilisateur dans la feuille de calcul Excel](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="f1c75-171">Outlook</span><span class="sxs-lookup"><span data-stu-id="f1c75-171">Outlook</span></span>

<span data-ttu-id="f1c75-172">Pour tester un complément Outlook, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="f1c75-172">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="f1c75-173">Une fois le processus de configuration de l’authentification unique terminé, exécutez la commande suivante pour créer le projet et démarrer le serveur web local.</span><span class="sxs-lookup"><span data-stu-id="f1c75-173">When the SSO configuration process completes, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f1c75-174">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="f1c75-174">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f1c75-175">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f1c75-175">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="f1c75-176">Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](/outlook/add-ins/sideload-outlook-add-ins-for-testing) pour charger le complément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="f1c75-176">Follow the instructions in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="f1c75-177">N'oubliez pas de vous connecter avec un utilisateur membre de la même organisation Office 365 que le compte d’administrateur Office 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="f1c75-177">Make sure that you're signed in to Outlook with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="f1c75-178">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-178">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="f1c75-179">Rédigez un nouveau message dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="f1c75-179">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="f1c75-180">Dans la fenêtre de composition du message, choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet du complément.</span><span class="sxs-lookup"><span data-stu-id="f1c75-180">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton du complément Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="f1c75-182">Au bas du volet des tâches, sélectionnez le bouton **Obtenir mes informations de profil utilisateur** pour lancer le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="f1c75-182">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="f1c75-183">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="f1c75-183">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="f1c75-184">Cela peut se produire lorsque l’administrateur du locataire n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Office 365 (« professionnel ou scolaire »).</span><span class="sxs-lookup"><span data-stu-id="f1c75-184">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="f1c75-185">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="f1c75-185">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Boîte de dialogue demande d’autorisation](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="f1c75-187">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="f1c75-187">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="f1c75-188">Le complément récupère les informations du profil de l’utilisateur connecté et les écrit dans le corps de l'e-mail.</span><span class="sxs-lookup"><span data-stu-id="f1c75-188">The add-in retrieves profile information for the signed-in user and writes it to the body of the email message.</span></span> 

    ![Informations du profil utilisateur dans un message Outlook](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="f1c75-190">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="f1c75-190">Next steps</span></span>

<span data-ttu-id="f1c75-191">Félicitations, vous avez créé un complément de volet Office qui utilise l’authentification unique lorsque c’est possible, et utilise une autre méthode d’authentification utilisateur lorsque l’authentification unique n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="f1c75-191">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="f1c75-192">Pour en savoir plus sur les étapes de configuration de l’authentification unique effectuées automatiquement par le générateur Yeoman et le code facilitant le processus d’authentification unique, veuillez consultez le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="f1c75-192">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="f1c75-193">Consultez aussi</span><span class="sxs-lookup"><span data-stu-id="f1c75-193">See also</span></span>

- [<span data-ttu-id="f1c75-194">Activer l’authentification unique pour des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f1c75-194">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="f1c75-195">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="f1c75-195">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="f1c75-196">Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="f1c75-196">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)