---
title: Utiliser le générateur Yeoman pour créer un complément Office qui utilise l’authentification unique (aperçu)
description: Utiliser le générateur Yeoman pour créer un complément Office Node.js qui utilise l’authentification unique (aperçu)
ms.date: 01/30/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: db3567a17a01af76c9db5f859a35dba46fd4858d
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163877"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="19898-103">Utiliser le générateur Yeoman pour créer un complément Office qui utilise l’authentification unique (aperçu)</span><span class="sxs-lookup"><span data-stu-id="19898-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="19898-104">Dans cet article, vous allez découvrir le processus d’utilisation du générateur Yeoman pour créer un complément Office pour Excel, Outlook, Word ou PowerPoint qui utilise l’authentification unique (SSO) lorsque c’est possible, et utilise une autre méthode d’authentification utilisateur lorsque l’authentification unique n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="19898-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Outlook, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="19898-105">Avant d'essayer de terminer ce démarrage rapide, consultez la section [Activer l'authentification unique pour les compléments Office](../develop/sso-in-office-add-ins.md) pour apprendre les concepts de base de l'authentification unique dans les compléments Office.</span><span class="sxs-lookup"><span data-stu-id="19898-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="19898-106">Le générateur Yeoman simplifie le processus de création d’un complément d’authentification unique en automatisant les étapes nécessaires pour configurer l’authentification unique dans Azure et la génération du code nécessaire pour qu’un complément utilise l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="19898-107">Si vous souhaitez avoir une description détaillée de la procédure à suivre pour effectuer manuellement les étapes que le générateur Yeoman automatise, veuillez consulter le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="19898-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="19898-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="19898-108">Prerequisites</span></span>

* <span data-ttu-id="19898-109">[Node.js](https://nodejs.org) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="19898-109">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

* <span data-ttu-id="19898-110">La dernière version de[Yeoman](https://github.com/yeoman/yo) et de [Yeoman Générateur de compléments Office](https://github.com/OfficeDev/generator-office). Pour installer ces outils globalement, exécutez la commande suivante à partir de l’invite de commande :</span><span class="sxs-lookup"><span data-stu-id="19898-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="19898-111">Si vous utilisez un Mac et que l'interface de ligne de commande (CLI) Azure n’est pas installée sur votre ordinateur, vous devez installer [Homebrew](https://brew.sh/).</span><span class="sxs-lookup"><span data-stu-id="19898-111">If you're using a Mac and don't have the Azure CLI installed on your machine, you must install [Homebrew](https://brew.sh/).</span></span> <span data-ttu-id="19898-112">Le script de configuration de l’authentification unique exécuté lors de ce démarrage rapide utilise homebrew pour installer l’interface de ligne de commande Azure, puis utilise la CLI pour configurer l’authentification unique dans Azure.</span><span class="sxs-lookup"><span data-stu-id="19898-112">The SSO configuration script that you'll run during this quick start will use Homebrew to install the Azure CLI, and will then use the Azure CLI to configure SSO within Azure.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="19898-113">Création du projet de complément</span><span class="sxs-lookup"><span data-stu-id="19898-113">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="19898-114">Le générateur Yeoman peut créer un complément Office avec authentification unique pour Excel, Outlook, Word ou PowerPoint, et peut être créé avec des scripts de type JavaScript ou TypeScript.</span><span class="sxs-lookup"><span data-stu-id="19898-114">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Outlook, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="19898-115">Les instructions suivantes indiquent `JavaScript` et `Excel`, mais vous devez choisir le type de script et l’application client Office les mieux adaptées à votre scénario.</span><span class="sxs-lookup"><span data-stu-id="19898-115">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="19898-116">**Sélectionnez un type de projet :** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="19898-116">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="19898-117">**Sélectionnez un type de script :** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="19898-117">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="19898-118">**Comment souhaitez-vous nommer votre complément ?**</span><span class="sxs-lookup"><span data-stu-id="19898-118">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="19898-119">**Quelle application client Office voulez-vous prendre en charge ?**</span><span class="sxs-lookup"><span data-stu-id="19898-119">**Which Office client application would you like to support?**</span></span> `Excel`

![Capture d’écran des invites et des réponses relatives au générateur Yeoman](../images/yo-office-sso-excel.png)

<span data-ttu-id="19898-121">Après avoir exécuté l’assistant, le générateur crée le projet et installe les composants Node de prise en charge.</span><span class="sxs-lookup"><span data-stu-id="19898-121">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="19898-122">Explorer le projet</span><span class="sxs-lookup"><span data-stu-id="19898-122">Explore the project</span></span>

<span data-ttu-id="19898-123">Le projet de complément que vous avez créé à l’aide du générateur Yeoman contient un code pour un complément de volet Office avec authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-123">The add-in project that you've created with the Yeoman generator contains code for an SSO-enabled task pane add-in.</span></span>

- <span data-ttu-id="19898-124">Le fichier **./manifest.xml** du répertoire racine du projet définit les paramètres et fonctionnalités du complément.</span><span class="sxs-lookup"><span data-stu-id="19898-124">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="19898-125">Le fichier **./src/taskpane/taskpane.html** contient les balises HTML du volet Office.</span><span class="sxs-lookup"><span data-stu-id="19898-125">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="19898-126">Le fichier **./src/taskpane/taskpane.css** contient le style CSS appliqué au contenu du volet Office.</span><span class="sxs-lookup"><span data-stu-id="19898-126">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="19898-127">Le fichier **./src/taskpane/taskpane.js** contient le code de l’API JavaScript pour Office qui facilite l’interaction entre le volet Office et l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="19898-127">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

- <span data-ttu-id="19898-128">Le fichier **./src/helpers/documentHelper.js** utilise la bibliothèque JavaScript Office pour ajouter les données de Microsoft Graph au document Office.</span><span class="sxs-lookup"><span data-stu-id="19898-128">The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>
- <span data-ttu-id="19898-129">Le fichier **./src/helpers/fallbackauthdialog.html** est la page sans interface utilisateur qui charge le code JavaScript de la méthode d’authentification de secours.</span><span class="sxs-lookup"><span data-stu-id="19898-129">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.</span></span>
- <span data-ttu-id="19898-130">Le fichier **./src/helpers/fallbackauthdialog.html** contient le code JavaScript de la méthode d’authentification de secours qui se connecte l'utilisateur avec msal.js.</span><span class="sxs-lookup"><span data-stu-id="19898-130">The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.</span></span>
- <span data-ttu-id="19898-131">Le fichier **./SRC/helpers/fallbackauthhelper.js** contient le volet Office JavaScript qui appelle la méthode d’authentification de secours dans les scénarios lorsque l’authentification unique n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="19898-131">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.</span></span>
- <span data-ttu-id="19898-132">Le fichier **./src/helpers/ssoauthhelper.js** contient l’appel JavaScript à l’API de l’authentification unique, `getAccessToken`, reçoit le jeton d’amorçage, initialise le remplacement du jeton d’amorçage pour un jeton d’accès à Microsoft Graph et appelle Microsoft Graph pour les données.</span><span class="sxs-lookup"><span data-stu-id="19898-132">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>

- <span data-ttu-id="19898-133">Le fichier **./ENV** dans le répertoire racine du projet définit les constantes utilisées par le projet de complément.</span><span class="sxs-lookup"><span data-stu-id="19898-133">The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>
    > [!NOTE]
    > <span data-ttu-id="19898-134">Certaines des constantes définies dans ce fichier sont utilisées pour simplifier le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-134">Some of the constants defined in this file are used to facilitate the SSO process.</span></span> <span data-ttu-id="19898-135">Vous pouvez mettre à jour les valeurs de ce fichier pour qu'elles correspondent à votre scénario spécifique.</span><span class="sxs-lookup"><span data-stu-id="19898-135">You may want to update values in this file to match your specific scenario.</span></span> <span data-ttu-id="19898-136">Par exemple, vous pouvez mettre à jour ce fichier pour spécifier une autre étendue, si votre complément nécessite une autre valeur que `User.Read`.</span><span class="sxs-lookup"><span data-stu-id="19898-136">For example, you can update this file to specify a different scope, if your add-in requires something other than `User.Read`.</span></span>

## <a name="configure-sso"></a><span data-ttu-id="19898-137">Configurer l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="19898-137">Configure SSO</span></span>

<span data-ttu-id="19898-138">À ce stade, votre projet de complément a été créé et contient le code nécessaire pour simplifier le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-138">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="19898-139">Ensuite, procédez comme suit pour configurer l’authentification unique pour votre complément.</span><span class="sxs-lookup"><span data-stu-id="19898-139">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="19898-140">Accédez au dossier racine du projet.</span><span class="sxs-lookup"><span data-stu-id="19898-140">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="19898-141">Exécutez la commande suivante pour configurer l’authentification unique pour le complément.</span><span class="sxs-lookup"><span data-stu-id="19898-141">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="19898-142">Cette commande échouera si votre locataire est configuré pour nécessiter une authentification à deux facteurs.</span><span class="sxs-lookup"><span data-stu-id="19898-142">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="19898-143">Dans ce scénario, vous devez effectuer manuellement les étapes d’inscription et de configuration de l’authentification unique de l’application Azure, comme décrit dans le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="19898-143">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="19898-144">Une fenêtre de navigateur web s’ouvre et vous invite à vous connecter à Azure.</span><span class="sxs-lookup"><span data-stu-id="19898-144">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="19898-145">Connectez-vous à Azura à l’aide de vos informations d’identification d’administrateur Office 365.</span><span class="sxs-lookup"><span data-stu-id="19898-145">Sign in to Azure using your Office 365 administrator credentials.</span></span> <span data-ttu-id="19898-146">Ces informations d’identification sont utilisées pour inscrire une nouvelle application dans Azure et configurer les paramètres requis par l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-146">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="19898-147">Si vous vous connectez à Azure à l’aide d’informations d’identification non-administrateur au cours de cette étape, le script `configure-sso` ne peut pas fournir d’accord d’administrateur pour le complément aux utilisateurs au sein de votre organisation.</span><span class="sxs-lookup"><span data-stu-id="19898-147">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="19898-148">Par conséquent, l’authentification unique ne sera pas disponible pour les utilisateurs du complément. vous serez invité à vous connecter.</span><span class="sxs-lookup"><span data-stu-id="19898-148">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="19898-149">Une fois vos informations d'identification saisies, fermez la fenêtre du navigateur et revenez à l'invite de commande.</span><span class="sxs-lookup"><span data-stu-id="19898-149">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="19898-150">Au fur et à mesure du processus de configuration de l’authentification unique, les messages d’État s’affichent sur la console.</span><span class="sxs-lookup"><span data-stu-id="19898-150">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="19898-151">Comme décrit dans la section messages de la console, les fichiers du projet de complément que le générateur Yeoman a créé sont automatiquement mis à jour avec les données requises par le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-151">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="19898-152">Try it out</span><span class="sxs-lookup"><span data-stu-id="19898-152">Try it out</span></span>

<span data-ttu-id="19898-153">Si vous avez créé un complément Excel, Word ou PowerPoint, suivez les étapes décrites dans la section suivante pour le tester. Si vous avez créé un complément Outlook, suivez les étapes décrites dans la section [d'Outlook](#outlook) à la place.</span><span class="sxs-lookup"><span data-stu-id="19898-153">If you've created an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If you've created an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="19898-154">Excel, Word et PowerPoint</span><span class="sxs-lookup"><span data-stu-id="19898-154">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="19898-155">Pour tester un complément Excel, Word ou PowerPoint, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="19898-155">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="19898-156">Une fois le processus de configuration de l’authentification unique terminé, exécutez la commande suivante pour créer le projet, démarrez le serveur web local et mettez votre complément en sideload dans l’application client Office précédemment sélectionnée.</span><span class="sxs-lookup"><span data-stu-id="19898-156">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="19898-157">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="19898-157">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="19898-158">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="19898-158">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="19898-159">Dans l’application client Office qui s’ouvre lorsque vous exécutez la commande précédente (par exemple, Excel, Word ou PowerPoint), assurez-vous que vous êtes connecté avec un utilisateur membre de la même organisation Office 365 que le compte d’administrateur Office 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="19898-159">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="19898-160">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-160">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="19898-161">Dans l’application client Office, sélectionnez l’onglet **Accueil**, puis choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet Office du complément.</span><span class="sxs-lookup"><span data-stu-id="19898-161">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="19898-162">L’image ci-après illustre ce bouton dans Excel.</span><span class="sxs-lookup"><span data-stu-id="19898-162">The following image shows this button in Excel.</span></span>

    ![Bouton Complément Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="19898-164">Au bas du volet Office, sélectionnez le bouton **Obtenir mes informations de profil utilisateur** pour lancer le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-164">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="19898-165">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="19898-165">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="19898-166">Cela peut se produire lorsque l’administrateur du locataire n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Office 365 (« professionnel ou scolaire »).</span><span class="sxs-lookup"><span data-stu-id="19898-166">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="19898-167">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="19898-167">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Boîte de dialogue demande d’autorisation](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="19898-169">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="19898-169">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="19898-170">Le complément récupère les informations de profil de l’utilisateur connecté et écrit celui-ci dans le document.</span><span class="sxs-lookup"><span data-stu-id="19898-170">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="19898-171">L’image suivante montre un exemple d’informations de profil écrites dans une feuille de calcul Excel.</span><span class="sxs-lookup"><span data-stu-id="19898-171">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Informations de profil utilisateur dans la feuille de calcul Excel](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="19898-173">Outlook</span><span class="sxs-lookup"><span data-stu-id="19898-173">Outlook</span></span>

<span data-ttu-id="19898-174">Pour tester un complément Outlook, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="19898-174">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="19898-175">Une fois le processus de configuration de l’authentification unique terminé, exécutez la commande suivante pour créer le projet et démarrer le serveur web local.</span><span class="sxs-lookup"><span data-stu-id="19898-175">When the SSO configuration process completes, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="19898-176">Les compléments Office doivent utiliser le protocole HTTPS, et non HTTP, même lorsque vous développez.</span><span class="sxs-lookup"><span data-stu-id="19898-176">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="19898-177">Si vous êtes invité à installer un certificat après avoir exécuté la commande suivante, acceptez d’installer le certificat fourni par le générateur Yeoman.</span><span class="sxs-lookup"><span data-stu-id="19898-177">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="19898-178">Suivez les instructions indiquées dans l’article [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md) pour charger le complément dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="19898-178">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="19898-179">N'oubliez pas de vous connecter avec un utilisateur membre de la même organisation Office 365 que le compte d’administrateur Office 365 que vous avez utilisé pour vous connecter à Azure lors de la configuration de l’authentification unique à l’étape 3 de la [section précédente](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="19898-179">Make sure that you're signed in to Outlook with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="19898-180">Cette opération permet d’établir les conditions appropriées pour la réussite de l’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-180">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="19898-181">Rédigez un nouveau message dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="19898-181">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="19898-182">Dans la fenêtre de composition du message, choisissez le bouton **Afficher le volet Office** du ruban pour ouvrir le volet du complément.</span><span class="sxs-lookup"><span data-stu-id="19898-182">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bouton du complément Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="19898-184">Au bas du volet des tâches, sélectionnez le bouton **Obtenir mes informations de profil utilisateur** pour lancer le processus d’authentification unique.</span><span class="sxs-lookup"><span data-stu-id="19898-184">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="19898-185">Si une boîte de dialogue s’affiche pour demander des autorisations pour le compte du complément, cela signifie que l’authentification unique n’est pas prise en charge pour votre scénario et que le complément est plutôt repassé à une autre méthode d’authentification des utilisateurs.</span><span class="sxs-lookup"><span data-stu-id="19898-185">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="19898-186">Cela peut se produire lorsque l’administrateur du locataire n’a pas accordé le consentement du complément pour accéder à Microsoft Graph, ou lorsque l’utilisateur n’est pas connecté à Office à l’aide d’un compte Microsoft valide ou d’un compte Office 365 (« professionnel ou scolaire »).</span><span class="sxs-lookup"><span data-stu-id="19898-186">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="19898-187">Sélectionnez le bouton **Accepter** dans la fenêtre de boîte de dialogue pour continuer.</span><span class="sxs-lookup"><span data-stu-id="19898-187">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Boîte de dialogue demande d’autorisation](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="19898-189">Une fois qu’un utilisateur a accepté cette demande d’autorisation, il n’est plus invité à le faire à l’avenir.</span><span class="sxs-lookup"><span data-stu-id="19898-189">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="19898-190">Le complément récupère les informations du profil de l’utilisateur connecté et les écrit dans le corps de l'e-mail.</span><span class="sxs-lookup"><span data-stu-id="19898-190">The add-in retrieves profile information for the signed-in user and writes it to the body of the email message.</span></span> 

    ![Informations du profil utilisateur dans un message Outlook](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="19898-192">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="19898-192">Next steps</span></span>

<span data-ttu-id="19898-193">Félicitations, vous avez créé un complément de volet Office qui utilise l’authentification unique lorsque c’est possible, et utilise une autre méthode d’authentification utilisateur lorsque l’authentification unique n’est pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="19898-193">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="19898-194">Pour en savoir plus sur les étapes de configuration de l’authentification unique effectuées automatiquement par le générateur Yeoman et le code facilitant le processus d’authentification unique, veuillez consultez le didacticiel [Créer un complément Office Node.js qui utilise l’authentification unique](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="19898-194">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="19898-195">Consultez aussi</span><span class="sxs-lookup"><span data-stu-id="19898-195">See also</span></span>

- [<span data-ttu-id="19898-196">Activer l’authentification unique pour des compléments Office</span><span class="sxs-lookup"><span data-stu-id="19898-196">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="19898-197">Créer un complément Office Node.js qui utilise l’authentification unique</span><span class="sxs-lookup"><span data-stu-id="19898-197">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="19898-198">Résolution des problèmes de messages d’erreur pour l’authentification unique (SSO)</span><span class="sxs-lookup"><span data-stu-id="19898-198">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)