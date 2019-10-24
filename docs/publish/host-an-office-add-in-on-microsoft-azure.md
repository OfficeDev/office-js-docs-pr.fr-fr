---
title: Héberger un complément pour Office sur Microsoft Azure | Microsoft Docs
description: Découvrez comment déployer une application web de complément sur Azure et charger une version test du complément pour le tester dans une application cliente Office.
ms.date: 10/16/2019
localization_priority: Priority
ms.openlocfilehash: 0cfddacf48bda9ed7b63d4018e3ae0437f15bcd9
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626977"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="af8b3-103">Héberger un complément pour Office sur Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="af8b3-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="af8b3-p101">Le complément Office le plus simple est constitué d’un fichier manifeste XML et d’une page HTML. Le fichier manifeste XML décrit les caractéristiques du complément, telles que son nom, les applications clientes Office dans lesquelles il peut s’exécuter et l’URL de la page HTML du complément. La page HTML est contenue dans une application web avec laquelle les utilisateurs interagissent lorsqu’ils installent et exécutent votre complément au sein d’une application cliente Office. Vous pouvez héberger l’application web d’un complément Office sur n’importe quelle plateforme d’hébergement web, y compris Azure.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="af8b3-108">Cet article décrit comment déployer une application web de complément sur Azure et [charger une version test du complément](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) pour le tester dans une application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="af8b3-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="af8b3-109">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="af8b3-109">Prerequisites</span></span> 

1. <span data-ttu-id="af8b3-110">Installez [Visual Studio 2019](https://www.visualstudio.com/downloads) et choisissez d’inclure la charge de travail de **développement Azure**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-110">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="af8b3-111">Si vous avez déjà installé Visual Studio 2019, [utilisez le programme d’installation Visual Studio Installer](/visualstudio/install/modify-visual-studio) pour vous assurer que la charge de travail de **développement Azure** est installée.</span><span class="sxs-lookup"><span data-stu-id="af8b3-111">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="af8b3-112">Installation d’Office.</span><span class="sxs-lookup"><span data-stu-id="af8b3-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="af8b3-113">Si vous n’avez pas encore Office, vous pouvez vous [inscrire pour obtenir un essai gratuit d’un mois](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span><span class="sxs-lookup"><span data-stu-id="af8b3-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="af8b3-114">Obtenez un abonnement Azure.</span><span class="sxs-lookup"><span data-stu-id="af8b3-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="af8b3-115">Si vous n’avez pas encore d’abonnement Azure, vous pouvez [en obtenir un dans le cadre de votre abonnement Visual Studio](https://azure.microsoft.com/fr-FR/pricing/member-offers/visual-studio-subscriptions/) ou vous [inscrire pour obtenir une version d’évaluation gratuite](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="af8b3-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/fr-FR/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="af8b3-116">Étape 1 : Créer un dossier partagé pour héberger le fichier manifeste XML de votre complément</span><span class="sxs-lookup"><span data-stu-id="af8b3-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="af8b3-117">Ouvrez l’explorateur de fichiers sur votre ordinateur de développement.</span><span class="sxs-lookup"><span data-stu-id="af8b3-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="af8b3-118">Cliquez avec le bouton droit de la souris sur le lecteur C:\, puis choisissez **Nouveau** > **Dossier**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="af8b3-119">Nommez le nouveau dossier AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="af8b3-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="af8b3-120">Cliquez avec le bouton droit de la souris sur le dossier AddinManifests, puis choisissez **Partager avec** > **Des personnes spécifiques**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="af8b3-121">Dans **Partage de fichiers**, sélectionnez la flèche déroulante vers le bas, puis choisissez **Tout le monde** > **Ajouter** > **Partager**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="af8b3-p102">Dans cette procédure, vous utilisez un partage de fichiers local en tant que catalogue approuvé où vous allez stocker le fichier manifeste XML du complément. Dans un scénario réel, vous pouvez choisir de [déployer le fichier manifeste XML dans un catalogue SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou de [publier le complément dans AppSource](/office/dev/store/submit-to-the-office-store), à la place.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="af8b3-124">Étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés</span><span class="sxs-lookup"><span data-stu-id="af8b3-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="af8b3-125">Démarrez Word et créez un document.</span><span class="sxs-lookup"><span data-stu-id="af8b3-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="af8b3-126">Bien que cet exemple utilise Word, vous pouvez utiliser n’importe quelle application Office qui prend en charge des compléments Office comme Excel, Outlook, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="af8b3-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="af8b3-127">Choisissez **Fichier**  >  **Options**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="af8b3-128">Dans la boîte de dialogue **Options Word**, choisissez **Centre de gestion de la confidentialité**, puis **Paramètres du Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="af8b3-p103">Dans la boîte de dialogue **Centre de gestion de la confidentialité**, choisissez **Catalogues de compléments approuvés**. Saisissez le chemin d’accès UNC (Universal Naming Convention) pour le partage de fichiers que vous avez créé précédemment en tant qu’**URL du catalogue** (par exemple, \\\NomDeVotreOrdinateur\AddinManifests), puis choisissez **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="af8b3-131">Activez la case **Afficher dans le menu**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="af8b3-132">Lorsque vous stockez un fichier manifeste XML de complément sur un partage qui est défini comme un catalogue de compléments web approuvés, le complément apparaît sous **Dossier partagé** dans la boîte de dialogue **Compléments Office** lorsque l’utilisateur accède à l’onglet **Insérer** dans le ruban et choisit **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="af8b3-133">Fermez Word.</span><span class="sxs-lookup"><span data-stu-id="af8b3-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a><span data-ttu-id="af8b3-134">Étape 3 : Créer une application web dans Azure à l’aide du Portail Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="af8b3-134">Step 3: Create a web app in Azure using the Azure portal</span></span>

<span data-ttu-id="af8b3-135">Pour créer l’application web à l’aide du portail Azure, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="af8b3-135">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="af8b3-136">Connectez-vous au [portail Azure](https://portal.azure.com/) à l’aide de vos informations d’identification Azure.</span><span class="sxs-lookup"><span data-stu-id="af8b3-136">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="af8b3-137">Sous **Azure services**, sélectionnez \*\*Applications web \*\*.</span><span class="sxs-lookup"><span data-stu-id="af8b3-137">Under **Azure Services** select **Web Apps**.</span></span>

3. <span data-ttu-id="af8b3-138">Dans la page **Service d’applications**, sélectionnez **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-138">On the **App Service** page, select **Add**.</span></span> <span data-ttu-id="af8b3-139">Fournissez ces informations :</span><span class="sxs-lookup"><span data-stu-id="af8b3-139">Provide this information:</span></span>

      - <span data-ttu-id="af8b3-140">Choisissez l’**abonnement** à utiliser pour créer ce site.</span><span class="sxs-lookup"><span data-stu-id="af8b3-140">Choose the **Subscription** to use for creating this site.</span></span>
      
      - <span data-ttu-id="af8b3-p105">Choisissez le **groupe de ressources** pour votre site. Si vous créez un groupe, vous devez également le nommer.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p105">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
      
      - <span data-ttu-id="af8b3-143">Entrez un **nom d’application** unique pour votre site.</span><span class="sxs-lookup"><span data-stu-id="af8b3-143">Enter a unique **App name** for your site.</span></span> <span data-ttu-id="af8b3-144">Azure vérifie que le nom du site est unique dans le domaine apps.net azureweb.</span><span class="sxs-lookup"><span data-stu-id="af8b3-144">Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="af8b3-145">Indiquez si vous souhaitez publier à l'aide d'un code ou d'un conteneur docker.</span><span class="sxs-lookup"><span data-stu-id="af8b3-145">Choose whether to publish using code or a docker container.</span></span>

      - <span data-ttu-id="af8b3-146">Spécifiez une **pile d’exécution**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-146">Specify a **Runtime stack**.</span></span>

      - <span data-ttu-id="af8b3-147">Choisissez le **système d’exploitation** de votre site.</span><span class="sxs-lookup"><span data-stu-id="af8b3-147">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="af8b3-148">Choisissez une **Région**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-148">Choose a geographical  **Region** appropriate for you.</span></span>

      - <span data-ttu-id="af8b3-149">Choisissez le **plan de service d’applications** à utiliser pour créer ce site.</span><span class="sxs-lookup"><span data-stu-id="af8b3-149">Choose the **App Service plan** to use for creating this site.</span></span>

      - <span data-ttu-id="af8b3-150">Sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-150">Choose **Create**.</span></span>

4. <span data-ttu-id="af8b3-151">La page suivante vous indique que votre déploiement est en cours et quand il prend fin.</span><span class="sxs-lookup"><span data-stu-id="af8b3-151">The next page will let you know that your deployment is underway and when it completes.</span></span> <span data-ttu-id="af8b3-152">Une fois l’opération terminée, sélectionnez **Accéder à la ressource**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-152">When it is completed, select **Go to resource**.</span></span>  

5. <span data-ttu-id="af8b3-153">Dans la section **Vue d’ensemble**, choisissez l’URL qui est affichée sous **URL**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-153">In the **Overview** section, choose the URL that is displayed under **URL**.</span></span> <span data-ttu-id="af8b3-154">Votre navigateur s’ouvre et affiche une page web avec le message « Votre application Service d’applications est opérationnelle. »</span><span class="sxs-lookup"><span data-stu-id="af8b3-154">Your browser opens and displays a webpage with the message "Your App Service app is up and running."</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="af8b3-155">Les sites web Azure [!include[HTTPS guidance](../includes/https-guidance.md)] fournissent automatiquement un point de terminaison HTTPS.</span><span class="sxs-lookup"><span data-stu-id="af8b3-155">[!include[HTTPS guidance](../includes/https-guidance.md)] Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="af8b3-156">Étape 4 : Créer un complément Office dans Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="af8b3-156">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="af8b3-157">Démarrez Visual Studio en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="af8b3-157">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="af8b3-158">Choisissez **Créer un nouveau projet**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-158">\*\*\*\* Create a new project</span></span>

3. <span data-ttu-id="af8b3-159">À l’aide de la zone de recherche, entrez **complément**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-159">Using the search box, enter **add-in**.</span></span>

4. <span data-ttu-id="af8b3-160">Choisissez **Complément Word web** comme type de projet, puis cliquez sur **Suivant** pour accepter les paramètres par défaut.</span><span class="sxs-lookup"><span data-stu-id="af8b3-160">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>

<span data-ttu-id="af8b3-161">Visual Studio crée un complément Word de base que vous pourrez publier tel quel, sans apporter de modifications à son projet web.</span><span class="sxs-lookup"><span data-stu-id="af8b3-161">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span> <span data-ttu-id="af8b3-162">Pour créer un complément pour un autre type d’hôte Office (par exemple, Excel), répétez les étapes et choisissez un type de projet avec l’hôte Office souhaité.</span><span class="sxs-lookup"><span data-stu-id="af8b3-162">To make an add-in for a different Office host type, such as Excel, repeat the steps and choose a project type with your desired Office host.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="af8b3-163">Étape 5 : Publier votre application web de complément Office sur Azure</span><span class="sxs-lookup"><span data-stu-id="af8b3-163">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="af8b3-164">Avec votre projet de complément ouvert dans Visual Studio, développez le nœud de solutions dans **Explorateur de solutions**, puis sélectionnez **Service d’applications**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-164">With your add-in project open in Visual Studio, expand the solution node in Solution Explorer so that you see both projects for the solution.</span></span>

2. <span data-ttu-id="af8b3-p110">Cliquez avec le bouton droit de la souris sur le projet web, puis choisissez **Publier**. Le projet web contient les fichiers d’application web du complément Office, et il s’agit donc du projet que vous publiez sur Azure.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p110">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="af8b3-167">Sur l’onglet **Publier** :</span><span class="sxs-lookup"><span data-stu-id="af8b3-167">On the **Publish** tab:</span></span>

      - <span data-ttu-id="af8b3-168">Choisissez **Microsoft Azure Application Service**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-168">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="af8b3-169">Choisissez **Sélectionner**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-169">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="af8b3-170">Choisissez **Publier**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-170">Choose **Publish**.</span></span>

4. <span data-ttu-id="af8b3-p111">Visual Studio publie le projet web pour votre complément Office sur votre site web Azure. Une fois le projet web publié par Visual Studio, votre navigateur s’ouvre et affiche une page web avec le texte « Votre application de service d’application a été créée. » Il s’agit de la page active par défaut pour l’application web.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p111">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

5. <span data-ttu-id="af8b3-174">Copiez l’URL racine (par exemple : https://YourDomain.azurewebsites.net) ; vous en aurez besoin lorsque vous modifierez le fichier manifeste de complément plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="af8b3-174">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="af8b3-175">Étape 6 : Modifier et déployer le fichier manifeste XML</span><span class="sxs-lookup"><span data-stu-id="af8b3-175">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="af8b3-176">Dans Visual Studio avec l’exemple de complément Office ouvert dans l’**explorateur de solutions**, développez la solution pour que les deux projets s’affichent.</span><span class="sxs-lookup"><span data-stu-id="af8b3-176">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="af8b3-p112">Développez le projet macro complémentaire Office (par exemple WordWebAddIn), le dossier manifeste d’avec le bouton droit de la souris et sélectionnez **Ouvrir**. Le fichier manifeste XML du complément s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p112">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="af8b3-p113">Dans le fichier manifeste XML, recherchez et remplacez toutes les instances de « ~remoteAppUrl » par l’URL racine de l’application web du complément sur Azure. Il s’agit de l’URL que vous avez copiée précédemment une fois que vous avez publié l’application web du complément sur Azure (par exemple : https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="af8b3-p113">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="af8b3-181">Choisissez **Fichier**, puis **Enregistrer tout**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-181">Choose **File** and then choose **Save All**.</span></span> <span data-ttu-id="af8b3-182">Ensuite, copiez le fichier manifeste XML du complément (par exemple, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="af8b3-182">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span>

5. <span data-ttu-id="af8b3-183">À l’aide du programme **Explorateur de fichier**, accédez au partage de fichiers réseau que vous avez créé à l’[Étape 1 : Créer un dossier partagé](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file), puis collez le fichier manifeste dans le dossier.</span><span class="sxs-lookup"><span data-stu-id="af8b3-183">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="af8b3-184">Étape 7 : insérer et exécuter le complément dans l’application cliente Office</span><span class="sxs-lookup"><span data-stu-id="af8b3-184">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="af8b3-185">Démarrez Word et créez un document.</span><span class="sxs-lookup"><span data-stu-id="af8b3-185">Start Word and create a document.</span></span>

2. <span data-ttu-id="af8b3-186">Sur le ruban, cliquez sur **Insérer** > **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="af8b3-186">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="af8b3-p115">Dans la boîte de dialogue **Compléments Office**, choisissez **DOSSIER PARTAGÉ**. Word recherche le dossier que vous avez désigné comme catalogue de compléments approuvés (à l’[étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) et affiche les compléments dans la boîte de dialogue. Vous devriez voir l’icône de votre exemple de complément.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p115">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="af8b3-p116">Cliquez sur l’icône de votre complément, puis choisissez **Ajouter**. Un bouton **Afficher le volet de tâches** pour votre complément est ajouté au ruban.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p116">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="af8b3-p117">Dans le ruban de l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches**. Le complément s’ouvre dans un volet de tâches à droite du document actif.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p117">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="af8b3-p118">Vérifiez que le complément fonctionne en sélectionnant du texte dans le document et en choisissant le bouton **Mettre en surbrillance** dans le volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="af8b3-p118">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="af8b3-196">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="af8b3-196">See also</span></span>

- [<span data-ttu-id="af8b3-197">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="af8b3-197">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="af8b3-198">Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication</span><span class="sxs-lookup"><span data-stu-id="af8b3-198">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
