---
title: Héberger un complément pour Office sur Microsoft Azure
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: f0d6a5a10d2ce0620b42566be03e2d36f8a922f2
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438808"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="51b1d-102">Héberger un complément pour Office sur Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="51b1d-102">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="51b1d-p101">Le complément Office le plus simple est constitué d’un fichier manifeste XML et d’une page HTML. Le fichier manifeste XML décrit les caractéristiques du complément, telles que son nom, les applications clientes Office dans lesquelles il peut s’exécuter et l’URL de la page HTML du complément. La page HTML est contenue dans une application web avec laquelle les utilisateurs interagissent lorsqu’ils installent et exécutent votre complément au sein d’une application cliente Office. Vous pouvez héberger l’application web d’un complément Office sur n’importe quelle plateforme d’hébergement web, y compris Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="51b1d-107">Cet article décrit comment déployer une application web de complément sur Azure et [charger une version test du complément](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) pour le tester dans une application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="51b1d-107">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="51b1d-108">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="51b1d-108">Prerequisites</span></span> 

1. <span data-ttu-id="51b1d-109">Installez [Visual Studio 2017](https://www.visualstudio.com/downloads) et choisissez d’inclure la charge de travail de **développement Azure**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-109">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="51b1d-110">Si vous avez déjà installé Visual Studio 2017, [utilisez le programme d’installation Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) pour vous assurer que la charge de travail de **développement Azure** est installée.</span><span class="sxs-lookup"><span data-stu-id="51b1d-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="51b1d-111">Installez Office 2016.</span><span class="sxs-lookup"><span data-stu-id="51b1d-111">Install Office 2016.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="51b1d-112">Si vous n’avez pas encore Office 2016, vous pouvez vous [inscrire pour obtenir une version d’évaluation gratuite d’un mois](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span><span class="sxs-lookup"><span data-stu-id="51b1d-112">If you don't already have Office 2016, you can [register for a free 1-month trial](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span></span>

3.  <span data-ttu-id="51b1d-113">Obtenez un abonnement Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-113">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="51b1d-114">Si vous n’avez pas encore d’abonnement Azure, vous pouvez [en obtenir un dans le cadre de votre abonnement MSDN](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/) ou vous [inscrire pour obtenir une version d’évaluation gratuite](https://azure.microsoft.com/en-us/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="51b1d-114">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/) or [register for a free trial](https://azure.microsoft.com/en-us/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="51b1d-115">Étape 1 : Créer un dossier partagé pour héberger le fichier manifeste XML de votre complément</span><span class="sxs-lookup"><span data-stu-id="51b1d-115">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="51b1d-116">Ouvrez l’explorateur de fichiers sur votre ordinateur de développement.</span><span class="sxs-lookup"><span data-stu-id="51b1d-116">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="51b1d-117">Cliquez avec le bouton droit de la souris sur le lecteur C:\, puis choisissez **Nouveau** > **Dossier**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-117">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="51b1d-118">Nommez le nouveau dossier AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="51b1d-118">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="51b1d-119">Cliquez avec le bouton droit de la souris sur le dossier AddinManifests, puis choisissez **Partager avec** > **Des personnes spécifiques**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-119">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="51b1d-120">Dans **Partage de fichiers**, sélectionnez la flèche déroulante vers le bas, puis choisissez **Tout le monde** > **Ajouter** > **Partager**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-120">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="51b1d-p102">Dans cette procédure, vous utilisez un partage de fichiers local en tant que catalogue approuvé où vous allez stocker le fichier manifeste XML du complément. Dans un scénario réel, vous pouvez choisir de [déployer le fichier manifeste XML dans un catalogue SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou de [publier le complément dans AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store), à la place.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="51b1d-123">Étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés</span><span class="sxs-lookup"><span data-stu-id="51b1d-123">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="51b1d-124">Démarrez Word 2016 et créez un document.</span><span class="sxs-lookup"><span data-stu-id="51b1d-124">Start Word 2016 and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="51b1d-125">Bien que cet exemple utilise Word 2016, vous pouvez utiliser n’importe quelle application Office qui prend en charge des compléments Office comme Excel, Outlook, PowerPoint ou Project 2016.</span><span class="sxs-lookup"><span data-stu-id="51b1d-125">Although this example uses Word 2016, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project 2016.</span></span>
    
2.  <span data-ttu-id="51b1d-126">Choisissez **Fichier**  >  **Options**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-126">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="51b1d-127">Dans la boîte de dialogue **Options Word**, choisissez **Centre de gestion de la confidentialité**, puis **Paramètres du Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-127">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="51b1d-p103">Dans la boîte de dialogue **Centre de gestion de la confidentialité**, choisissez **Catalogues de compléments approuvés**. Saisissez le chemin d’accès UNC (Universal Naming Convention) pour le partage de fichiers que vous avez créé précédemment en tant qu’**URL du catalogue** (par exemple, \\\NomDeVotreOrdinateur\AddinManifests), puis choisissez **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="51b1d-130">Activez la case **Afficher dans le menu**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-130">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="51b1d-131">Lorsque vous stockez un fichier manifeste XML de complément sur un partage qui est défini comme un catalogue de compléments web approuvés, le complément apparaît sous **Dossier partagé** dans la boîte de dialogue **Compléments Office** lorsque l’utilisateur accède à l’onglet **Insérer** dans le ruban et choisit **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-131">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="51b1d-132">Fermez Word 2016.</span><span class="sxs-lookup"><span data-stu-id="51b1d-132">Close Word 2016.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="51b1d-133">Étape 3 : Créer une application web dans Azure</span><span class="sxs-lookup"><span data-stu-id="51b1d-133">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="51b1d-134">Créez une application web vide dans Azure en utilisant [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) ou le [portail Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span><span class="sxs-lookup"><span data-stu-id="51b1d-134">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="51b1d-135">Utilisation de Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="51b1d-135">Using Visual Studio 2017</span></span>

<span data-ttu-id="51b1d-136">Pour créer l’application web à l’aide de Visual Studio 2017, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="51b1d-136">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="51b1d-p104">Dans Visual Studio, dans le menu **Affichage**, sélectionnez **Explorateur de serveurs**. Cliquez avec le bouton droit de la souris sur **Azure** et choisissez **Se connecter à un abonnement Microsoft Azure**. Suivez les instructions pour vous connecter à votre abonnement Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="51b1d-140">Dans Visual Studio, dans **Explorateur de serveurs**, développez **Azure**, cliquez avec le bouton droit de la souris sur **App Service**, puis choisissez **Créer un App Service**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-140">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="51b1d-141">Dans la boîte de dialogue **Créer App Service**, indiquez les informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="51b1d-141">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="51b1d-p105">Entrez un **nom d’application web** unique pour votre site. Azure vérifie que le nom du site est unique dans le domaine azurewebsites.net.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="51b1d-144">Choisissez l’**abonnement** à utiliser pour créer ce site.</span><span class="sxs-lookup"><span data-stu-id="51b1d-144">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="51b1d-p106">Choisissez le **groupe de ressources** pour votre site. Si vous créez un groupe, vous devez également le nommer.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="51b1d-p107">Choisissez le **plan de service d'applications** à utiliser pour créer ce site. Si vous créez un plan, vous devez également le nommer.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="51b1d-149">Sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-149">Choose **Create**.</span></span>

    <span data-ttu-id="51b1d-150">La nouvelle application web s’affiche dans **Explorateur de serveurs** sous **Azure** >> **App Service** >> (le groupe de ressources choisi).</span><span class="sxs-lookup"><span data-stu-id="51b1d-150">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="51b1d-p108">Cliquez avec le bouton droit de la souris sur la nouvelle application web, puis choisissez **Afficher dans le navigateur**. Votre navigateur s’ouvre et affiche une page web avec le message « Votre service d’application a été créé. ».</span><span class="sxs-lookup"><span data-stu-id="51b1d-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="51b1d-153">Dans la barre d’adresse du navigateur, modifiez l’URL de l’application web pour qu’elle utilise le protocole HTTPS et appuyez sur **Entrée** pour confirmer que le protocole HTTPS est activé.</span><span class="sxs-lookup"><span data-stu-id="51b1d-153">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="51b1d-154"> Les sites Web Azure fournissent automatiquement un point de terminaison HTTPS.</span><span class="sxs-lookup"><span data-stu-id="51b1d-154">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="51b1d-155">Utilisation du portail Azure</span><span class="sxs-lookup"><span data-stu-id="51b1d-155">Using the Azure portal</span></span>

<span data-ttu-id="51b1d-156">Pour créer l’application web à l’aide du portail Azure, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="51b1d-156">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="51b1d-157">Connectez-vous au [portail Azure](https://portal.azure.com/) à l’aide de vos informations d’identification Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-157">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="51b1d-158">Choisissez **Nouveau** > **Web + mobile** > **Application web**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-158">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="51b1d-159">Dans la boîte de dialogue **Créer une application web**, renseignez ces informations :</span><span class="sxs-lookup"><span data-stu-id="51b1d-159">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="51b1d-p109">Entrez un **nom d’application** unique pour votre site. Azure vérifie que le nom du site est unique dans le domaine apps.net azureweb.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="51b1d-162">Choisissez l’**abonnement** à utiliser pour créer ce site.</span><span class="sxs-lookup"><span data-stu-id="51b1d-162">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="51b1d-p110">Choisissez le **groupe de ressources** pour votre site. Si vous créez un groupe, vous devez également le nommer.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="51b1d-165">Choisissez le **système d’exploitation** de votre site.</span><span class="sxs-lookup"><span data-stu-id="51b1d-165">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="51b1d-p111">Choisissez le **plan de service d’applications** à utiliser pour créer ce site. Si vous créez un plan, vous devez également le nommer.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="51b1d-168">Sélectionnez **Créer**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-168">Choose **Create**.</span></span>

4. <span data-ttu-id="51b1d-169">Choisissez **Notifications** (l’icône représentant une cloche qui se trouve sur le bord supérieur du portail Azure), puis choisissez la notification **Déploiements réussis** pour ouvrir la page **Vue d’ensemble** du site dans le portail Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-169">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="51b1d-170">La notification passera de **Déploiement en cours** à **Déploiements réussis** quand le déploiement du site sera terminé.</span><span class="sxs-lookup"><span data-stu-id="51b1d-170">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="51b1d-p112">Dans la section **Essentials** de la page **Vue d’ensemble** du site dans le portail Azure, sélectionnez l’URL qui s’affiche sous **URL**. Votre navigateur s’ouvre et affiche une page web avec le message « Votre service d’application a été créé. ».</span><span class="sxs-lookup"><span data-stu-id="51b1d-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="51b1d-173">Dans la barre d’adresse du navigateur, modifiez l’URL de l’application web pour qu’elle utilise le protocole HTTPS et appuyez sur **Entrée** pour confirmer que le protocole HTTPS est activé.</span><span class="sxs-lookup"><span data-stu-id="51b1d-173">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="51b1d-174"> Les sites Web Azure fournissent automatiquement un point de terminaison HTTPS.</span><span class="sxs-lookup"><span data-stu-id="51b1d-174">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="51b1d-175">Étape 4 : Créer un complément Office dans Visual Studio</span><span class="sxs-lookup"><span data-stu-id="51b1d-175">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="51b1d-176">Démarrez Visual Studio en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="51b1d-176">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="51b1d-177">Choisissez **Fichier** > **Nouveau** > **Projet**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-177">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="51b1d-178">Sous **Modèles**, développez **Visual C#** (ou **Visual Basic**), développez **Office/SharePoint** et choisissez **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-178">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="51b1d-179">Choisissez **Complément Word web**, puis cliquez sur **OK** pour accepter les paramètres par défaut.</span><span class="sxs-lookup"><span data-stu-id="51b1d-179">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="51b1d-180">Visual Studio crée un complément Word de base que vous pourrez publier tel quel, sans apporter de modifications à son projet web.</span><span class="sxs-lookup"><span data-stu-id="51b1d-180">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="51b1d-181">Étape 5 : Publier votre application web de complément Office sur Azure</span><span class="sxs-lookup"><span data-stu-id="51b1d-181">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="51b1d-182">Avec votre projet de complément ouvert dans Visual Studio, développez le nœud de solution dans l’**explorateur de solutions** pour voir les deux projets pour la solution.</span><span class="sxs-lookup"><span data-stu-id="51b1d-182">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="51b1d-p113">Cliquez avec le bouton droit de la souris sur le projet web, puis choisissez **Publier**. Le projet web contient les fichiers d’application web du complément Office, et il s’agit donc du projet que vous publiez sur Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="51b1d-185">Sur l’onglet **Publier** :</span><span class="sxs-lookup"><span data-stu-id="51b1d-185">On the **Publish** tab:</span></span>

      - <span data-ttu-id="51b1d-186">Choisissez **Microsoft Azure Application Service**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-186">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="51b1d-187">Choisissez **Sélectionner**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-187">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="51b1d-188">Choisissez **Publier**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-188">Choose **Publish**.</span></span> 

6. <span data-ttu-id="51b1d-189">Dans la boîte de dialogue **App Service**, recherchez et sélectionnez l’application web que vous avez créée à l’[étape 3 : Créer une application web dans Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure), puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-189">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="51b1d-p114">Visual Studio publie le projet web pour votre complément Office sur votre site web Azure. Une fois le projet web publié par Visual Studio, votre navigateur s’ouvre et affiche une page web avec le texte « Votre application de service d’application a été créée. » Il s’agit de la page active par défaut pour l’application web.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="51b1d-193">Pour voir la page Web de votre complément, modifiez l'URL afin qu'elle utilise HTTPS et spécifie le chemin d'accès de la page HTML de votre complément (par exemple : https://YourDomain.azurewebsites.net/Home.html).</span><span class="sxs-lookup"><span data-stu-id="51b1d-193">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="51b1d-194">Cela confirme que l'application Web de votre complément est désormais hébergée sur Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-194">This confirms that your add-in's website is now hosted on Azure.</span></span> <span data-ttu-id="51b1d-195">Copiez l'URL racine (par exemple : https://YourDomain.azurewebsites.net); vous en aurez besoin lorsque vous modifierez le fichier manifeste du complément plus loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="51b1d-195">Copy this URL because you'll need it when you edit the add-in manifest file later in this topic.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="51b1d-196">Étape 6 : Modifier et déployer le fichier manifeste XML du complément</span><span class="sxs-lookup"><span data-stu-id="51b1d-196">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="51b1d-197">Dans Visual Studio avec l’exemple de complément Office ouvert dans l’**explorateur de solutions**, développez la solution pour que les deux projets s’affichent.</span><span class="sxs-lookup"><span data-stu-id="51b1d-197">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="51b1d-p116">Développez le projet de complément Office (par exemple WordWebAddIn), cliquez sur le dossier manifeste avec le bouton droit de la souris et sélectionnez **Ouvrir**. Le fichier manifeste XML du complément s’ouvre.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="51b1d-200">Dans le fichier manifeste XML, recherchez et remplacez toutes les instances de "~remoteAppUrl" par l'URL racine de l'application Web du complément sur Azure.</span><span class="sxs-lookup"><span data-stu-id="51b1d-200">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> <span data-ttu-id="51b1d-201">Il s'agit de l'URL que vous avez copiée plus tôt après la publication de l'application Web du complément dans Azure (par exemple : https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="51b1d-201">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="51b1d-p118">Choisissez **Fichier**, puis **Enregistrer tout**. Fermez le fichier manifeste XML du complément.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="51b1d-204">Retournez dans l’**explorateur de solutions**, cliquez avec le bouton droit de la souris sur le dossier du fichier manifeste et choisissez **Ouvrir le dossier dans l'Explorateur de fichiers**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-204">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="51b1d-205">Copiez le fichier manifeste XML du complément (par exemple, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="51b1d-205">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="51b1d-206">Accédez au partage de fichiers réseau que vous avez créé à l’[étape 1 : Créer un dossier partagé](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) et collez le fichier manifeste dans le dossier.</span><span class="sxs-lookup"><span data-stu-id="51b1d-206">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="51b1d-207">Étape 7 : insérer et exécuter le complément dans l’application cliente Office</span><span class="sxs-lookup"><span data-stu-id="51b1d-207">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="51b1d-208">Démarrez Word 2016 et créez un document.</span><span class="sxs-lookup"><span data-stu-id="51b1d-208">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="51b1d-209">Sur le ruban, cliquez sur **Insérer** > **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="51b1d-209">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="51b1d-p119">Dans la boîte de dialogue **Compléments Office**, choisissez **DOSSIER PARTAGÉ**. Word recherche le dossier que vous avez désigné comme catalogue de compléments approuvés (à l’[étape 2 : Ajouter le partage de fichiers au catalogue de compléments approuvés](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) et affiche les compléments dans la boîte de dialogue. Vous devriez voir l’icône de votre exemple de complément.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="51b1d-p120">Cliquez sur l’icône de votre complément, puis choisissez **Ajouter**. Un bouton **Afficher le volet de tâches** pour votre complément est ajouté au ruban.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="51b1d-p121">Dans le ruban de l’onglet **Accueil**, choisissez le bouton **Afficher le volet de tâches**. Le complément s’ouvre dans un volet de tâches à droite du document actif.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="51b1d-p122">Vérifiez que le complément fonctionne en sélectionnant du texte dans le document et en choisissant le bouton **Mettre en surbrillance** dans le volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="51b1d-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="51b1d-219">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="51b1d-219">See also</span></span>

- [<span data-ttu-id="51b1d-220">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="51b1d-220">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="51b1d-221">Créer le package de votre complément à l’aide de Visual Studio pour préparer la publication</span><span class="sxs-lookup"><span data-stu-id="51b1d-221">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    
