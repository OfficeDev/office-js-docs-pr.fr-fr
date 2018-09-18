---
title: Publication de compléments du volet Office et de contenu dans un catalogue SharePoint
description: Pour rendre les compléments Office accessibles aux utilisateurs au sein de leur organisation, les administrateurs peuvent télécharger des fichiers manifestes des compléments Office dans le catalogue de compléments pour leur organisation.
ms.date: 01/23/2018
ms.openlocfilehash: 2d1328b9944366d063934ff5781029beccfc82c8
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944845"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="3744e-103">Publication de compléments du volet Office et de contenu dans un catalogue SharePoint</span><span class="sxs-lookup"><span data-stu-id="3744e-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="3744e-p101">Un catalogue de compléments est une collection de sites dédiée dans une application web SharePoint ou une location SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office et SharePoint. Pour rendre les compléments Office accessibles aux utilisateurs dans leur organisation, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue de compléments pour leur organisation. Lorsqu’un administrateur enregistre un catalogue de compléments en tant que catalogue approuvé, les utilisateurs peuvent insérer le complément à partir de l’interface utilisateur d’insertion dans une application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="3744e-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="3744e-106">Les catalogues de compléments sur SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud `VersionOverrides` du [manifeste de complément](../develop/add-in-manifests.md), comme les commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="3744e-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="3744e-107">Si vous ciblez un environnement de cloud ou hybride, nous vous recommandons d’[utiliser un déploiement centralisé via le centre d’administration Office 365](../publish/centralized-deployment.md) pour publier vos compléments.</span><span class="sxs-lookup"><span data-stu-id="3744e-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="3744e-108">Les catalogues SharePoint ne sont pas pris en charge dans Office pour Mac.</span><span class="sxs-lookup"><span data-stu-id="3744e-108">SharePoint catalogs are not supported for Office 2016 for Mac.</span></span> <span data-ttu-id="3744e-109">Pour déployer des compléments Office à des clients Mac, vous devez les envoyer à [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="3744e-109">To deploy Office Add-ins to Mac clients, you must submit them to the [Office Store](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="set-up-an-add-in-catalog"></a><span data-ttu-id="3744e-110">Configuration d’un catalogue de compléments</span><span class="sxs-lookup"><span data-stu-id="3744e-110">Set up an add-in catalog</span></span>

<span data-ttu-id="3744e-111">Suivez les étapes décrites dans l’une des sections suivantes pour configurer un catalogue de compléments sur SharePoint ou Office 365.</span><span class="sxs-lookup"><span data-stu-id="3744e-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a><span data-ttu-id="3744e-112">Configuration d’un catalogue de compléments sur SharePoint</span><span class="sxs-lookup"><span data-stu-id="3744e-112">To set up an add-in catalog on SharePoint</span></span>

1. <span data-ttu-id="3744e-113">Accédez au **site Administration centrale** (**Démarrer** > **Tous les programmes** > **Produits Microsoft SharePoint 2013** > **Administration centrale SharePoint 2013**).</span><span class="sxs-lookup"><span data-stu-id="3744e-113">Browse to the  **Central Administration Site** ( **Start** > **All Programs** > **Microsoft SharePoint 2013 Products** > **SharePoint 2013 Central Administration**).</span></span>
    
2. <span data-ttu-id="3744e-114">Dans le volet Office de gauche, cliquez sur  **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="3744e-114">In the left task pane, choose  **Add-ins**.</span></span>
    
3. <span data-ttu-id="3744e-115">Sur la page  **Compléments**, sous  **Gestion des compléments**, choisissez  **Gérer le catalogue de compléments**.</span><span class="sxs-lookup"><span data-stu-id="3744e-115">On the  **Add-ins** page, under **Add-in Management**, choose  **Manage Add-in Catalog**.</span></span>
    
4. <span data-ttu-id="3744e-116">Sur la page  **Gérer le catalogue de compléments**, vérifiez que vous avez sélectionné l’application web appropriée dans  **Sélecteur d’applications web**.</span><span class="sxs-lookup"><span data-stu-id="3744e-116">On the  **Manage Add-in Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>
    
5. <span data-ttu-id="3744e-117">Choisissez  **Afficher les paramètres du site**.</span><span class="sxs-lookup"><span data-stu-id="3744e-117">Choose  **View site settings**.</span></span>
    
6. <span data-ttu-id="3744e-118">Sur la page  **Paramètre du site**, choisissez  **Administrateurs de collections de sites** pour spécifier les administrateurs de collection de sites, puis choisissez **OK**.</span><span class="sxs-lookup"><span data-stu-id="3744e-118">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>
    
7. <span data-ttu-id="3744e-119">Pour accorder des autorisations de site aux utilisateurs, choisissez  **Autorisations de site**, puis choisissez  **Accorder des autorisations**.</span><span class="sxs-lookup"><span data-stu-id="3744e-119">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>
    
8. <span data-ttu-id="3744e-120">Dans la boîte de dialogue  **Partager le site de catalogue d’applications**, spécifiez des utilisateurs de site, définissez les autorisations appropriées pour ces derniers, puis éventuellement d’autres options, puis choisissez  **Partager**.</span><span class="sxs-lookup"><span data-stu-id="3744e-120">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>
    
9. <span data-ttu-id="3744e-121">Pour ajouter un complément au catalogue de compléments Office, choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="3744e-121">To add an add-in to the Office Add-ins add-in catalog, choose **Office Add-ins**.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a><span data-ttu-id="3744e-122">Configuration d’un catalogue de compléments sur Office 365</span><span class="sxs-lookup"><span data-stu-id="3744e-122">To set up an add-in catalog on Office 365</span></span>

1. <span data-ttu-id="3744e-123">Sur la page Centre d’administration Office 365, sélectionnez **Administrateur**, puis **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="3744e-123">On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.</span></span>
    
2. <span data-ttu-id="3744e-124">Dans le volet Office situé à gauche, cliquez sur  **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="3744e-124">In the left task pane, choose  **add-ins**.</span></span>
    
3. <span data-ttu-id="3744e-125">Sur la page  **Compléments**, cliquez sur  **Catalogue de compléments**.</span><span class="sxs-lookup"><span data-stu-id="3744e-125">On the  **add-ins** page, choose **Add-in Catalog**.</span></span>
    
4. <span data-ttu-id="3744e-126">Sur la page  **Site de catalogue de compléments**, cliquez sur  **OK** pour accepter l’option par défaut et créer un site de catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="3744e-126">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>
    
5. <span data-ttu-id="3744e-127">Sur la page  **Créer une collection de sites de catalogue de compléments**, indiquez le titre de votre site de catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="3744e-127">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>
    
6. <span data-ttu-id="3744e-128">Spécifiez l’adresse du site web.</span><span class="sxs-lookup"><span data-stu-id="3744e-128">Specify the web site address.</span></span>
    
7. <span data-ttu-id="3744e-p103">Définissez l’option  **Quota de stockage** sur la plus faible valeur possible (actuellement 110). Vous n’installerez que des packages de complément sur cette collection de sites et ils sont peu volumineux.</span><span class="sxs-lookup"><span data-stu-id="3744e-p103">Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.</span></span>
    
8. <span data-ttu-id="3744e-p104">Définissez l’option  **Quota de ressources du serveur** sur 0 (zéro). (Le quota de ressources du serveur est lié à la limitation des solutions bac à sable (sandbox) dont les performances sont médiocres, mais vous n’installerez aucune solution bac à sable (sandbox) sur votre site de catalogue de compléments.)</span><span class="sxs-lookup"><span data-stu-id="3744e-p104">Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>
    
9. <span data-ttu-id="3744e-133">Sélectionnez **OK**.</span><span class="sxs-lookup"><span data-stu-id="3744e-133">Choose  **OK**.</span></span>
    
10. <span data-ttu-id="3744e-p105">Pour ajouter un complément au site de catalogue de compléments, accédez au site que vous venez de créer. Dans le volet de navigation de gauche, choisissez **Compléments Office**, puis, pour télécharger un fichier manifeste de complément Office, sélectionnez **Nouveau complément**.</span><span class="sxs-lookup"><span data-stu-id="3744e-p105">To add an add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.</span></span>

## <a name="publish-an-add-in-to-an-add-in-catalog"></a><span data-ttu-id="3744e-136">Publication d’un complément dans un catalogue de compléments</span><span class="sxs-lookup"><span data-stu-id="3744e-136">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="3744e-137">Pour publier un complément dans un catalogue de compléments, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="3744e-137">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="3744e-138">Accédez au catalogue de compléments :</span><span class="sxs-lookup"><span data-stu-id="3744e-138">Browse to the add-in catalog:</span></span>

    - <span data-ttu-id="3744e-139">Ouvrez la page principale de l’Administration centrale de SharePoint.</span><span class="sxs-lookup"><span data-stu-id="3744e-139">Open the SharePoint Central Administration main page.</span></span>
    
    - <span data-ttu-id="3744e-140">Sélectionnez **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="3744e-140">Select  **Add-ins**.</span></span>
    
    - <span data-ttu-id="3744e-141">Sélectionnez **Gérer le catalogue de compléments**.</span><span class="sxs-lookup"><span data-stu-id="3744e-141">Select  **Manage Add-in Catalog**.</span></span>
    
    - <span data-ttu-id="3744e-142">Sélectionnez le lien fourni, puis choisissez **Compléments Office** dans la barre de navigation située à gauche.</span><span class="sxs-lookup"><span data-stu-id="3744e-142">Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.</span></span>
    
2. <span data-ttu-id="3744e-143">Sélectionnez le lien **Cliquer pour ajouter un nouvel élément**.</span><span class="sxs-lookup"><span data-stu-id="3744e-143">Choose the  **Click to add new item** link.</span></span>
    
3. <span data-ttu-id="3744e-144">Choisissez **Parcourir**, puis spécifiez le [manifeste](../develop/add-in-manifests.md) à télécharger.</span><span class="sxs-lookup"><span data-stu-id="3744e-144">Choose  **Browse**, and then specify the [manifest](../develop/add-in-manifests.md) to upload.</span></span>
    
    <span data-ttu-id="3744e-p106">Les compléments de contenu et de volet Office de ce catalogue sont désormais disponibles dans la boîte de dialogue **Compléments Office**. Pour y accéder, choisissez **Mes compléments** sous l’onglet **Insérer**, puis choisissez **MON ORGANISATION**.</span><span class="sxs-lookup"><span data-stu-id="3744e-p106">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="3744e-147">Expérience des utilisateurs finaux avec le catalogue des compléments</span><span class="sxs-lookup"><span data-stu-id="3744e-147">End user experience with the add-in catalog</span></span>

<span data-ttu-id="3744e-148">Les utilisateurs finaux peuvent accéder au catalogue des compléments dans une application Office en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="3744e-148">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="3744e-149">Dans l’application Office, accédez à **Fichier**  >  **Options**  >  **Centre de gestion de la confidentialité**  >  **Paramètres du centre de gestion de la confidentialité**  >  **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="3744e-149">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
    
2. <span data-ttu-id="3744e-150">Spécifiez l’URL de la _collection de sites SharePoint parente_ du catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="3744e-150">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 
    
    <span data-ttu-id="3744e-151">Par exemple, si l’URL du catalogue de compléments Office est :</span><span class="sxs-lookup"><span data-stu-id="3744e-151">For example, if the URL of the Office Add-ins catalog is:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    <span data-ttu-id="3744e-152">Spécifiez simplement l’URL de la collection de sites parente :</span><span class="sxs-lookup"><span data-stu-id="3744e-152">Specify just the URL of the parent site collection:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. <span data-ttu-id="3744e-p107">Fermez puis rouvrez l’application Office. Le catalogue de compléments est disponible dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="3744e-p107">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="3744e-155">Par ailleurs, un administrateur peut spécifier un catalogue de compléments Office sur SharePoint à l’aide d’une stratégie de groupe.</span><span class="sxs-lookup"><span data-stu-id="3744e-155">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="3744e-156">Pour plus d’informations, reportez-vous à la section [Utilisation d’une stratégie de groupe pour gérer la manière dont les utilisateurs peuvent installer et utiliser des compléments Office](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span><span class="sxs-lookup"><span data-stu-id="3744e-156">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office) on TechNet.</span></span>
