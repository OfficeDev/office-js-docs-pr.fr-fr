---
title: Publication de compléments du volet Office et de contenu dans un catalogue SharePoint
description: Pour rendre les compléments Office accessibles aux utilisateurs, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue de compléments pour leur organisation.
ms.date: 05/22/2019
localization_priority: Priority
ms.openlocfilehash: bffbf3e83a2e6d8d0c63252c27ba54826611f78b
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432242"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="40aa5-103">Publication de compléments du volet Office et de contenu dans un catalogue SharePoint</span><span class="sxs-lookup"><span data-stu-id="40aa5-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="40aa5-p101">Un catalogue de compléments est une collection de sites dédiée dans une application web SharePoint ou une location SharePoint Online qui héberge des bibliothèques de documents pour des compléments Office et SharePoint. Pour rendre les compléments Office accessibles aux utilisateurs dans leur organisation, les administrateurs peuvent charger des fichiers manifeste de compléments Office vers le catalogue de compléments pour leur organisation. Lorsqu’un administrateur enregistre un catalogue de compléments en tant que catalogue approuvé, les utilisateurs peuvent insérer le complément à partir de l’interface utilisateur d’insertion dans une application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="40aa5-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="40aa5-106">Les catalogues de compléments sur SharePoint ne prennent pas en charge les fonctionnalités de complément qui sont implémentées dans le nœud `VersionOverrides` du [manifeste de complément](../develop/add-in-manifests.md), comme les commandes de complément.</span><span class="sxs-lookup"><span data-stu-id="40aa5-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="40aa5-107">Si vous ciblez un environnement de cloud ou hybride, nous vous recommandons d’[utiliser un déploiement centralisé via le centre d’administration Office 365](../publish/centralized-deployment.md) pour publier vos compléments.</span><span class="sxs-lookup"><span data-stu-id="40aa5-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="40aa5-108">Les catalogues SharePoint ne sont pas pris en charge dans Office pour Mac.</span><span class="sxs-lookup"><span data-stu-id="40aa5-108">SharePoint catalogs are not supported for Office for Mac.</span></span> <span data-ttu-id="40aa5-109">Pour déployer des compléments Office sur les clients Mac, vous devez les envoyer à [AppSource](/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="40aa5-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="create-an-add-in-catalog"></a><span data-ttu-id="40aa5-110">Création d’un catalogue de compléments</span><span class="sxs-lookup"><span data-stu-id="40aa5-110">Create an add-in catalog</span></span>

<span data-ttu-id="40aa5-111">Suivez les étapes décrites dans l’une des sections suivantes pour créer un catalogue de compléments sur SharePoint ou Office 365.</span><span class="sxs-lookup"><span data-stu-id="40aa5-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="40aa5-112">Création d’un catalogue de compléments sur SharePoint local</span><span class="sxs-lookup"><span data-stu-id="40aa5-112">To set up an add-in catalog for on-premises SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="40aa5-113">L’interface utilisateur dans SharePoint local fait toujours référence aux compléments en tant qu’**applications**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="40aa5-114">Accédez au **site Administration centrale**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-114">Browse to the  **Central Administration Site**.</span></span>

2. <span data-ttu-id="40aa5-115">Dans le volet Office situé à gauche, cliquez sur **Applications**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-115">In the left task pane, choose  **Apps**.</span></span>

3. <span data-ttu-id="40aa5-116">Sur la page **Applications**, sous **Gestion des applications**, sélectionnez **	Gérer le catalogue d’applications**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>

4. <span data-ttu-id="40aa5-117">Sur la page **Gérer le catalogue d’applications**, vérifiez que vous avez sélectionné l’application web appropriée dans **Sélecteur d’applications web**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-117">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>

5. <span data-ttu-id="40aa5-118">Choisissez  **Afficher les paramètres du site**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-118">Choose  **View site settings**.</span></span>

6. <span data-ttu-id="40aa5-119">Sur la page  **Paramètre du site**, choisissez  **Administrateurs de collections de sites** pour spécifier les administrateurs de collection de sites, puis choisissez **OK**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>

7. <span data-ttu-id="40aa5-120">Pour accorder des autorisations de site aux utilisateurs, choisissez  **Autorisations de site**, puis choisissez  **Accorder des autorisations**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>

8. <span data-ttu-id="40aa5-121">Dans la boîte de dialogue  **Partager le site de catalogue d’applications**, spécifiez des utilisateurs de site, définissez les autorisations appropriées pour ces derniers, puis éventuellement d’autres options, puis choisissez  **Partager**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>

9. <span data-ttu-id="40aa5-122">Pour ajouter un complément au catalogue de compléments Office, choisissez **Applications pour Office**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-122">To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="40aa5-123">Pour créer catalogue d’applications Office 365</span><span class="sxs-lookup"><span data-stu-id="40aa5-123">To create an app catalog on Office 365</span></span>

<span data-ttu-id="40aa5-124">SharePoint l’appelle un catalogue d’« Applications », mais vous pouvez également enregistrer des compléments Office dans le catalogue.</span><span class="sxs-lookup"><span data-stu-id="40aa5-124">Even though SharePoint names the catalog an "app" catalog, you can register Office Add-ins in the app catalog.</span></span>

1. <span data-ttu-id="40aa5-125">Aller au Centre d’administration Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="40aa5-125">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="40aa5-126">Pour plus d’informations sur comment accéder au centre d’administration, voir [À propos du centre d’administration Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="40aa5-126">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="40aa5-127">Dans la page Centre d’administration Microsoft 365, développez la liste des **centres d’administration**, puis sélectionnez **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-127">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="40aa5-128">Vous devez utiliser le centre d’administration SharePoint classique pour créer le catalogue.</span><span class="sxs-lookup"><span data-stu-id="40aa5-128">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="40aa5-129">Si c’est la première fois que vous accédez au centre d’administration SharePoint, sélectionnez **Centre d’administration SharePoint classique** dans le volet gauche.</span><span class="sxs-lookup"><span data-stu-id="40aa5-129">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="40aa5-130">Dans le volet Office situé à gauche, choisissez **Applications**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-130">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="40aa5-131">Dans la page d’**applications**, choisissez **Catalogue d’applications**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-131">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="40aa5-132">Si un catalogue d’applications est déjà créé et apparaît dans cette page, vous pouvez ignorer le reste de ces étapes et accéder à la section suivante de cet article pour publier votre complément dans le catalogue.</span><span class="sxs-lookup"><span data-stu-id="40aa5-132">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="40aa5-133">Dans la page **Site de catalogue d’applications**, cliquez sur **OK** pour accepter l’option par défaut et créer un site de catalogue.</span><span class="sxs-lookup"><span data-stu-id="40aa5-133">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>

6. <span data-ttu-id="40aa5-134">Dans la page **Créer une collection de sites de catalogue d’applications**, indiquez le titre de votre site de catalogue.</span><span class="sxs-lookup"><span data-stu-id="40aa5-134">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="40aa5-135">Spécifiez l’**adresse du site web**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-135">Specify the web site address.</span></span>

8. <span data-ttu-id="40aa5-136">Précisez qui est l’\*\*administrateur \*\*.</span><span class="sxs-lookup"><span data-stu-id="40aa5-136">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="40aa5-137">Choisissez 0 (zéro) comme **quota de ressources du serveur**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-137">Set the **Server Resource Quota** to 0 (zero).</span></span> <span data-ttu-id="40aa5-138">(Le quota de ressources du serveur est lié à la limitation des solutions bac à sable (sandbox) dont les performances sont médiocres, mais vous n’installerez aucune solution bac à sable (sandbox) sur votre site de catalogue d’applications.)</span><span class="sxs-lookup"><span data-stu-id="40aa5-138">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="40aa5-139">Sélectionnez **OK**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-139">Choose **OK**.</span></span>

<span data-ttu-id="40aa5-140">Le catalogue d’applications est créé.</span><span class="sxs-lookup"><span data-stu-id="40aa5-140">The app catalog is now created.</span></span>

## <a name="publish-an-add-in-to-an-app-catalog"></a><span data-ttu-id="40aa5-141">Publication d’un complément dans un catalogue d’applications</span><span class="sxs-lookup"><span data-stu-id="40aa5-141">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="40aa5-142">Pour publier un complément dans un catalogue d’applications existant, procédez comme suit.</span><span class="sxs-lookup"><span data-stu-id="40aa5-142">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="40aa5-143">Aller au Centre d’administration Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="40aa5-143">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="40aa5-144">Pour plus d’informations sur comment accéder au centre d’administration, voir [À propos du centre d’administration Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="40aa5-144">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="40aa5-145">Dans la page Centre d’administration Microsoft 365, développez la liste des **centres d’administration**, puis sélectionnez **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-145">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="40aa5-146">Vous devez utiliser le centre d’administration SharePoint classique pour créer le catalogue.</span><span class="sxs-lookup"><span data-stu-id="40aa5-146">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="40aa5-147">Si c’est la première fois que vous accédez au centre d’administration SharePoint, sélectionnez **Centre d’administration SharePoint classique** dans le volet gauche.</span><span class="sxs-lookup"><span data-stu-id="40aa5-147">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="40aa5-148">Dans le volet Office situé à gauche, choisissez **Applications**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-148">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="40aa5-149">Dans la page d’**applications**, choisissez **Catalogue d’applications**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-149">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="40aa5-150">Choisissez **Distribuer des applications pour Office**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-150">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="40aa5-151">Dans la page **Applications pour Office**, cliquez sur **Nouveau**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-151">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="40aa5-152">Dans la boîte de dialogue **Ajouter un document**, sélectionnez le bouton **Choisir un fichier**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-152">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="40aa5-153">Recherchez et spécifiez le fichier [manifeste](../develop/add-in-manifests.md) à télécharger, puis sélectionnez **Ouvrir**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-153">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="40aa5-154">Dans la boîte de dialogue **Ajouter un document**, cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-154">In the **Add a document** dialog box, choose **OK**.</span></span>

    <span data-ttu-id="40aa5-p108">Les compléments de contenu et de volet Office de ce catalogue sont désormais disponibles dans la boîte de dialogue **Compléments Office**. Pour y accéder, choisissez **Mes compléments** sous l’onglet **Insérer**, puis choisissez **MON ORGANISATION**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-p108">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="40aa5-157">Expérience des utilisateurs finaux avec le catalogue des compléments</span><span class="sxs-lookup"><span data-stu-id="40aa5-157">End user experience with the add-in catalog</span></span>

<span data-ttu-id="40aa5-158">Les utilisateurs finaux peuvent accéder au catalogue des compléments dans une application Office en procédant comme suit :</span><span class="sxs-lookup"><span data-stu-id="40aa5-158">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="40aa5-159">Dans l’application Office, accédez à **Fichier**  >  **Options**  >  **Centre de gestion de la confidentialité**  >  **Paramètres du centre de gestion de la confidentialité**  >  **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-159">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>

2. <span data-ttu-id="40aa5-160">Spécifiez l’URL de la _collection de sites SharePoint parente_ du catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="40aa5-160">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 

    <span data-ttu-id="40aa5-161">Par exemple, si l’URL du catalogue de compléments Office est :</span><span class="sxs-lookup"><span data-stu-id="40aa5-161">For example, if the URL of the Office Add-ins catalog is:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    <span data-ttu-id="40aa5-162">Spécifiez simplement l’URL de la collection de sites parente :</span><span class="sxs-lookup"><span data-stu-id="40aa5-162">Specify just the URL of the parent site collection:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. <span data-ttu-id="40aa5-p109">Fermez puis rouvrez l’application Office. Le catalogue de compléments est disponible dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="40aa5-p109">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="40aa5-165">Par ailleurs, un administrateur peut spécifier un catalogue de compléments Office sur SharePoint à l’aide d’une stratégie de groupe.</span><span class="sxs-lookup"><span data-stu-id="40aa5-165">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="40aa5-166">Pour plus d’informations, reportez-vous à la section relative à l’[utilisation d’une stratégie de groupe pour gérer la manière dont les utilisateurs peuvent installer et utiliser des compléments Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span><span class="sxs-lookup"><span data-stu-id="40aa5-166">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>
