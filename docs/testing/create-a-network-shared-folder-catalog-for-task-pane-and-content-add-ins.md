---
title: Chargement de compléments Office pour des tests
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: acd16bb8a8d08aa6dd05f0f56921e285ee1e2e1f
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127017"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="e9d78-102">Chargement de compléments Office pour des tests</span><span class="sxs-lookup"><span data-stu-id="e9d78-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="e9d78-103">Vous pouvez installer un complément Office à des fins de test dans un client Office s’exécutant sur Windows à l’aide d’un catalogue de dossiers partagés pour publier le manifeste sur un partage de fichiers réseau.</span><span class="sxs-lookup"><span data-stu-id="e9d78-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="e9d78-104">Si votre projet de complément a été créé avec le [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office), il existe une autre méthode de chargement indépendant pouvant vous convenir.</span><span class="sxs-lookup"><span data-stu-id="e9d78-104">If your add-in project was created with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), there is an alternate way of sideloading the add-in that might work for you.</span></span> <span data-ttu-id="e9d78-105">Pour plus de détails, reportez-vous à [Chargement de versions test de compléments Office à l’aide de la commande sideload](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="e9d78-105">For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="e9d78-106">Cet article s’applique uniquement aux tests de compléments Word, Excel, PowerPoint ou Project sur Windows.</span><span class="sxs-lookup"><span data-stu-id="e9d78-106">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins on Windows.</span></span> <span data-ttu-id="e9d78-107">Si vous souhaitez tester sur une autre plateforme ou tester un complément Outlook, consultez une des rubriques suivantes pour charger une version de votre complément :</span><span class="sxs-lookup"><span data-stu-id="e9d78-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="e9d78-108">Chargement de versions test des compléments Office dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="e9d78-108">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="e9d78-109">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="e9d78-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="e9d78-110">Chargement de version test des compléments Outlook pour les tester</span><span class="sxs-lookup"><span data-stu-id="e9d78-110">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

<span data-ttu-id="e9d78-111">La vidéo suivante présente la procédure de chargement de version test de votre complément dans Office sur le web ou le bureau à l’aide d’un catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="e9d78-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="e9d78-112">Partager un dossier</span><span class="sxs-lookup"><span data-stu-id="e9d78-112">Share a folder</span></span>

1. <span data-ttu-id="e9d78-113">Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="e9d78-113">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="e9d78-114">Ouvrez le menu contextuel pour le dossier que vous souhaitez utiliser comme catalogue de dossiers partagés (cliquez sur le dossier avec le bouton droit) et choisissez **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="e9d78-115">Dans la boîte de dialogue **Propriétés**, ouvrez l’onglet **Partage**, puis choisissez le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le bouton Partager mis en évidence](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="e9d78-117">Dans la boîte de dialogue **Accès réseau**, ajoutez-vous ainsi que les autres utilisateurs et/ou groupes avec lesquels vous souhaitez partager votre complément.</span><span class="sxs-lookup"><span data-stu-id="e9d78-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="e9d78-118">Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier.</span><span class="sxs-lookup"><span data-stu-id="e9d78-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="e9d78-119">Une fois que vous avez choisi les utilisateurs avec lesquels vous souhaitez effectuer le partage, sélectionnez le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="e9d78-120">Lorsqu’un message de confirmation indiquant que **votre dossier est partagé** apparaît, notez le chemin d’accès complet du réseau qui s’affiche juste après le nom du dossier.</span><span class="sxs-lookup"><span data-stu-id="e9d78-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="e9d78-121">(Vous devrez entrer cette valeur comme **URL du catalogue** lorsque vous [spécifierez le dossier partagé comme un catalogue approuvé](#specify-the-shared-folder-as-a-trusted-catalog), tel que décrit dans la section suivante de cet article.) Sélectionnez le bouton **Terminé** pour fermer la boîte de dialogue **Accès réseau**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Boîte de dialogue Accès réseau avec le chemin d’accès partagé mis en évidence](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="e9d78-123">Choisissez le bouton **Fermer** pour fermer la boîte de dialogue **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-123">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="e9d78-124">Spécifier le dossier partagé en tant que catalogue approuvé</span><span class="sxs-lookup"><span data-stu-id="e9d78-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="e9d78-125">Ouvrez un nouveau document dans Excel, Word, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="e9d78-125">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>
    
2. <span data-ttu-id="e9d78-126">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="e9d78-127">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="e9d78-128">Choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-128">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="e9d78-129">Dans la zone **URL du catalogue**, entrez le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment.</span><span class="sxs-lookup"><span data-stu-id="e9d78-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="e9d78-130">Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="e9d78-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès du réseau mis en évidence](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="e9d78-132">Après avoir entré le chemin d’accès complet du réseau du dossier dans la zone **URL du catalogue**, choisissez le bouton **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="e9d78-133">Cochez la case **Afficher dans le menu** pour l’élément nouvellement ajouté, puis choisissez le bouton **OK** pour fermer la boîte de dialogue **Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Boîte de dialogue Centre de gestion de la confidentialité avec le catalogue sélectionné](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="e9d78-135">Sélectionnez le bouton **OK** pour fermer la boîte de dialogue **Options Word**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-135">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="e9d78-136">Fermez et ouvrez de nouveau l’application Office afin que vos modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="e9d78-136">Close and reopen the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="e9d78-137">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="e9d78-137">Sideload your add-in</span></span>


1. <span data-ttu-id="e9d78-138">Placez le fichier XML manifeste d’un complément en cours de test dans le catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="e9d78-138">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="e9d78-139">Notez que vous déployez l’application web sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="e9d78-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="e9d78-140">Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="e9d78-140">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="e9d78-141">Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.</span><span class="sxs-lookup"><span data-stu-id="e9d78-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="e9d78-142">Dans Project, sélectionnez **Mes compléments** sous l’onglet **Project** du ruban.</span><span class="sxs-lookup"><span data-stu-id="e9d78-142">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span> 

3. <span data-ttu-id="e9d78-143">Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="e9d78-143">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="e9d78-144">Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.</span><span class="sxs-lookup"><span data-stu-id="e9d78-144">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="e9d78-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e9d78-145">See also</span></span>

- [<span data-ttu-id="e9d78-146">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="e9d78-146">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="e9d78-147">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="e9d78-147">Publish your Office Add-in</span></span>](../publish/publish.md)
    
