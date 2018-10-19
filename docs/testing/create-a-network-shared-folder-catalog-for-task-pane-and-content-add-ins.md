---
title: Chargement de version test de compléments Office
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 6ee8e4e9a2413b34cb8991b09d61e16888a0e6a6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640021"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="dc158-102">Chargement de version test de compléments Office</span><span class="sxs-lookup"><span data-stu-id="dc158-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="dc158-103">Vous pouvez installer un complément Office à tester dans un client Office s’exécutant sous Windows en publiant le manifeste sur un partage de fichiers réseau (instructions ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="dc158-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="dc158-p101">Si votre projet de complément a été créé avec l’outil [**Yo Office**](https://github.com/OfficeDev/generator-office), il existe une façon alternative de charger la version test correspondante qui pourrait fonctionner pour vous. Pour plus de détails, voir [Charger une version test des compléments Office à l’aide de la commande de chargement indépendant](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="dc158-p101">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you. For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="dc158-p102">Cet article s’applique uniquement aux tests des compléments Word, Excel ou PowerPoint sur Windows. Si vous souhaitez tester sur une autre plateforme ou si vous souhaitez tester un complément Outlook, consultez l'une des rubriques suivantes pour charger la version test de votre complément :</span><span class="sxs-lookup"><span data-stu-id="dc158-p102">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows. If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="dc158-108">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="dc158-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="dc158-109">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="dc158-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="dc158-110">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="dc158-110">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="dc158-111">La vidéo suivante vous guide à travers la procédure de chargement indépendant de votre complément dans la version de bureau Office ou Office Online à l’aide du catalogue d’un dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="dc158-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="dc158-112">Partager un dossier</span><span class="sxs-lookup"><span data-stu-id="dc158-112">Share a folder</span></span>

1. <span data-ttu-id="dc158-113">Dans l’Explorateur de fichiers sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="dc158-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="dc158-114">Ouvrez le menu contextuel du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés (cliquez avec le bouton droit sur le dossier) et sélectionnez **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="dc158-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="dc158-115">Dans la boîte de dialogue **Propriétés** , cliquez sur l’onglet **Partage** , puis choisissez le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="dc158-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![boîte de dialogue Propriétés du dossier avec l’onglet Partage et le bouton Partager en surbrillance](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="dc158-117">Dans la boîte de dialogue **de l’accès réseau**, ajoutez-vous ainsi que tous les autres utilisateurs et/ou groupes avec lesquels vous souhaitez partager votre complément.</span><span class="sxs-lookup"><span data-stu-id="dc158-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="dc158-118">Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier.</span><span class="sxs-lookup"><span data-stu-id="dc158-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="dc158-119">Après avoir terminé de choisir les personnes avec qui vous partagez, cliquez sur le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="dc158-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="dc158-120">Lorsque vous voyez la confirmation que **votre dossier est partagé**, notez le chemin d’accès complet du réseau qui s’affiche immédiatement après le nom du dossier.</span><span class="sxs-lookup"><span data-stu-id="dc158-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="dc158-121">(Vous devrez saisir cette valeur comme **URL de catalogue** lorsque vous [spécifiez que ce dossier partagé est un catalogue approuvé](#specify-the-shared-folder-as-a-trusted-catalog), comme le décrit la section suivante de cet article.) Cliquez sur le bouton **Terminé** pour fermer la boîte de dialogue **Accès réseau**.</span><span class="sxs-lookup"><span data-stu-id="dc158-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Boîte de dialogue Accès réseau avec le chemin d’accès de partage en surbrillance](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="dc158-123">Cliquez sur le bouton **Fermer** pour fermer la boîte de dialogue **Propriétés** .</span><span class="sxs-lookup"><span data-stu-id="dc158-123">Choose the **Close** button to close the **Workbook Connections** dialog box.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="dc158-124">Spécifier le dossier partagé en tant que catalogue approuvé</span><span class="sxs-lookup"><span data-stu-id="dc158-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="dc158-125">Ouvrez un nouveau document dans Excel, Word ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="dc158-125">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="dc158-126">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="dc158-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="dc158-127">Choisissez l’onglet **Centre de gestion de la confidentialité**, puis choisissez le bouton **Paramètres du Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="dc158-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="dc158-128">Choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="dc158-128">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="dc158-129">Dans la zone **URL du catalogue** , entrez le chemin d’accès complet du réseau vers le dossier que vous avez auparavant [partagé](#share-a-folder).</span><span class="sxs-lookup"><span data-stu-id="dc158-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="dc158-130">Si vous n’avez pas noté le chemin réseau complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="dc158-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![boîte de dialogue Propriétés du dossier avec l'onglet Partage et le chemin d’accès réseau en surbrillance](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="dc158-132">Une fois que vous avez saisi le chemin d’accès réseau complet du dossier dans la zone **URL du catalogue**, cliquez sur le bouton **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="dc158-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="dc158-133">Sélectionnez la case à cocher **Afficher dans le Menu** de l’élément nouvellement ajouté, puis cliquez sur le bouton **OK** pour fermer la boîte de dialogue **Centre de gestion de la confidentialité** .</span><span class="sxs-lookup"><span data-stu-id="dc158-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Boîte de dialogue Centre de gestion de la confidentialité avec catalogue sélectionné](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="dc158-135">Choisissez le bouton **OK** pour fermer la boîte de dialogue **Options Word** .</span><span class="sxs-lookup"><span data-stu-id="dc158-135">Choose the  **OK** button to close the **Internet Options** dialog box.</span></span>

9. <span data-ttu-id="dc158-136">Fermez et ouvrez de nouveau l’application Office afin que vos modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="dc158-136">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="dc158-137">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="dc158-137">Sideload your add-in</span></span>


1. <span data-ttu-id="dc158-138">Placez le fichier manifeste XML d’un complément en cours de test dans le catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="dc158-138">Put the manifest file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="dc158-139">Notez que vous déployez l’application web elle-même sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="dc158-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="dc158-140">Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="dc158-140">Deploy the web application itself to a web server and specify the URL in the  **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="dc158-141">Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.</span><span class="sxs-lookup"><span data-stu-id="dc158-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="dc158-142">Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="dc158-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="dc158-143">Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.</span><span class="sxs-lookup"><span data-stu-id="dc158-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="dc158-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dc158-144">See also</span></span>

- [<span data-ttu-id="dc158-145">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="dc158-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="dc158-146">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="dc158-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
