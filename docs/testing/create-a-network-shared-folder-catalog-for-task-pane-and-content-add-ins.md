---
title: Chargement de compléments Office pour des tests
description: ''
ms.date: 08/15/2019
localization_priority: Priority
ms.openlocfilehash: 19cd599ea743fc577a5139d3f278dd3f993ec5b1
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477928"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="7dff5-102">Chargement de compléments Office pour des tests</span><span class="sxs-lookup"><span data-stu-id="7dff5-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="7dff5-103">Vous pouvez installer un complément Office à des fins de test dans un client Office s’exécutant sur Windows à l’aide d’un catalogue de dossiers partagés pour publier le manifeste sur un partage de fichiers réseau.</span><span class="sxs-lookup"><span data-stu-id="7dff5-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="7dff5-104">Si votre projet de complément a été créé avec une version suffisamment récente du [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office), le complément se charge automatiquement en version de test dans le client de bureau Office lors de l’exécution de `npm start`.</span><span class="sxs-lookup"><span data-stu-id="7dff5-104">If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), the add-in will automatically sideload in the Office desktop client when you run `npm start`.</span></span>

<span data-ttu-id="7dff5-105">Cet article s’applique uniquement aux tests de compléments Word, Excel, PowerPoint ou Project sur Windows.</span><span class="sxs-lookup"><span data-stu-id="7dff5-105">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins on Windows.</span></span> <span data-ttu-id="7dff5-106">Si vous souhaitez tester sur une autre plateforme ou tester un complément Outlook, consultez une des rubriques suivantes pour charger une version de votre complément :</span><span class="sxs-lookup"><span data-stu-id="7dff5-106">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="7dff5-107">Chargement de versions test des compléments Office dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="7dff5-107">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="7dff5-108">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="7dff5-108">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="7dff5-109">Chargement de version test des compléments Outlook pour les tester</span><span class="sxs-lookup"><span data-stu-id="7dff5-109">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

<span data-ttu-id="7dff5-110">La vidéo suivante présente la procédure de chargement de version test de votre complément dans Office sur le web ou le bureau à l’aide d’un catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="7dff5-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="7dff5-111">Partager un dossier</span><span class="sxs-lookup"><span data-stu-id="7dff5-111">Share a folder</span></span>

1. <span data-ttu-id="7dff5-112">Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="7dff5-112">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="7dff5-113">Ouvrez le menu contextuel pour le dossier que vous souhaitez utiliser comme catalogue de dossiers partagés (cliquez sur le dossier avec le bouton droit) et choisissez **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-113">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="7dff5-114">Dans la boîte de dialogue **Propriétés**, ouvrez l’onglet **Partage**, puis choisissez le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-114">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le bouton Partager mis en évidence](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="7dff5-116">Dans la boîte de dialogue **Accès réseau**, ajoutez-vous ainsi que les autres utilisateurs et/ou groupes avec lesquels vous souhaitez partager votre complément.</span><span class="sxs-lookup"><span data-stu-id="7dff5-116">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="7dff5-117">Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier.</span><span class="sxs-lookup"><span data-stu-id="7dff5-117">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="7dff5-118">Une fois que vous avez choisi les utilisateurs avec lesquels vous souhaitez effectuer le partage, sélectionnez le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-118">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="7dff5-119">Lorsqu’un message de confirmation indiquant que **votre dossier est partagé** apparaît, notez le chemin d’accès complet du réseau qui s’affiche juste après le nom du dossier.</span><span class="sxs-lookup"><span data-stu-id="7dff5-119">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="7dff5-120">(Vous devrez entrer cette valeur comme **URL du catalogue** lorsque vous [spécifierez le dossier partagé comme un catalogue approuvé](#specify-the-shared-folder-as-a-trusted-catalog), tel que décrit dans la section suivante de cet article.) Sélectionnez le bouton **Terminé** pour fermer la boîte de dialogue **Accès réseau**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-120">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Boîte de dialogue Accès réseau avec le chemin d’accès partagé mis en évidence](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="7dff5-122">Choisissez le bouton **Fermer** pour fermer la boîte de dialogue **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-122">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="7dff5-123">Spécifier le dossier partagé en tant que catalogue approuvé</span><span class="sxs-lookup"><span data-stu-id="7dff5-123">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="7dff5-124">Ouvrez un nouveau document dans Excel, Word, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="7dff5-124">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>
    
2. <span data-ttu-id="7dff5-125">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-125">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="7dff5-126">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-126">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="7dff5-127">Choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-127">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="7dff5-128">Dans la zone **URL du catalogue**, entrez le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment.</span><span class="sxs-lookup"><span data-stu-id="7dff5-128">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="7dff5-129">Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="7dff5-129">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès du réseau mis en évidence](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="7dff5-131">Après avoir entré le chemin d’accès complet du réseau du dossier dans la zone **URL du catalogue**, choisissez le bouton **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-131">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="7dff5-132">Cochez la case **Afficher dans le menu** pour l’élément nouvellement ajouté, puis choisissez le bouton **OK** pour fermer la boîte de dialogue **Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-132">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Boîte de dialogue Centre de gestion de la confidentialité avec le catalogue sélectionné](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="7dff5-134">Sélectionnez le bouton **OK** pour fermer la boîte de dialogue **Options Word**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-134">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="7dff5-135">Fermez et ouvrez de nouveau l’application Office afin que vos modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="7dff5-135">Close and reopen the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="7dff5-136">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="7dff5-136">Sideload your add-in</span></span>


1. <span data-ttu-id="7dff5-137">Placez le fichier XML manifeste d’un complément en cours de test dans le catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="7dff5-137">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="7dff5-138">Notez que vous déployez l’application web sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="7dff5-138">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="7dff5-139">Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="7dff5-139">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="7dff5-140">Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.</span><span class="sxs-lookup"><span data-stu-id="7dff5-140">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="7dff5-141">Dans Project, sélectionnez **Mes compléments** sous l’onglet **Project** du ruban.</span><span class="sxs-lookup"><span data-stu-id="7dff5-141">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span> 

3. <span data-ttu-id="7dff5-142">Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="7dff5-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="7dff5-143">Sélectionnez le nom du complément, puis choisissez **OK** pour insérer celui-ci.</span><span class="sxs-lookup"><span data-stu-id="7dff5-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="7dff5-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7dff5-144">See also</span></span>

- [<span data-ttu-id="7dff5-145">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="7dff5-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="7dff5-146">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="7dff5-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
