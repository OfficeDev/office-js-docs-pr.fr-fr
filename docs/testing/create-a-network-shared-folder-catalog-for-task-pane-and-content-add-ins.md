---
title: " Chargement de version test de compléments Office"
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: b143999422866dba9b43432359c12f3607261c60
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782811"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="e0566-102"> Chargement de version test de compléments Office</span><span class="sxs-lookup"><span data-stu-id="e0566-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="e0566-103">Vous pouvez installer un complément Office à tester dans un client Office s’exécutant sous Windows en publiant le manifeste sur un partage de fichiers réseau (instructions ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="e0566-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="e0566-104">Si votre projet de complément a été créé avec l’outil [**Yo Office**](https://github.com/OfficeDev/generator-office), il existe une façon alternative de charger la version test correspondante qui pourrait fonctionner pour vous.</span><span class="sxs-lookup"><span data-stu-id="e0566-104">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="e0566-105">Pour plus de détails, voir [Charger une version test des compléments Office à l’aide de la commande de chargement indépendant](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="e0566-105">Sideload Office Add-ins using the sideload command</span></span>

<span data-ttu-id="e0566-106">Cet article s’applique uniquement aux tests des compléments Word, Excel ou PowerPoint sur Windows.</span><span class="sxs-lookup"><span data-stu-id="e0566-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="e0566-107">Si vous souhaitez tester sur une autre plateforme ou si vous souhaitez tester un complément Outlook, consultez l'une des rubriques suivantes pour charger la version test de votre complément :</span><span class="sxs-lookup"><span data-stu-id="e0566-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="e0566-108">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="e0566-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="e0566-109">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="e0566-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="e0566-110">Chargement de version test de compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="e0566-110">Sideload Outlook add-ins for testing</span></span>](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="e0566-111">La vidéo suivante présente vous guide à travers la procédure de chargement indépendant de votre complément dans la version de bureau Office ou Office Online à l’aide du catalogue d'un dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="e0566-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="e0566-112">Partager un dossier</span><span class="sxs-lookup"><span data-stu-id="e0566-112">Share a folder</span></span>

1. <span data-ttu-id="e0566-113">Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="e0566-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="e0566-114">Ouvrez le menu contextuel du dossier (clic droit), puis choisissez **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="e0566-114">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="e0566-115">Ouvrez l’onglet **Partage**.</span><span class="sxs-lookup"><span data-stu-id="e0566-115">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="e0566-p103">Dans la page **Choisir les utilisateurs...**, ajoutez votre nom et celui des utilisateurs avec lesquels vous souhaitez partager votre complément. S’ils sont tous membres d’un groupe de sécurité, vous pouvez ajouter le groupe. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier.</span><span class="sxs-lookup"><span data-stu-id="e0566-p103">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="e0566-119">Choisissez **Partager** > **Terminer** > **Fermer**.</span><span class="sxs-lookup"><span data-stu-id="e0566-119">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="e0566-120">Spécifier le dossier partagé en tant que catalogue approuvé</span><span class="sxs-lookup"><span data-stu-id="e0566-120">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="e0566-121">Ouvrez un nouveau document dans Excel, Word ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="e0566-121">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="e0566-122">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="e0566-122">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="e0566-123">Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="e0566-123">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="e0566-124">Choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="e0566-124">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="e0566-125">Dans la zone **URL du catalogue**, entrez le chemin d’accès réseau complet au catalogue de dossiers partagés, puis choisissez **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="e0566-125">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="e0566-126">Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="e0566-126">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="e0566-127">Fermez l’application Office afin que vos modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="e0566-127">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="e0566-128">Charger votre complément</span><span class="sxs-lookup"><span data-stu-id="e0566-128">Sideload your add-in</span></span>

1. <span data-ttu-id="e0566-p104">Placez le fichier manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="e0566-p104">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="e0566-132">Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.</span><span class="sxs-lookup"><span data-stu-id="e0566-132">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="e0566-133">Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="e0566-133">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="e0566-134">Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.</span><span class="sxs-lookup"><span data-stu-id="e0566-134">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="e0566-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e0566-135">See also</span></span>

- [<span data-ttu-id="e0566-136">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="e0566-136">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="e0566-137">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="e0566-137">Publish your Office Add-in</span></span>](../publish/publish.md)
    
