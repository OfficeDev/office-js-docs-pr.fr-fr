---
title: Chargement de compléments Office pour des tests
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e5769ef40868ec996194725d98913e61b76279bc
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/27/2018
ms.locfileid: "21270292"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="52bdc-102">Chargez les compléments Office en version test effectuer des tests</span><span class="sxs-lookup"><span data-stu-id="52bdc-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="52bdc-103">Vous pouvez installer un complément Office pour tester dans un client Office s'exécutant sous Windows par l'une des méthodes suivantes :</span><span class="sxs-lookup"><span data-stu-id="52bdc-103">You can install an Office Add-in for testing in an Office client running on Windows by one of the following methods:</span></span>

- <span data-ttu-id="52bdc-104">Utilisez un catalogue de dossiers partagés pour publier le manifeste sur un partage de fichiers réseau (instructions ci-dessous)</span><span class="sxs-lookup"><span data-stu-id="52bdc-104">Using a shared folder catalog to publish the manifest to a network file share (instructions below)</span></span>
- [<span data-ttu-id="52bdc-105">Exécutez la commande **« npm run sideload »** à partir de la racine du dossier de projet du complément.</span><span class="sxs-lookup"><span data-stu-id="52bdc-105">Running the "**npm run sideload**" command from the root of the add-in project folder.</span></span>](sideload-office-addin-using-sideload-command.md) 
>[!NOTE]
><span data-ttu-id="52bdc-106">La méthode « npm run sideload » ne fonctionne que pour les compléments Excel, Word et PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="52bdc-106">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

<span data-ttu-id="52bdc-107">Si vous ne testez pas un complément Word, Excel ou PowerPoint sous Windows, consultez une des rubriques suivantes pour charger la version test de votre complément :</span><span class="sxs-lookup"><span data-stu-id="52bdc-107">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="52bdc-108">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="52bdc-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="52bdc-109">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="52bdc-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

<span data-ttu-id="52bdc-110">La vidéo suivante présente la procédure de chargement indépendant de votre complément dans la version de bureau Office ou Office Online à l'aide du catalogue d'un dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="52bdc-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="52bdc-111">Partager un dossier</span><span class="sxs-lookup"><span data-stu-id="52bdc-111">Share a folder</span></span>

1. <span data-ttu-id="52bdc-112">Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="52bdc-112">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="52bdc-113">Ouvrez le menu contextuel du dossier (clic droit), puis choisissez **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-113">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="52bdc-114">Ouvrez l’onglet **Partage**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-114">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="52bdc-p101">Dans la page **Choisir les utilisateurs...**, ajoutez votre nom et celui des utilisateurs avec lesquels vous souhaitez partager votre complément. S’ils sont tous membres d’un groupe de sécurité, vous pouvez ajouter le groupe. Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier.</span><span class="sxs-lookup"><span data-stu-id="52bdc-p101">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="52bdc-118">Choisissez **Partager** > **Terminer** > **Fermer**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-118">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="52bdc-119">Spécifier le dossier partagé en tant que catalogue approuvé</span><span class="sxs-lookup"><span data-stu-id="52bdc-119">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="52bdc-120">Ouvrez un nouveau document dans Excel, Word ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="52bdc-120">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="52bdc-121">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-121">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="52bdc-122">Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-122">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="52bdc-123">Choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-123">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="52bdc-124">Dans la zone **URL du catalogue**, entrez le chemin d’accès réseau complet au catalogue de dossiers partagés, puis choisissez **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-124">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="52bdc-125">Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-125">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="52bdc-126">Fermez l’application Office afin que vos modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="52bdc-126">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="52bdc-127">Charger votre complément</span><span class="sxs-lookup"><span data-stu-id="52bdc-127">Sideload your add-in</span></span>

1. <span data-ttu-id="52bdc-p102">Placez le fichier manifeste d’un complément en cours de test dans le catalogue de dossiers partagés. Notez que vous déployez l’application web sur un serveur web. Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="52bdc-p102">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="52bdc-131">Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.</span><span class="sxs-lookup"><span data-stu-id="52bdc-131">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="52bdc-132">Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="52bdc-132">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="52bdc-133">Sélectionnez le nom du complément, puis choisissez **OK** pour insérer le complément.</span><span class="sxs-lookup"><span data-stu-id="52bdc-133">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="52bdc-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="52bdc-134">See also</span></span>

- [<span data-ttu-id="52bdc-135">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="52bdc-135">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="52bdc-136">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="52bdc-136">Publish your Office Add-in</span></span>](../publish/publish.md)
    
