---
title: Chargement de Office pour les tester à partir d'un partage réseau
description: Découvrez comment recharger une version test d Office pour le test à partir d'un partage réseau
ms.date: 06/02/2020
localization_priority: Normal
ms.openlocfilehash: 79ab14ca34530c870d3be4cc962a7e3df358e9bc
ms.sourcegitcommit: 6057afc1776e1667b231d2e9809d261d372151f6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/30/2021
ms.locfileid: "52100277"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a><span data-ttu-id="bf8ae-103">Chargement de Office pour les tester à partir d'un partage réseau</span><span class="sxs-lookup"><span data-stu-id="bf8ae-103">Sideload Office Add-ins for testing from a network share</span></span>

<span data-ttu-id="bf8ae-104">Vous pouvez tester un Office dans un client Office qui se trouve sur Windows en publiant le manifeste sur un partage de fichiers réseau (instructions ci-dessous).</span><span class="sxs-lookup"><span data-stu-id="bf8ae-104">You can test an Office Add-in in an Office client that is on Windows by publishing the manifest to a network file share (instructions below).</span></span> <span data-ttu-id="bf8ae-105">Cette option de déploiement est destinée à être utilisée lorsque vous avez terminé le développement et le test sur un localhost et que vous souhaitez tester le module à partir d'un serveur non local ou d'un compte cloud.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-105">This deployment option is intended to be used when you have completed development and testing on a localhost and want to test the add-in from a non-local server or cloud account.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bf8ae-106">Le déploiement par partage réseau n'est pas pris en charge pour les modules de production. Cette méthode présente les limitations suivantes :</span><span class="sxs-lookup"><span data-stu-id="bf8ae-106">Deployment by network share is not supported for production add-ins. This method has the following limitations:</span></span>
> 
> - <span data-ttu-id="bf8ae-107">Le module peut uniquement être installé sur Windows ordinateurs.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-107">The add-in can only be installed on Windows computers.</span></span>
> - <span data-ttu-id="bf8ae-108">Si une nouvelle version d'un add-in change le ruban, chaque utilisateur devra réinstaller le module.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-108">If a new version of an add-in changes the ribbon, each user will have to reinstall the add-in.</span></span>


> [!NOTE]
> <span data-ttu-id="bf8ae-109">Si votre projet de complément a été créé avec une version suffisamment récente du [générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office), le complément se charge automatiquement en version de test dans le client de bureau Office lors de l’exécution de `npm start`.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-109">If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), the add-in will automatically sideload in the Office desktop client when you run `npm start`.</span></span>

<span data-ttu-id="bf8ae-110">Cet article s'applique uniquement aux tests de Excel, de PowerPoint et de Project et uniquement sur Windows.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-110">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins and only on Windows.</span></span> <span data-ttu-id="bf8ae-111">Si vous souhaitez tester sur une autre plateforme ou tester un complément Outlook, consultez une des rubriques suivantes pour charger une version de votre complément :</span><span class="sxs-lookup"><span data-stu-id="bf8ae-111">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="bf8ae-112">Chargement de versions test des compléments Office dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="bf8ae-112">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="bf8ae-113">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="bf8ae-113">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="bf8ae-114">Chargement de version test des compléments Outlook pour les tester</span><span class="sxs-lookup"><span data-stu-id="bf8ae-114">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

<span data-ttu-id="bf8ae-115">La vidéo suivante présente la procédure de chargement de version test de votre complément dans Office sur le web ou le bureau à l’aide d’un catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-115">The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="bf8ae-116">Partager un dossier</span><span class="sxs-lookup"><span data-stu-id="bf8ae-116">Share a folder</span></span>

1. <span data-ttu-id="bf8ae-117">Sur l’ordinateur Windows sur lequel vous voulez héberger votre complément, accédez au dossier parent ou à la lettre de lecteur du dossier que vous souhaitez utiliser comme catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-117">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="bf8ae-118">Ouvrez le menu contextuel pour le dossier que vous souhaitez utiliser comme catalogue de dossiers partagés (cliquez sur le dossier avec le bouton droit) et choisissez **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-118">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="bf8ae-119">Dans la boîte de dialogue **Propriétés**, ouvrez l’onglet **Partage**, puis choisissez le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-119">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le bouton Partager mis en évidence](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="bf8ae-121">Dans la boîte de dialogue **Accès réseau**, ajoutez-vous ainsi que les autres utilisateurs et/ou groupes avec lesquels vous souhaitez partager votre complément.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-121">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="bf8ae-122">Vous aurez besoin d’au moins une autorisation d’accès en **lecture/écriture** au dossier.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-122">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="bf8ae-123">Une fois que vous avez choisi les utilisateurs avec lesquels vous souhaitez effectuer le partage, sélectionnez le bouton **Partager**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-123">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="bf8ae-124">Lorsqu’un message de confirmation indiquant que **votre dossier est partagé** apparaît, notez le chemin d’accès complet du réseau qui s’affiche juste après le nom du dossier.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-124">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="bf8ae-125">(Vous devrez entrer cette valeur comme **URL du catalogue** lorsque vous [spécifierez le dossier partagé comme un catalogue approuvé](#specify-the-shared-folder-as-a-trusted-catalog), tel que décrit dans la section suivante de cet article.) Sélectionnez le bouton **Terminé** pour fermer la boîte de dialogue **Accès réseau**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-125">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Boîte de dialogue Accès réseau avec le chemin d’accès partagé mis en évidence](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="bf8ae-127">Choisissez le bouton **Fermer** pour fermer la boîte de dialogue **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-127">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="bf8ae-128">Spécifier le dossier partagé en tant que catalogue approuvé</span><span class="sxs-lookup"><span data-stu-id="bf8ae-128">Specify the shared folder as a trusted catalog</span></span>

### <a name="configure-the-trust-manually"></a><span data-ttu-id="bf8ae-129">Configurer l’approbation manuellement</span><span class="sxs-lookup"><span data-stu-id="bf8ae-129">Configure the trust manually</span></span>

1. <span data-ttu-id="bf8ae-130">Ouvrez un nouveau document dans Excel, Word, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-130">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="bf8ae-131">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-131">Choose the **File** tab, and then choose **Options**.</span></span>

3. <span data-ttu-id="bf8ae-132">Choisissez l’onglet **Fichier**, puis choisissez **Options**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-132">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

4. <span data-ttu-id="bf8ae-133">Choisissez **Catalogues de compléments approuvés**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-133">Choose **Trusted Add-in Catalogs**.</span></span>

5. <span data-ttu-id="bf8ae-134">Dans la zone **URL du catalogue**, entrez le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-134">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="bf8ae-135">Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-135">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span>

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès du réseau mis en évidence](../images/sideload-windows-properties-dialog-2.png)

6. <span data-ttu-id="bf8ae-137">Après avoir entré le chemin d’accès complet du réseau du dossier dans la zone **URL du catalogue**, choisissez le bouton **Ajouter un catalogue**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-137">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="bf8ae-138">Cochez la case **Afficher dans le menu** pour l’élément nouvellement ajouté, puis choisissez le bouton **OK** pour fermer la boîte de dialogue **Centre de gestion de la confidentialité**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-138">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Boîte de dialogue Centre de gestion de la confidentialité avec le catalogue sélectionné](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="bf8ae-140">Sélectionnez le **bouton OK** pour fermer la boîte de dialogue **Options.**</span><span class="sxs-lookup"><span data-stu-id="bf8ae-140">Choose the **OK** button to close the **Options** dialog window.</span></span>

9. <span data-ttu-id="bf8ae-141">Fermez et ouvrez de nouveau l’application Office afin que vos modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-141">Close and reopen the Office application so your changes will take effect.</span></span>

### <a name="configure-the-trust-with-a-registry-script"></a><span data-ttu-id="bf8ae-142">Configurer l’approbation à l’aide d’un script du Registre</span><span class="sxs-lookup"><span data-stu-id="bf8ae-142">Configure the trust with a Registry script</span></span>

1. <span data-ttu-id="bf8ae-143">Dans un éditeur de texte, créez un fichier nommé TrustNetworkShareCatalog.reg.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-143">In a text editor, create a file named TrustNetworkShareCatalog.reg.</span></span>

2. <span data-ttu-id="bf8ae-144">Ajoutez le contenu suivant au fichier :</span><span class="sxs-lookup"><span data-stu-id="bf8ae-144">Add the following content to the file:</span></span>

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```
3. <span data-ttu-id="bf8ae-145">Utilisez l’un des nombreux outils de génération de GUID en ligne, tels que le [Générateur de GUID](https://guidgenerator.com/), pour générer un GUID aléatoire, et dans le fichier TrustNetworkShareCatalog.reg, remplacez la chaîne « -Random-GUID-here- » *dans les deux emplacements* par le GUID.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-145">Use one of the many online GUID generation tools, such as [GUID Generator](https://guidgenerator.com/), to generate a random GUID, and within the TrustNetworkShareCatalog.reg file, replace the string "-random-GUID-here-" *in both places* with the GUID.</span></span> <span data-ttu-id="bf8ae-146">(Les symboles `{}` englobantes doivent subsister).</span><span class="sxs-lookup"><span data-stu-id="bf8ae-146">(The enclosing `{}` symbols should remain.)</span></span>

4. <span data-ttu-id="bf8ae-147">Remplacez la valeur`Url`, par le chemin d’accès complet du réseau vers le dossier que vous avez [partagé](#share-a-folder) précédemment.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-147">Replace the `Url` value with the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="bf8ae-148">(Notez que les caractères `\` de l’URL doivent être doublés) Si vous n’avez pas noté le chemin d’accès complet du réseau lorsque vous avez partagé le dossier, vous pouvez le récupérer dans la boîte de dialogue **Propriétés** du dossier, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-148">(Note that any `\` characters in the URL must be doubled.) If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span>

    ![Boîte de dialogue Propriétés du dossier avec l’onglet Partage et le chemin d’accès du réseau mis en évidence](../images/sideload-windows-properties-dialog-2.png)

5. <span data-ttu-id="bf8ae-150">Le fichier doit désormais se présenter comme suit.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-150">The file should now look like the following.</span></span> <span data-ttu-id="bf8ae-151">Enregistrez-le.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-151">Save it.</span></span>

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

6. <span data-ttu-id="bf8ae-152">Fermez *toutes* les applications Office.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-152">Close *all* Office applications.</span></span>

7. <span data-ttu-id="bf8ae-153">Exécutez le fichier TrustNetworkShareCatalog.reg comme vous le feriez pour n’importe quel exécutable, par exemple, double-cliquez sur celui-ci.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-153">Run the TrustNetworkShareCatalog.reg just as you would any executable, such as double-clicking it.</span></span>

## <a name="sideload-your-add-in"></a><span data-ttu-id="bf8ae-154">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="bf8ae-154">Sideload your add-in</span></span>

1. <span data-ttu-id="bf8ae-155">Placez le fichier XML manifeste d’un complément en cours de test dans le catalogue de dossiers partagés.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-155">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="bf8ae-156">Notez que vous déployez l’application web sur un serveur web.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-156">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="bf8ae-157">Veillez à spécifier l’URL dans l’élément **SourceLocation** du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-157">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

    > [!NOTE]
    > <span data-ttu-id="bf8ae-158">Pour Visual Studio projets, utilisez le manifeste créé par le projet dans le `{projectfolder}\bin\Debug\OfficeAppManifests` dossier.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-158">For Visual Studio projects, use the manifest built by the project in the `{projectfolder}\bin\Debug\OfficeAppManifests` folder.</span></span>

2. <span data-ttu-id="bf8ae-159">Dans Excel, Word ou PowerPoint, sélectionnez **Mes compléments** dans l’onglet **Insérer** du ruban.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-159">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="bf8ae-160">Dans Project, sélectionnez **Mes compléments** sous l’onglet **Project** du ruban.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-160">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span>

3. <span data-ttu-id="bf8ae-161">Choisissez **DOSSIER PARTAGÉ** dans la boîte de dialogue **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-161">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="bf8ae-162">Sélectionnez le nom du complément, puis choisissez **OK** pour insérer celui-ci.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-162">Select the name of the add-in and choose **Add** to insert the add-in.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="bf8ae-163">Supprimer un add-in chargé de nouveau</span><span class="sxs-lookup"><span data-stu-id="bf8ae-163">Remove a sideloaded add-in</span></span>

<span data-ttu-id="bf8ae-164">Vous pouvez supprimer un add-in précédemment chargé de nouveau en effantant le cache Office sur votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="bf8ae-164">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="bf8ae-165">Vous pouvez trouver plus d'informations sur la façon de Windows cache dans l'article Effacer [le cache Office cache.](clear-cache.md#clear-the-office-cache-on-windows)</span><span class="sxs-lookup"><span data-stu-id="bf8ae-165">Details on how to clear the cache on Windows can be found in the article [Clear the Office cache](clear-cache.md#clear-the-office-cache-on-windows).</span></span>

## <a name="see-also"></a><span data-ttu-id="bf8ae-166">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bf8ae-166">See also</span></span>

- [<span data-ttu-id="bf8ae-167">Valider le manifeste d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="bf8ae-167">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="bf8ae-168">Vider le cache Office</span><span class="sxs-lookup"><span data-stu-id="bf8ae-168">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="bf8ae-169">Publier votre complément Office</span><span class="sxs-lookup"><span data-stu-id="bf8ae-169">Publish your Office Add-in</span></span>](../publish/publish.md)
