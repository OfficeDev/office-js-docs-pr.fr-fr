---
title: Chargement de version test des compléments Office dans Office Online
description: Tester votre complément Office dans Office Online par chargement de version test
ms.date: 04/29/2019
localization_priority: Priority
ms.openlocfilehash: 2bcab7b41fa7f5b9590aacc19645253ee822eeb8
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517085"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="33f0e-103">Chargement de version test des compléments Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="33f0e-103">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="33f0e-104">Vous procéder à un chargement de version test pour installer un complément Office sans avoir à le placer au préalable dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="33f0e-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="33f0e-105">Le chargement de version test s’effectue dans Office 365 ou Office Online.</span><span class="sxs-lookup"><span data-stu-id="33f0e-105">Sideloading can be done in either Office 365 or Office Online.</span></span> <span data-ttu-id="33f0e-106">La procédure est légèrement différente entre les deux plateformes.</span><span class="sxs-lookup"><span data-stu-id="33f0e-106">The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="33f0e-107">Lorsque vous chargez une version test d’un complément, le manifeste du complément est stocké dans le stockage local du navigateur. Ainsi, si vous videz le cache du navigateur ou si vous basculez vers un autre navigateur, vous devez à nouveau charger une version test de complément.</span><span class="sxs-lookup"><span data-stu-id="33f0e-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="33f0e-p102">Tel que décrit dans cet article, le chargement de version test est pris en charge dans Word, Excel et PowerPoint. Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="33f0e-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="33f0e-110">La vidéo suivante présente la procédure de chargement de version test de votre complément dans la version de bureau Office ou Office Online.</span><span class="sxs-lookup"><span data-stu-id="33f0e-110">The following video walks you through the process of sideloading your add-in in Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-online"></a><span data-ttu-id="33f0e-111">Chargement de version test d’un complément Office dans Office Online</span><span class="sxs-lookup"><span data-stu-id="33f0e-111">Sideload an Office Add-in in Office Online</span></span>

1. <span data-ttu-id="33f0e-112">Ouvrez [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="33f0e-112">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="33f0e-113">Dans **Commencer à utiliser les applications en ligne maintenant**, choisissez **Excel**, **Word** ou **PowerPoint**, puis ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="33f0e-113">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="33f0e-114">Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="33f0e-115">Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MES COMPLÉMENTS**, choisissez **Gérer mes compléments**, puis **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-115">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="33f0e-117">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. <span data-ttu-id="33f0e-p103">Vérifiez que votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître dans le ruban ou dans le menu contextuel. S’il s’agit d’un complément du volet Office, le volet doit apparaître.</span><span class="sxs-lookup"><span data-stu-id="33f0e-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="33f0e-122">Pour tester votre complément Office avec Edge, deux étapes de configuration sont nécessaires :</span><span class="sxs-lookup"><span data-stu-id="33f0e-122">To test your Office Add-in with Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="33f0e-123">Depuis une invite de commandes Windows, exécutez la ligne suivante : `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="33f0e-123">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="33f0e-124">Entrez « **about:flags** » dans la barre de recherche Edge pour afficher les options des Paramètres de développeur.</span><span class="sxs-lookup"><span data-stu-id="33f0e-124">Enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="33f0e-125">Activez l’option « **Autoriser le bouclage localhost** », puis redémarrez Edge.</span><span class="sxs-lookup"><span data-stu-id="33f0e-125">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![Option Autoriser le bouclage localhost de Edge avec la case à cocher activée.](../images/allow-localhost-loopback.png)


## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="33f0e-127">Chargement de version test d’un complément Office dans Office 365</span><span class="sxs-lookup"><span data-stu-id="33f0e-127">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="33f0e-128">Connectez-vous à votre compte Office 365.</span><span class="sxs-lookup"><span data-stu-id="33f0e-128">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="33f0e-129">Ouvrez le lanceur d’applications à l’extrémité gauche de la barre d’outils et sélectionnez **Excel**,  **Word** ou **PowerPoint**, puis créez un document.</span><span class="sxs-lookup"><span data-stu-id="33f0e-129">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="33f0e-130">Les étapes 3 à 6 sont identiques à celles de la section précédente, **Chargement d’une version de test d’un complément Office dans Office Online**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-130">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office Online**.</span></span>


## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="33f0e-131">Chargement d’une version test d’un complément lors de l’utilisation de Visual Studio</span><span class="sxs-lookup"><span data-stu-id="33f0e-131">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="33f0e-132">Si vous développez votre complément à l’aide de Visual Studio, le processus de chargement d’une version de teste est similaire.</span><span class="sxs-lookup"><span data-stu-id="33f0e-132">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="33f0e-133">La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** dans votre manifeste afin d’inclure l’URL complète de déploiement du complément.</span><span class="sxs-lookup"><span data-stu-id="33f0e-133">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="33f0e-134">Si vous pouvez charger une version test des compléments à partir de Visual Studio vers Office Online, vous ne pouvez pas les déboguer à partir de Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="33f0e-134">Although you can sideload add-ins from Visual Studio to Office Online, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="33f0e-135">Pour déboguer, vous devrez utiliser les outils de débogage du navigateur.</span><span class="sxs-lookup"><span data-stu-id="33f0e-135">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="33f0e-136">Pour plus d’informations, voir [Débogage de compléments dans Office Online](debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="33f0e-136">For more information, see [Debug add-ins in Office Online](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="33f0e-137">Dans Visual Studio, affichez la fenêtre **Propriétés** en choisissant **Affichage** -> **Fenêtre Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-137">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="33f0e-138">Dans l’**Explorateur de solutions**, sélectionnez le projet web.</span><span class="sxs-lookup"><span data-stu-id="33f0e-138">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="33f0e-139">Cela a pour effet d’afficher les propriétés du projet dans la fenêtre **Propriétés**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-139">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="33f0e-140">Dans la fenêtre Propriétés, copiez l’**URL SSL**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-140">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="33f0e-141">Dans le projet de complément, ouvrez le fichier XML de manifeste.</span><span class="sxs-lookup"><span data-stu-id="33f0e-141">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="33f0e-142">Veillez à modifier le code XML source.</span><span class="sxs-lookup"><span data-stu-id="33f0e-142">Be sure you are editing the source XML.</span></span> <span data-ttu-id="33f0e-143">Pour certains types de projets, Visual Studio ouvre un affichage visuel du code XML qui ne fonctionnera pas pour l’étape suivante.</span><span class="sxs-lookup"><span data-stu-id="33f0e-143">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="33f0e-144">Cherchez toutes les instances de **~remoteAppUrl/** et remplacez-les par l’URL SSL que vous venez de copier.</span><span class="sxs-lookup"><span data-stu-id="33f0e-144">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="33f0e-145">Vous verrez plusieurs remplacements en fonction du type de projet, et les nouvelles URL ressembleront à `https://localhost:44300/Home.html`.</span><span class="sxs-lookup"><span data-stu-id="33f0e-145">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="33f0e-146">Enregistrez le fichier XML.</span><span class="sxs-lookup"><span data-stu-id="33f0e-146">Save the XML file.</span></span>
7. <span data-ttu-id="33f0e-147">Cliquez avec le bouton droit sur le projet web, puis sélectionnez **Déboguer** -> **Démarrer une nouvelle instance**.</span><span class="sxs-lookup"><span data-stu-id="33f0e-147">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="33f0e-148">Cela a pour effet d’exécuter le projet web sans lancer Office.</span><span class="sxs-lookup"><span data-stu-id="33f0e-148">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="33f0e-149">À partir d’Office Online, chargez la version test du complément en suivant les étapes décrites précédemment dans [Chargement de version test d’un complément Office dans Office Online](#sideload-an-office-add-in-in-office-online).</span><span class="sxs-lookup"><span data-stu-id="33f0e-149">From Office Online, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office Online](#sideload-an-office-add-in-in-office-online).</span></span>
