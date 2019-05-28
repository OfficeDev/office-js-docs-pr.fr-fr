---
title: Chargement de version test des compléments Office sur iPad et Mac
description: ''
ms.date: 05/21/2019
localization_priority: Priority
ms.openlocfilehash: f666ffee2c7c139ea18859d995384281a31238dd
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432221"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="91f6c-102">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="91f6c-102">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="91f6c-p101">Pour voir comment votre complément s’exécutera dans Office pour iOS, vous pouvez charger une version test du manifeste de votre complément sur un iPad à l’aide d’iTunes ou directement dans Office pour Mac. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="91f6c-p101">To see how your add-in will run in Office for iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office for Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span> 

## <a name="prerequisites-for-office-for-ios"></a><span data-ttu-id="91f6c-105">Configuration requise pour Office pour iOS</span><span class="sxs-lookup"><span data-stu-id="91f6c-105">Prerequisites for Office for iOS</span></span>

- <span data-ttu-id="91f6c-106">Un ordinateur Windows ou Mac sur lequel [iTunes](https://www.apple.com/itunes/download/) est installé.</span><span class="sxs-lookup"><span data-stu-id="91f6c-106">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
    
- <span data-ttu-id="91f6c-107">Un iPad fonctionnant sous iOS 8.2 ou version ultérieure sur lequel [Excel pour iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) est installé et disposant d’un câble de synchronisation.</span><span class="sxs-lookup"><span data-stu-id="91f6c-107">An iPad running iOS 8.2 or later with [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>
    
- <span data-ttu-id="91f6c-108">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="91f6c-108">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="prerequisites-for-office-for-mac"></a><span data-ttu-id="91f6c-109">Configuration requise pour Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="91f6c-109">Prerequisites for Office for Mac</span></span>

- <span data-ttu-id="91f6c-110">Un Mac fonctionnant sous OS X v10.10 « Yosemite » ou une version ultérieure, avec [Office pour Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installé.</span><span class="sxs-lookup"><span data-stu-id="91f6c-110">A Mac running OS X v10.10 "Yosemite" or later with [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>
    
- <span data-ttu-id="91f6c-111">Word pour Mac version 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="91f6c-111">Word for Mac version 15.18 (160109).</span></span>
   
- <span data-ttu-id="91f6c-112">Excel pour Mac version 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="91f6c-112">Excel for Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="91f6c-113">PowerPoint pour Mac version 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="91f6c-113">PowerPoint for Mac version 15.24 (160614)</span></span>
    
- <span data-ttu-id="91f6c-114">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="91f6c-114">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a><span data-ttu-id="91f6c-115">Chargement d’une version test d’un complément dans Excel ou Word pour iPad</span><span class="sxs-lookup"><span data-stu-id="91f6c-115">Sideload an add-in on Excel or Word for iPad</span></span>

1. <span data-ttu-id="91f6c-p102">Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Lorsque vous connectez l’iPad à votre ordinateur pour la première fois, le message **Approuver cet ordinateur ?** s’affiche. Sélectionnez **Approuver** pour continuer.</span><span class="sxs-lookup"><span data-stu-id="91f6c-p102">Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with  **Trust This Computer?**. Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="91f6c-119">Dans iTunes, sélectionnez l’icône **iPad** en dessous de la barre de menu.</span><span class="sxs-lookup"><span data-stu-id="91f6c-119">In iTunes, choose the  **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="91f6c-120">Sous  **Réglages** sur le côté gauche d’iTunes, sélectionnez **Applications**.</span><span class="sxs-lookup"><span data-stu-id="91f6c-120">Under  **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="91f6c-121">Sur le côté droite d’iTunes, faites défiler vers  **Partage de fichiers**, puis sélectionnez  **Excel** ou **Word** dans la colonne **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="91f6c-121">On the right side of iTunes, scroll down to  **File Sharing**, and then choose  **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="91f6c-122">Au bas de la colonne  **Excel** ou **Documents Word**, sélectionnez  **Ajouter un fichier**, puis sélectionnez le fichier .xml de manifeste du complément dont vous voulez charger une version test.</span><span class="sxs-lookup"><span data-stu-id="91f6c-122">At the bottom of the  **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span> 
    
6. <span data-ttu-id="91f6c-p103">Ouvrez l'application Excel ou Word sur votre iPad. Si l'application Excel ou Word est déjà en cours d'exécution, choisissez le bouton  **Home**, puis fermez et redémarrez l'application.</span><span class="sxs-lookup"><span data-stu-id="91f6c-p103">Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the  **Home** button, and then close and restart the app.</span></span>
    
7. <span data-ttu-id="91f6c-125">Ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="91f6c-125">Open a document.</span></span>
    
8. <span data-ttu-id="91f6c-126">Choisissez  **Compléments** dans l’onglet **Insérer**. La version test chargée de votre complément peut être insérée sous l’en-tête  **Développeur** dans l’interface utilisateur **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="91f6c-126">Choose  **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>
    
    ![Insérer des compléments dans l’application Excel](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-on-office-for-mac"></a><span data-ttu-id="91f6c-128">Chargement d’une version test de complément dans Office pour Mac</span><span class="sxs-lookup"><span data-stu-id="91f6c-128">Sideload an add-in on Office for Mac</span></span>

> [!NOTE]
> <span data-ttu-id="91f6c-129">Pour charger une version test d’un complément Outlook pour Mac, consultez [Chargement de version test des compléments Outlook](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="91f6c-129">To sideload Outlook for Mac add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

1. <span data-ttu-id="91f6c-p104">Ouvrez **Terminal** et accédez à l’un des dossiers suivants, dans lequel vous enregistrerez le fichier manifeste de votre complément. Si le dossier `wef` n’existe pas sur votre ordinateur, créez-le.</span><span class="sxs-lookup"><span data-stu-id="91f6c-p104">Open  **Terminal** and go to one of the following folders where you'll save your add-in's manifest file. If the `wef` folder doesn't exist on your computer, create it.</span></span>
    
    - <span data-ttu-id="91f6c-132">Pour Word : `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="91f6c-132">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span></span>    
    - <span data-ttu-id="91f6c-133">Pour Excel : `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="91f6c-133">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span></span>
    - <span data-ttu-id="91f6c-134">Pour PowerPoint : `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="91f6c-134">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span></span>
    
2. <span data-ttu-id="91f6c-p105">Ouvrez le dossier dans **Finder** à l’aide de la commande `open .` (sans oublier le point). Copier le fichier manifeste de votre complément dans ce dossier.</span><span class="sxs-lookup"><span data-stu-id="91f6c-p105">Open the folder in  **Finder** using the command `open .` (including the period or dot). Copy your add-in's manifest file to this folder.</span></span>
    
    ![Dossier WEF dans Office pour Mac](../images/all-my-files.png)

3. <span data-ttu-id="91f6c-p106">Ouvrez Word, puis ouvrez un document. Redémarrez Word si cette application est déjà en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="91f6c-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>
    
4. <span data-ttu-id="91f6c-140">Dans Word, choisissez **Insertion** > **Compléments** > **Mes compléments** (menu déroulant), puis choisissez votre complément.</span><span class="sxs-lookup"><span data-stu-id="91f6c-140">In Word, choose  **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>
    
    ![Mes compléments dans Office pour Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="91f6c-p107">Les versions test chargées de vos compléments ne s’afficheront pas dans la boîte de dialogue Mes compléments. Elles sont visibles uniquement dans le menu déroulant (petite flèche vers le bas à droite de Mes compléments dans l’onglet **Insérer**). Les versions test chargées de vos compléments sont répertoriées sous l’en-tête **Compléments de développeur** dans ce menu.</span><span class="sxs-lookup"><span data-stu-id="91f6c-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span> 
    
5. <span data-ttu-id="91f6c-145">Vérifiez que votre complément apparaît dans Word.</span><span class="sxs-lookup"><span data-stu-id="91f6c-145">Verify that your add-in is displayed in Word.</span></span>
    
    ![Complément Office affiché dans Office pour Mac](../images/lorem-ipsum-wikipedia.png)
    
### <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="91f6c-147">Effacement du cache de l’application Office sur un ordinateur Mac</span><span class="sxs-lookup"><span data-stu-id="91f6c-147">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="91f6c-148">Les compléments sont souvent mis en cache dans Office pour Mac, pour des raisons de performances.</span><span class="sxs-lookup"><span data-stu-id="91f6c-148">Add-ins are cached often in Office for Mac, for performance reasons.</span></span> <span data-ttu-id="91f6c-149">En règle générale, vous pouvez effacer le cache en rechargeant le complément.</span><span class="sxs-lookup"><span data-stu-id="91f6c-149">Normally, the cache is cleared by reloading the add-in.</span></span> <span data-ttu-id="91f6c-150">En présence de plusieurs compléments dans le même document, il se peut que le processus d’effacement automatique du cache lors du rechargement ne fonctionne pas systématiquement.</span><span class="sxs-lookup"><span data-stu-id="91f6c-150">If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="91f6c-151">Sur un ordinateur Mac, vous pouvez effacer le cache manuellement en supprimant le contenu du dossier `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="91f6c-151">On a Mac, you can clear the cache manually by deleting everything in the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="see-also"></a><span data-ttu-id="91f6c-152">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="91f6c-152">See also</span></span>

- [<span data-ttu-id="91f6c-153">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="91f6c-153">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
