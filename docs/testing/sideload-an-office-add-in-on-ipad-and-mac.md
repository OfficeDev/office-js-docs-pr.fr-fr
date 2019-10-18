---
title: Chargement de version test des compléments Office sur iPad et Mac
description: ''
ms.date: 07/29/2019
localization_priority: Priority
ms.openlocfilehash: 010812cf02bb96f26db64aa89d6e9fd3ce679ea9
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940870"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="67605-102">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="67605-102">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="67605-p101">Pour voir comment votre complément s’exécutera dans Office sur iOS, vous pouvez charger une version test du manifeste de votre complément sur un iPad à l’aide d’iTunes ou directement dans Office sur Mac. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="67605-p101">To see how your add-in will run in Office for iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office for Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span> 

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="67605-105">Configuration requise pour Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="67605-105">Prerequisites for Office for iOS</span></span>

- <span data-ttu-id="67605-106">Un ordinateur Windows ou Mac sur lequel [iTunes](https://www.apple.com/itunes/download/) est installé.</span><span class="sxs-lookup"><span data-stu-id="67605-106">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
    
- <span data-ttu-id="67605-107">Un iPad fonctionnant sous iOS 8.2 ou version ultérieure sur lequel [Excel sur iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) est installé et disposant d’un câble de synchronisation.</span><span class="sxs-lookup"><span data-stu-id="67605-107">An iPad running iOS 8.2 or later with [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>
    
- <span data-ttu-id="67605-108">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="67605-108">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="67605-109">Configuration requise pour Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="67605-109">Prerequisites for Office for Mac</span></span>

- <span data-ttu-id="67605-110">Un Mac fonctionnant sous OS X v10.10 « Yosemite » ou une version ultérieure, avec [Office sur Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installé.</span><span class="sxs-lookup"><span data-stu-id="67605-110">A Mac running OS X v10.10 "Yosemite" or later with [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>
    
- <span data-ttu-id="67605-111">Word sur Mac version 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="67605-111">Word for Mac version 15.18 (160109)</span></span>
   
- <span data-ttu-id="67605-112">Excel sur Mac version 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="67605-112">Excel for Mac version 15.19 (160206)</span></span>

- <span data-ttu-id="67605-113">PowerPoint sur Mac version 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="67605-113">PowerPoint for Mac version 15.24 (160614)</span></span>
    
- <span data-ttu-id="67605-114">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="67605-114">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="67605-115">Chargement d’une version test d’un complément dans Excel ou Word sur iPad</span><span class="sxs-lookup"><span data-stu-id="67605-115">Sideload an add-in on Excel or Word for iPad</span></span>

1. <span data-ttu-id="67605-p102">Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Lorsque vous connectez l’iPad à votre ordinateur pour la première fois, le message **Approuver cet ordinateur ?** s’affiche. Sélectionnez **Approuver** pour continuer.</span><span class="sxs-lookup"><span data-stu-id="67605-p102">Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with  **Trust This Computer?**. Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="67605-119">Dans iTunes, sélectionnez l’icône **iPad** en dessous de la barre de menu.</span><span class="sxs-lookup"><span data-stu-id="67605-119">In iTunes, choose the  **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="67605-120">Sous  **Réglages** sur le côté gauche d’iTunes, sélectionnez **Applications**.</span><span class="sxs-lookup"><span data-stu-id="67605-120">Under  **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="67605-121">Sur le côté droite d’iTunes, faites défiler vers  **Partage de fichiers**, puis sélectionnez  **Excel** ou **Word** dans la colonne **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="67605-121">On the right side of iTunes, scroll down to  **File Sharing**, and then choose  **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="67605-122">Au bas de la colonne  **Excel** ou **Documents Word**, sélectionnez  **Ajouter un fichier**, puis sélectionnez le fichier .xml de manifeste du complément dont vous voulez charger une version test.</span><span class="sxs-lookup"><span data-stu-id="67605-122">At the bottom of the  **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span> 
    
6. <span data-ttu-id="67605-p103">Ouvrez l'application Excel ou Word sur votre iPad. Si l'application Excel ou Word est déjà en cours d'exécution, choisissez le bouton  **Home**, puis fermez et redémarrez l'application.</span><span class="sxs-lookup"><span data-stu-id="67605-p103">Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the  **Home** button, and then close and restart the app.</span></span>
    
7. <span data-ttu-id="67605-125">Ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="67605-125">Open a document.</span></span>
    
8. <span data-ttu-id="67605-126">Choisissez  **Compléments** dans l’onglet **Insérer**. La version test chargée de votre complément peut être insérée sous l’en-tête  **Développeur** dans l’interface utilisateur **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="67605-126">Choose  **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>
    
    ![Insérer des compléments dans l’application Excel](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="67605-128">Chargement d’une version test de complément dans Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="67605-128">Sideload an add-in on Office for Mac</span></span>

> [!NOTE]
> <span data-ttu-id="67605-129">Pour charger une version test de complément Outlook sur Mac, voir l’article relatif au [chargement de version test des compléments Outlook](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="67605-129">To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

1. <span data-ttu-id="67605-p104">Ouvrez **Terminal** et accédez à l’un des dossiers suivants, dans lequel vous enregistrerez le fichier manifeste de votre complément. Si le dossier `wef` n’existe pas sur votre ordinateur, créez-le.</span><span class="sxs-lookup"><span data-stu-id="67605-p104">Open  **Terminal** and go to one of the following folders where you'll save your add-in's manifest file. If the `wef` folder doesn't exist on your computer, create it.</span></span>
    
    - <span data-ttu-id="67605-132">Pour Word : `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="67605-132">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>    
    - <span data-ttu-id="67605-133">Pour Excel : `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="67605-133">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="67605-134">Pour PowerPoint : `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="67605-134">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>
    
2. <span data-ttu-id="67605-p105">Ouvrez le dossier dans **Finder** à l’aide de la commande `open .` (sans oublier le point). Copier le fichier manifeste de votre complément dans ce dossier.</span><span class="sxs-lookup"><span data-stu-id="67605-p105">Open the folder in  **Finder** using the command `open .` (including the period or dot). Copy your add-in's manifest file to this folder.</span></span>
    
    ![Dossier WEF dans Office sur Mac](../images/all-my-files.png)

3. <span data-ttu-id="67605-p106">Ouvrez Word, puis ouvrez un document. Redémarrez Word si cette application est déjà en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="67605-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>
    
4. <span data-ttu-id="67605-140">Dans Word, choisissez **Insertion** > **Compléments** > **Mes compléments** (menu déroulant), puis choisissez votre complément.</span><span class="sxs-lookup"><span data-stu-id="67605-140">In Word, choose  **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>
    
    ![Mes compléments dans Office sur Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="67605-p107">Les versions test chargées de vos compléments ne s’afficheront pas dans la boîte de dialogue Mes compléments. Elles sont visibles uniquement dans le menu déroulant (petite flèche vers le bas à droite de Mes compléments dans l’onglet **Insérer**). Les versions test chargées de vos compléments sont répertoriées sous l’en-tête **Compléments de développeur** dans ce menu.</span><span class="sxs-lookup"><span data-stu-id="67605-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span> 
    
5. <span data-ttu-id="67605-145">Vérifiez que votre complément apparaît dans Word.</span><span class="sxs-lookup"><span data-stu-id="67605-145">Verify that your add-in is displayed in Word.</span></span>
    
    ![Complément Office affiché dans Office sur Mac](../images/lorem-ipsum-wikipedia.png)
    
### <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="67605-147">Effacement du cache de l’application Office sur un ordinateur Mac</span><span class="sxs-lookup"><span data-stu-id="67605-147">Clearing the Office application's cache on a Mac or iPad</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="see-also"></a><span data-ttu-id="67605-148">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="67605-148">See also</span></span>

- [<span data-ttu-id="67605-149">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="67605-149">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
