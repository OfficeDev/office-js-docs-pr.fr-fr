---
title: Chargement de version test des compléments Office sur iPad et Mac
description: Testez votre complément Office sur iPad et Mac par chargement
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 4863a55d21ab37411e76810a744f103cc364f7c1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719776"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="71e7d-103">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="71e7d-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="71e7d-p101">Pour voir comment votre complément s’exécutera dans Office sur iOS, vous pouvez charger une version test du manifeste de votre complément sur un iPad à l’aide d’iTunes ou directement dans Office sur Mac. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="71e7d-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="71e7d-106">Configuration requise pour Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="71e7d-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="71e7d-107">Un ordinateur Windows ou Mac sur lequel [iTunes](https://www.apple.com/itunes/download/) est installé.</span><span class="sxs-lookup"><span data-stu-id="71e7d-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>

- <span data-ttu-id="71e7d-108">Un iPad fonctionnant sous iOS 8.2 ou version ultérieure sur lequel [Excel sur iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) est installé et disposant d’un câble de synchronisation.</span><span class="sxs-lookup"><span data-stu-id="71e7d-108">An iPad running iOS 8.2 or later with [Excel on iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>

- <span data-ttu-id="71e7d-109">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="71e7d-109">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="71e7d-110">Configuration requise pour Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="71e7d-110">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="71e7d-111">Un Mac fonctionnant sous OS X v10.10 « Yosemite » ou une version ultérieure, avec [Office sur Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installé.</span><span class="sxs-lookup"><span data-stu-id="71e7d-111">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="71e7d-112">Word sur Mac version 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="71e7d-112">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="71e7d-113">Excel sur Mac version 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="71e7d-113">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="71e7d-114">PowerPoint sur Mac version 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="71e7d-114">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="71e7d-115">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="71e7d-115">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="71e7d-116">Chargement d’une version test d’un complément dans Excel ou Word sur iPad</span><span class="sxs-lookup"><span data-stu-id="71e7d-116">Sideload an add-in on Excel or Word on iPad</span></span>

1. <span data-ttu-id="71e7d-117">Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="71e7d-117">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="71e7d-118">Si vous connectez l’ordinateur iPad à votre ordinateur pour la première fois, vous êtes invité à **approuver cet ordinateur ?**.</span><span class="sxs-lookup"><span data-stu-id="71e7d-118">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="71e7d-119">Sélectionnez **approuver** pour continuer.</span><span class="sxs-lookup"><span data-stu-id="71e7d-119">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="71e7d-120">Dans iTunes, sélectionnez l’icône **iPad** en dessous de la barre de menu.</span><span class="sxs-lookup"><span data-stu-id="71e7d-120">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="71e7d-121">Sous **Réglages** sur le côté gauche d’iTunes, sélectionnez **Applications**.</span><span class="sxs-lookup"><span data-stu-id="71e7d-121">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="71e7d-122">Sur le côté droite d’iTunes, faites défiler vers **Partage de fichiers**, puis sélectionnez **Excel** ou **Word** dans la colonne **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="71e7d-122">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="71e7d-123">Au bas de la colonne **Excel** ou **documents Word** , sélectionnez **Ajouter un fichier**, puis sélectionnez le fichier manifest. xml du complément que vous souhaitez chargement.</span><span class="sxs-lookup"><span data-stu-id="71e7d-123">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="71e7d-124">Ouvrez l'application Excel ou Word sur votre iPad.</span><span class="sxs-lookup"><span data-stu-id="71e7d-124">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="71e7d-125">Si l’application Excel ou Word est déjà en cours d’exécution, cliquez sur le bouton **Accueil** , puis fermez et redémarrez l’application.</span><span class="sxs-lookup"><span data-stu-id="71e7d-125">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="71e7d-126">Ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="71e7d-126">Open a document.</span></span>

8. <span data-ttu-id="71e7d-127">Choisissez **compléments** sous l’onglet **insertion** . Votre complément versions test chargées peut être inséré sous le titre **développeur** dans l’interface utilisateur des **compléments** .</span><span class="sxs-lookup"><span data-stu-id="71e7d-127">Choose **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Insérer des compléments dans l’application Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="71e7d-129">Chargement d’une version test de complément dans Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="71e7d-129">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="71e7d-130">Pour charger une version test de complément Outlook sur Mac, voir l’article relatif au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="71e7d-130">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="71e7d-131">Ouvrez le **Terminal** et accédez à l’un des dossiers suivants, dans lequel vous allez enregistrer le fichier manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="71e7d-131">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="71e7d-132">Si le dossier `wef` n’existe pas sur votre ordinateur, créez-le.</span><span class="sxs-lookup"><span data-stu-id="71e7d-132">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="71e7d-133">Pour Word : `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="71e7d-133">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>    
    - <span data-ttu-id="71e7d-134">Pour Excel : `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="71e7d-134">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="71e7d-135">Pour PowerPoint : `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="71e7d-135">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="71e7d-136">Ouvrez le dossier dans **Finder** à l’aide `open .` de la commande (y compris le point ou le point).</span><span class="sxs-lookup"><span data-stu-id="71e7d-136">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="71e7d-137">Copier le fichier manifeste de votre complément dans ce dossier.</span><span class="sxs-lookup"><span data-stu-id="71e7d-137">Copy your add-in's manifest file to this folder.</span></span>

    ![Dossier WEF dans Office sur Mac](../images/all-my-files.png)

3. <span data-ttu-id="71e7d-p106">Ouvrez Word, puis ouvrez un document. Redémarrez Word si cette application est déjà en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="71e7d-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="71e7d-141">Dans Word, sélectionnez **Insérer** > des**compléments** > **My Add-ins** (menu déroulant), puis choisissez votre complément.</span><span class="sxs-lookup"><span data-stu-id="71e7d-141">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Mes compléments dans Office sur Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="71e7d-p107">Les versions test chargées de vos compléments ne s’afficheront pas dans la boîte de dialogue Mes compléments. Elles sont visibles uniquement dans le menu déroulant (petite flèche vers le bas à droite de Mes compléments dans l’onglet **Insérer**). Les versions test chargées de vos compléments sont répertoriées sous l’en-tête **Compléments de développeur** dans ce menu.</span><span class="sxs-lookup"><span data-stu-id="71e7d-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="71e7d-146">Vérifiez que votre complément apparaît dans Word.</span><span class="sxs-lookup"><span data-stu-id="71e7d-146">Verify that your add-in is displayed in Word.</span></span>

    ![Complément Office affiché dans Office sur Mac](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="71e7d-148">Supprimer un complément versions test chargées</span><span class="sxs-lookup"><span data-stu-id="71e7d-148">Remove a sideloaded add-in</span></span>

<span data-ttu-id="71e7d-149">Vous pouvez supprimer un complément précédemment versions test chargées en effaçant le cache Office sur votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="71e7d-149">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="71e7d-150">Pour plus d’informations sur la façon d’effacer le cache de chaque plateforme et hôte, consultez l’article [effacer le cache Office](clear-cache.md).</span><span class="sxs-lookup"><span data-stu-id="71e7d-150">Details on how to clear the cache for each platform and host can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="71e7d-151">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="71e7d-151">See also</span></span>

- [<span data-ttu-id="71e7d-152">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="71e7d-152">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
