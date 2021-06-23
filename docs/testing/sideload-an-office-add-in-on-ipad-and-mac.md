---
title: Chargement de version test des compléments Office sur iPad et Mac
description: Testez votre Office sur iPad Mac en chargeant une version test.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: b3d7d7fa3ee69e849c112c888b66fa9deed23d88
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076202"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="a8a31-103">Chargement de version test des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="a8a31-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="a8a31-p101">Pour voir comment votre complément s’exécutera dans Office sur iOS, vous pouvez charger une version test du manifeste de votre complément sur un iPad à l’aide d’iTunes ou directement dans Office sur Mac. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.</span><span class="sxs-lookup"><span data-stu-id="a8a31-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="a8a31-106">Configuration requise pour Office sur iOS</span><span class="sxs-lookup"><span data-stu-id="a8a31-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="a8a31-107">Un ordinateur Windows ou Mac sur lequel [iTunes](https://www.apple.com/itunes/download/) est installé.</span><span class="sxs-lookup"><span data-stu-id="a8a31-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
  > [!IMPORTANT]
  > <span data-ttu-id="a8a31-108">Si vous exécutez macOS Journal, [iTunes](https://support.apple.com/HT210200) n’est plus disponible. Vous devez donc suivre les instructions de la section Chargement de version de version ultérieure d’un module de chargement de version ultérieure d’un module sur Excel ou Word sur iPad à l’aide de [macOS Importez plus](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="a8a31-108">If you're running macOS Catalina, [iTunes is no longer available](https://support.apple.com/HT210200) so you should follow the instructions in the section [Sideload an add-in on Excel or Word on iPad using macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) later in this article.</span></span>

- <span data-ttu-id="a8a31-109">Un iPad exécutant iOS 8.2 ou [](https://apps.apple.com/app/microsoft-excel/id586683407) une Excel [ou Word,](https://apps.apple.com/app/microsoft-word/id586447913) et un câble de synchronisation.</span><span class="sxs-lookup"><span data-stu-id="a8a31-109">An iPad running iOS 8.2 or later with [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) or [Word](https://apps.apple.com/app/microsoft-word/id586447913) installed, and a sync cable.</span></span>

- <span data-ttu-id="a8a31-110">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="a8a31-110">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="a8a31-111">Configuration requise pour Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="a8a31-111">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="a8a31-112">Un Mac fonctionnant sous OS X v10.10 « Yosemite » ou une version ultérieure, avec [Office sur Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installé.</span><span class="sxs-lookup"><span data-stu-id="a8a31-112">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="a8a31-113">Word sur Mac version 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="a8a31-113">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="a8a31-114">Excel sur Mac version 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="a8a31-114">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="a8a31-115">PowerPoint sur Mac version 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="a8a31-115">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="a8a31-116">Le fichier .xml de manifeste pour le complément que vous voulez tester.</span><span class="sxs-lookup"><span data-stu-id="a8a31-116">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a><span data-ttu-id="a8a31-117">Chargement d’une version de version Excel ou Word sur iPad à l’aide d’iTunes</span><span class="sxs-lookup"><span data-stu-id="a8a31-117">Sideload an add-in on Excel or Word on iPad using iTunes</span></span>

1. <span data-ttu-id="a8a31-118">Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="a8a31-118">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="a8a31-119">Si vous connectez le iPad à votre ordinateur pour la première fois, vous êtes invité à utiliser **l’outil Trust This Computer ?**.</span><span class="sxs-lookup"><span data-stu-id="a8a31-119">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="a8a31-120">Sélectionnez **Confiance** pour continuer.</span><span class="sxs-lookup"><span data-stu-id="a8a31-120">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="a8a31-121">Dans iTunes, sélectionnez l’icône **iPad** en dessous de la barre de menu.</span><span class="sxs-lookup"><span data-stu-id="a8a31-121">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="a8a31-122">Sous **Réglages** sur le côté gauche d’iTunes, sélectionnez **Applications**.</span><span class="sxs-lookup"><span data-stu-id="a8a31-122">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="a8a31-123">Sur le côté droite d’iTunes, faites défiler vers **Partage de fichiers**, puis sélectionnez **Excel** ou **Word** dans la colonne **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="a8a31-123">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="a8a31-124">En bas de la **colonne Excel** ou **Documents Word,** sélectionnez Ajouter un **fichier,** puis sélectionnez le fichier .xml manifeste du module de chargement de version de version.</span><span class="sxs-lookup"><span data-stu-id="a8a31-124">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="a8a31-125">Ouvrez l'application Excel ou Word sur votre iPad.</span><span class="sxs-lookup"><span data-stu-id="a8a31-125">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="a8a31-126">Si l Excel ou l’application Word est  déjà en cours d’exécution, sélectionnez le bouton Accueil, puis fermez et redémarrez l’application.</span><span class="sxs-lookup"><span data-stu-id="a8a31-126">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="a8a31-127">Ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="a8a31-127">Open a document.</span></span>

8. <span data-ttu-id="a8a31-128">Sélectionnez **Les add-ins** sous l’onglet Insérer. (Sous l’onglet Insertion, vous devrez **peut-être** faire défiler horizontalement jusqu’à ce que vous voyez le bouton De plus.)   Votre version de chargement de version de  votre application peut être insérée sous l’en-tête Développeur dans l’interface **utilisateur des applications.**</span><span class="sxs-lookup"><span data-stu-id="a8a31-128">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Insérez des applications dans l Excel appl.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a><span data-ttu-id="a8a31-130">Chargement d’une version de version Excel ou Word sur iPad à l’aide de macOS</span><span class="sxs-lookup"><span data-stu-id="a8a31-130">Sideload an add-in on Excel or Word on iPad using macOS Catalina</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a8a31-131">Avec l’introduction de macOS Android, Apple a abandonné [iTunes](https://support.apple.com/HT210200) sur Mac et les fonctionnalités intégrées requises pour télécharger une version de version de chargement d’applications dans **Finder**.</span><span class="sxs-lookup"><span data-stu-id="a8a31-131">With the introduction of macOS Catalina, [Apple discontinued iTunes on Mac](https://support.apple.com/HT210200) and integrated functionality required to sideload apps into **Finder**.</span></span>

1. <span data-ttu-id="a8a31-132">Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="a8a31-132">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="a8a31-133">Si vous connectez le iPad à votre ordinateur pour la première fois, vous êtes invité à utiliser **l’outil Trust This Computer ?**.</span><span class="sxs-lookup"><span data-stu-id="a8a31-133">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="a8a31-134">Sélectionnez **Confiance** pour continuer.</span><span class="sxs-lookup"><span data-stu-id="a8a31-134">Choose **Trust** to continue.</span></span> <span data-ttu-id="a8a31-135">Vous pouvez également être invité à savoir s’il s’agit d’iPad ou si vous en restétiez un.</span><span class="sxs-lookup"><span data-stu-id="a8a31-135">You may also be asked if this is a new iPad or if you're restoring one.</span></span>

2. <span data-ttu-id="a8a31-136">Dans Rechercher, sous **Emplacements,** sélectionnez **l’icône iPad** sous la barre de menus.</span><span class="sxs-lookup"><span data-stu-id="a8a31-136">In Finder, under **Locations**, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="a8a31-137">En haut de la fenêtre Finder, cliquez sur **Fichiers,** puis recherchez **Excel** ou **Word**.</span><span class="sxs-lookup"><span data-stu-id="a8a31-137">On the top of the Finder window, click on **Files**, and then locate **Excel** or **Word**.</span></span>

4. <span data-ttu-id="a8a31-138">Dans une autre fenêtre Finder, faites glisser et déposez le fichier manifest.xml du module que vous souhaitez charger de manière latérale sur le **fichier Excel** ou **Word** dans la première fenêtre Finder.</span><span class="sxs-lookup"><span data-stu-id="a8a31-138">From a different Finder window, drag and drop the manifest.xml file of the add-in you want to side load onto the **Excel** or **Word** file in the first Finder window.</span></span>

5. <span data-ttu-id="a8a31-139">Ouvrez l'application Excel ou Word sur votre iPad.</span><span class="sxs-lookup"><span data-stu-id="a8a31-139">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="a8a31-140">Si l Excel ou l’application Word est  déjà en cours d’exécution, sélectionnez le bouton Accueil, puis fermez et redémarrez l’application.</span><span class="sxs-lookup"><span data-stu-id="a8a31-140">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

6. <span data-ttu-id="a8a31-141">Ouvrez un document.</span><span class="sxs-lookup"><span data-stu-id="a8a31-141">Open a document.</span></span>

7. <span data-ttu-id="a8a31-142">Sélectionnez **Les add-ins** sous l’onglet Insérer. (Sous l’onglet Insertion, vous devrez **peut-être** faire défiler horizontalement jusqu’à ce que vous voyez le bouton De plus.)   Votre version de chargement de version de  votre application peut être insérée sous l’en-tête Développeur dans l’interface **utilisateur des applications.**</span><span class="sxs-lookup"><span data-stu-id="a8a31-142">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Insérez des applications dans l Excel appl.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="a8a31-144">Chargement d’une version test de complément dans Office sur Mac</span><span class="sxs-lookup"><span data-stu-id="a8a31-144">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="a8a31-145">Pour charger une version test de complément Outlook sur Mac, voir l’article relatif au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="a8a31-145">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="a8a31-146">Ouvrez **Terminal** et allez dans l’un des dossiers suivants où vous allez enregistrer le fichier manifeste de votre module.</span><span class="sxs-lookup"><span data-stu-id="a8a31-146">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="a8a31-147">Si le dossier `wef` n’existe pas sur votre ordinateur, créez-le.</span><span class="sxs-lookup"><span data-stu-id="a8a31-147">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="a8a31-148">Pour Word : `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="a8a31-148">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>
    - <span data-ttu-id="a8a31-149">Pour Excel : `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="a8a31-149">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="a8a31-150">Pour PowerPoint : `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="a8a31-150">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="a8a31-151">Ouvrez le dossier dans **Finder à** l’aide de la commande `open .` (y compris le point ou le point).</span><span class="sxs-lookup"><span data-stu-id="a8a31-151">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="a8a31-152">Copier le fichier manifeste de votre complément dans ce dossier.</span><span class="sxs-lookup"><span data-stu-id="a8a31-152">Copy your add-in's manifest file to this folder.</span></span>

    ![Dossier Wef dans Office sur Mac.](../images/all-my-files.png)

3. <span data-ttu-id="a8a31-p108">Ouvrez Word, puis ouvrez un document. Redémarrez Word si cette application est déjà en cours d'exécution.</span><span class="sxs-lookup"><span data-stu-id="a8a31-p108">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="a8a31-156">Dans Word, **sélectionnez Insérer** des modules de mes  >    >  **add-ins** (menu déroulant), puis choisissez votre module.</span><span class="sxs-lookup"><span data-stu-id="a8a31-156">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Mes Office sur Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="a8a31-p109">Les versions test chargées de vos compléments ne s’afficheront pas dans la boîte de dialogue Mes compléments. Elles sont visibles uniquement dans le menu déroulant (petite flèche vers le bas à droite de Mes compléments dans l’onglet **Insérer**). Les versions test chargées de vos compléments sont répertoriées sous l’en-tête **Compléments de développeur** dans ce menu.</span><span class="sxs-lookup"><span data-stu-id="a8a31-p109">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="a8a31-161">Vérifiez que votre complément apparaît dans Word.</span><span class="sxs-lookup"><span data-stu-id="a8a31-161">Verify that your add-in is displayed in Word.</span></span>

    ![Office Le add-in s’affiche Office sur Mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="a8a31-163">Supprimer un add-in chargé de nouveau</span><span class="sxs-lookup"><span data-stu-id="a8a31-163">Remove a sideloaded add-in</span></span>

<span data-ttu-id="a8a31-164">Vous pouvez supprimer un add-in précédemment chargé de nouveau en effasant le cache Office sur votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="a8a31-164">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="a8a31-165">Pour plus d’informations sur la façon de effacer le cache pour chaque plateforme et application, voir l’article Effacer [le cache Office cache.](clear-cache.md)</span><span class="sxs-lookup"><span data-stu-id="a8a31-165">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a8a31-166">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a8a31-166">See also</span></span>

- [<span data-ttu-id="a8a31-167">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="a8a31-167">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
