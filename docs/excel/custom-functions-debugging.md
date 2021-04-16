---
ms.date: 04/09/2021
description: Découvrez comment déboguer vos fonctions personnalisées Excel qui n'utilisent pas de volet de tâches.
title: Débogage de fonctions personnalisées sans interface utilisateur
localization_priority: Normal
ms.openlocfilehash: 5b27ca44dbb891c2e1f4ae86175595dc902b74ba
ms.sourcegitcommit: 094caf086c2696e78fbdfdc6030cb0c89d32b585
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/16/2021
ms.locfileid: "51862336"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="f91e4-103">Débogage de fonctions personnalisées sans interface utilisateur</span><span class="sxs-lookup"><span data-stu-id="f91e4-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="f91e4-104">Cet article traite du  débogage uniquement pour les fonctions personnalisées qui n'utilisent pas de volet de tâches ou d'autres éléments d'interface utilisateur (fonctions personnalisées sans interface utilisateur).</span><span class="sxs-lookup"><span data-stu-id="f91e4-104">This article discusses debugging *only* for custom functions that don't use a task pane or other user interface elements (UI-less custom functions).</span></span> 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="f91e4-105">Sur Windows :</span><span class="sxs-lookup"><span data-stu-id="f91e4-105">On Windows:</span></span>
- [<span data-ttu-id="f91e4-106">Débogger Excel Desktop and Visual Studio Code (VS Code)</span><span class="sxs-lookup"><span data-stu-id="f91e4-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="f91e4-107">Débogger Excel sur le web et VS Code</span><span class="sxs-lookup"><span data-stu-id="f91e4-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="f91e4-108">Outils Excel sur le web et navigateur</span><span class="sxs-lookup"><span data-stu-id="f91e4-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="f91e4-109">Ligne de commande</span><span class="sxs-lookup"><span data-stu-id="f91e4-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="f91e4-110">Sur Mac :</span><span class="sxs-lookup"><span data-stu-id="f91e4-110">On Mac:</span></span>
- [<span data-ttu-id="f91e4-111">Outils Excel sur le web et navigateur</span><span class="sxs-lookup"><span data-stu-id="f91e4-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="f91e4-112">Ligne de commande</span><span class="sxs-lookup"><span data-stu-id="f91e4-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="f91e4-113">Par souci de simplicité, cet article présente le débogage dans le contexte de l'utilisation de Visual Studio Code pour modifier, exécuter des tâches et, dans certains cas, utiliser l'affichage débogage.</span><span class="sxs-lookup"><span data-stu-id="f91e4-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="f91e4-114">Si vous utilisez un autre éditeur ou outil de ligne de commande, consultez les [instructions](#commands-for-building-and-running-your-add-in) de ligne de commande à la fin de cet article.</span><span class="sxs-lookup"><span data-stu-id="f91e4-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="f91e4-115">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f91e4-115">Requirements</span></span>

<span data-ttu-id="f91e4-116">Ce processus de  débogage fonctionne uniquement pour les fonctions personnalisées sans interface utilisateur, qui n'utilisent pas de volet de tâches ou d'autres éléments d'interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f91e4-116">This debugging process works **only** for UI-less custom functions, which don't use a task pane or other UI elements.</span></span> <span data-ttu-id="f91e4-117">Une fonction personnalisée sans interface utilisateur peut être créée en suivant les [étapes](../tutorials/excel-tutorial-create-custom-functions.md) du didacticiel Créer des fonctions personnalisées dans Excel, puis en supprimant tous les éléments du volet Office et de l'interface utilisateur installés par le générateur Yeoman pour les [add-ins Office.](https://www.npmjs.com/package/generator-office)</span><span class="sxs-lookup"><span data-stu-id="f91e4-117">A UI-less custom function can be created by following the steps in the [Create custom functions in Excel](../tutorials/excel-tutorial-create-custom-functions.md) tutorial, and then removing all of the task pane and UI elements that are installed by the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span>

<span data-ttu-id="f91e4-118">Notez que ce processus de débogage n'est pas compatible avec les projets de fonctions personnalisées à l'aide [d'un runtime partagé.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="f91e4-118">Note that this debugging process is not compatible with custom functions projects using a [shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="f91e4-119">Utiliser le débogger VS Code pour Excel Desktop</span><span class="sxs-lookup"><span data-stu-id="f91e4-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="f91e4-120">Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur dans Office Excel sur le Bureau.</span><span class="sxs-lookup"><span data-stu-id="f91e4-120">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="f91e4-121">Le débogage du bureau pour Mac n'est pas disponible, mais peut être réalisé à l'aide des outils de navigateur et de la ligne de commande pour [déboguer Excel sur le web).](#use-the-command-line-tools-to-debug)</span><span class="sxs-lookup"><span data-stu-id="f91e4-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="f91e4-122">Exécuter votre add-in à partir de VS Code</span><span class="sxs-lookup"><span data-stu-id="f91e4-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="f91e4-123">Ouvrez le dossier de projet racine de vos fonctions personnalisées dans [VS Code.](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="f91e4-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="f91e4-124">Choose **Terminal > Run Task** and type or select **Watch**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="f91e4-125">Cela surveillera et reconstruira les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="f91e4-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="f91e4-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="f91e4-127">Démarrer le débogger VS Code</span><span class="sxs-lookup"><span data-stu-id="f91e4-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="f91e4-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span><span class="sxs-lookup"><span data-stu-id="f91e4-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="f91e4-129">Dans le menu déroulant Exécuter, choisissez **Excel Desktop (Edge Chromium).**</span><span class="sxs-lookup"><span data-stu-id="f91e4-129">From the Run drop-down menu, choose **Excel Desktop (Edge Chromium)**.</span></span>
6. <span data-ttu-id="f91e4-130">Sélectionnez **F5** (ou **exécutez -> démarrer** le débogage à partir du menu) pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="f91e4-130">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="f91e4-131">Un nouveau workbook Excel s'ouvre avec votre add-in déjà chargé et prêt à l'emploi.</span><span class="sxs-lookup"><span data-stu-id="f91e4-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="f91e4-132">Démarrer le débogage</span><span class="sxs-lookup"><span data-stu-id="f91e4-132">Start debugging</span></span>

1. <span data-ttu-id="f91e4-133">Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f91e4-133">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="f91e4-134">[Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f91e4-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="f91e4-135">Dans le workbook Excel, entrez une formule qui utilise votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f91e4-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="f91e4-136">À ce stade, l'exécution s'arrête sur la ligne de code où vous définissez le point d'arrêt.</span><span class="sxs-lookup"><span data-stu-id="f91e4-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="f91e4-137">Vous pouvez désormais vous servir de votre code, définir des montres et utiliser les fonctionnalités de débogage VS Code dont vous avez besoin.</span><span class="sxs-lookup"><span data-stu-id="f91e4-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="f91e4-138">Utiliser le débogger VS Code pour Excel dans Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="f91e4-138">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="f91e4-139">Vous pouvez utiliser VS Code pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel dans le navigateur Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="f91e4-139">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="f91e4-140">Pour utiliser VS Code avec Microsoft Edge, vous devez installer le [débogger pour l'extension Microsoft Edge.](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)</span><span class="sxs-lookup"><span data-stu-id="f91e4-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="f91e4-141">Exécuter votre add-in à partir de VS Code</span><span class="sxs-lookup"><span data-stu-id="f91e4-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="f91e4-142">Ouvrez le dossier de projet racine de vos fonctions personnalisées dans [VS Code.](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="f91e4-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="f91e4-143">Choose **Terminal > Run Task** and type or select **Watch**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="f91e4-144">Cela surveillera et reconstruira les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="f91e4-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="f91e4-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="f91e4-146">Démarrer le débogger VS Code</span><span class="sxs-lookup"><span data-stu-id="f91e4-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="f91e4-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span><span class="sxs-lookup"><span data-stu-id="f91e4-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="f91e4-148">Dans les options Debug, choisissez **Office Online (Edge Chromium).**</span><span class="sxs-lookup"><span data-stu-id="f91e4-148">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="f91e4-149">Ouvrez Excel dans le navigateur Microsoft Edge et créez un nouveau workbook.</span><span class="sxs-lookup"><span data-stu-id="f91e4-149">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="f91e4-150">Choisissez **Partager** dans le ruban et copiez le lien pour l'URL de ce nouveau workbook.</span><span class="sxs-lookup"><span data-stu-id="f91e4-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="f91e4-151">Sélectionnez **F5** (ou **exécutez > démarrer le débogage** à partir du menu) pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="f91e4-151">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="f91e4-152">Une invite s'affiche, qui demande l'URL de votre document.</span><span class="sxs-lookup"><span data-stu-id="f91e4-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="f91e4-153">Collez l'URL de votre workbook et appuyez sur Entrée.</span><span class="sxs-lookup"><span data-stu-id="f91e4-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="f91e4-154">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="f91e4-154">Sideload your add-in</span></span>

1. <span data-ttu-id="f91e4-155">Sélectionnez **l'onglet** Insérer sur le ruban et, dans la section Des **add-ins,** choisissez **Les add-ins Office.**</span><span class="sxs-lookup"><span data-stu-id="f91e4-155">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="f91e4-156">Dans la boîte de dialogue **Des add-ins Office,** sélectionnez l'onglet MES **ADD-INS,** choisissez **Manage My Add-ins**, puis **Upload My Add-in**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-156">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="f91e4-158">**Accédez** au fichier manifeste du add-in, puis sélectionnez **Télécharger.**</span><span class="sxs-lookup"><span data-stu-id="f91e4-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="f91e4-160">Définir des points d'arrêt</span><span class="sxs-lookup"><span data-stu-id="f91e4-160">Set breakpoints</span></span>
1. <span data-ttu-id="f91e4-161">Dans VS Code, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f91e4-161">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="f91e4-162">[Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f91e4-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="f91e4-163">Dans le workbook Excel, entrez une formule qui utilise votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f91e4-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="f91e4-164">Utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées dans Excel sur le web</span><span class="sxs-lookup"><span data-stu-id="f91e4-164">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="f91e4-165">Vous pouvez utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées sans interface utilisateur dans Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="f91e4-165">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="f91e4-166">Les étapes suivantes fonctionnent pour Windows et macOS.</span><span class="sxs-lookup"><span data-stu-id="f91e4-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="f91e4-167">Exécuter votre add-in à partir de Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f91e4-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="f91e4-168">Ouvrez le dossier de projet racine de vos fonctions personnalisées [dans Visual Studio Code (VS Code).](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="f91e4-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="f91e4-169">Choose **Terminal > Run Task** and type or select **Watch**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="f91e4-170">Cela surveillera et reconstruira les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="f91e4-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="f91e4-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="f91e4-172">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="f91e4-172">Sideload your add-in</span></span>

1. <span data-ttu-id="f91e4-173">Ouvrez [Office sur le web.](https://office.live.com/)</span><span class="sxs-lookup"><span data-stu-id="f91e4-173">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="f91e4-174">Ouvrez un nouveau workbook Excel.</span><span class="sxs-lookup"><span data-stu-id="f91e4-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="f91e4-175">Ouvrez **l'onglet** Insérer sur le ruban et, dans la section Des **add-ins,** choisissez **Les add-ins Office.**</span><span class="sxs-lookup"><span data-stu-id="f91e4-175">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="f91e4-176">Dans la boîte de dialogue **Des add-ins Office,** sélectionnez l'onglet MES **ADD-INS,** choisissez **Manage My Add-ins**, puis **Upload My Add-in**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-176">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="f91e4-178">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="f91e4-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="f91e4-180">Une fois que vous avez chargé une version de version sideload dans le document, celui-ci reste chargé de nouveau à chaque ouverture du document.</span><span class="sxs-lookup"><span data-stu-id="f91e4-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="f91e4-181">Démarrer le débogage</span><span class="sxs-lookup"><span data-stu-id="f91e4-181">Start debugging</span></span>

1. <span data-ttu-id="f91e4-182">Ouvrez les outils de développement dans le navigateur.</span><span class="sxs-lookup"><span data-stu-id="f91e4-182">Open developer tools in the browser.</span></span> <span data-ttu-id="f91e4-183">Pour Chrome et la plupart des navigateurs F12 ouvrent les outils de développement.</span><span class="sxs-lookup"><span data-stu-id="f91e4-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="f91e4-184">Dans les outils de développement, ouvrez votre fichier de script de code source à l'aide de **Cmd+P** ou **Ctrl+P** (**functions.js** ou **functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f91e4-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="f91e4-185">[Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="f91e4-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="f91e4-186">Si vous avez besoin de modifier le code, vous pouvez apporter des modifications dans VS Code et enregistrer les modifications.</span><span class="sxs-lookup"><span data-stu-id="f91e4-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="f91e4-187">Actualisez le navigateur pour voir les modifications chargées.</span><span class="sxs-lookup"><span data-stu-id="f91e4-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="f91e4-188">Utiliser les outils de ligne de commande pour déboguer</span><span class="sxs-lookup"><span data-stu-id="f91e4-188">Use the command line tools to debug</span></span>

<span data-ttu-id="f91e4-189">Si vous n'utilisez pas VS Code, vous pouvez utiliser la ligne de commande (par exemple, Bash ou PowerShell) pour exécuter votre add-in.</span><span class="sxs-lookup"><span data-stu-id="f91e4-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="f91e4-190">Vous devez utiliser les outils de développement du navigateur pour déboguer votre code dans Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="f91e4-190">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="f91e4-191">Vous ne pouvez pas déboguer la version de bureau d'Excel à l'aide de la ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="f91e4-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="f91e4-192">À partir de la ligne de commande, `npm run watch` exécutez la commande pour observer et reconstruire lorsque des modifications de code se produisent.</span><span class="sxs-lookup"><span data-stu-id="f91e4-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="f91e4-193">Ouvrez une deuxième fenêtre de ligne de commande (la première sera bloquée lors de l'exécution de l'observation).)</span><span class="sxs-lookup"><span data-stu-id="f91e4-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="f91e4-194">Si vous souhaitez démarrer votre application dans la version de bureau d'Excel, exécutez la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="f91e4-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="f91e4-195">Ou si vous préférez démarrer votre application dans Excel sur le web, exécutez la commande suivante:</span><span class="sxs-lookup"><span data-stu-id="f91e4-195">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="f91e4-196">Pour Excel sur le web, vous devez également charger une version de version de votre application.</span><span class="sxs-lookup"><span data-stu-id="f91e4-196">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="f91e4-197">Suivez les étapes du chargement de version de version sideload de votre [add-in](#sideload-your-add-in) pour le chargement de version de votre module.</span><span class="sxs-lookup"><span data-stu-id="f91e4-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="f91e4-198">Ensuite, continuez jusqu'à la section suivante pour démarrer le débogage.</span><span class="sxs-lookup"><span data-stu-id="f91e4-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="f91e4-199">Ouvrez les outils de développement dans le navigateur.</span><span class="sxs-lookup"><span data-stu-id="f91e4-199">Open developer tools in the browser.</span></span> <span data-ttu-id="f91e4-200">Pour Chrome et la plupart des navigateurs F12 ouvrent les outils de développement.</span><span class="sxs-lookup"><span data-stu-id="f91e4-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="f91e4-201">Dans les outils de développement, ouvrez votre fichier de script de code source (**functions.js** ou **functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f91e4-201">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="f91e4-202">Votre code de fonctions personnalisées peut se trouver à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="f91e4-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="f91e4-203">Dans le code source de la fonction personnalisée, appliquez un point d'arrêt en sélectionnant une ligne de code.</span><span class="sxs-lookup"><span data-stu-id="f91e4-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="f91e4-204">Si vous devez modifier le code, vous pouvez effectuer des modifications dans Visual Studio et enregistrer les modifications.</span><span class="sxs-lookup"><span data-stu-id="f91e4-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="f91e4-205">Actualisez le navigateur pour voir les modifications chargées.</span><span class="sxs-lookup"><span data-stu-id="f91e4-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="f91e4-206">Commandes de création et d'exécution de votre add-in</span><span class="sxs-lookup"><span data-stu-id="f91e4-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="f91e4-207">Plusieurs tâches de build sont disponibles :</span><span class="sxs-lookup"><span data-stu-id="f91e4-207">There are several build tasks available:</span></span>
- <span data-ttu-id="f91e4-208">`npm run watch`: se construit pour le développement et se reconstruit automatiquement lorsqu'un fichier source est enregistré</span><span class="sxs-lookup"><span data-stu-id="f91e4-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="f91e4-209">`npm run build-dev`: builds pour le développement une fois</span><span class="sxs-lookup"><span data-stu-id="f91e4-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="f91e4-210">`npm run build`: builds pour la production</span><span class="sxs-lookup"><span data-stu-id="f91e4-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="f91e4-211">`npm run dev-server`: exécute le serveur web utilisé pour le développement</span><span class="sxs-lookup"><span data-stu-id="f91e4-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="f91e4-212">Vous pouvez utiliser les tâches suivantes pour démarrer le débogage sur un ordinateur de bureau ou en ligne.</span><span class="sxs-lookup"><span data-stu-id="f91e4-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="f91e4-213">`npm run start:desktop`: démarre Excel sur le bureau et charge une version de version de votre application.</span><span class="sxs-lookup"><span data-stu-id="f91e4-213">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="f91e4-214">`npm run start:web`: démarre Excel sur le web et charge une version de version de votre application.</span><span class="sxs-lookup"><span data-stu-id="f91e4-214">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="f91e4-215">`npm run stop`: arrête Excel et le débogage.</span><span class="sxs-lookup"><span data-stu-id="f91e4-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f91e4-216">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="f91e4-216">Next steps</span></span>
<span data-ttu-id="f91e4-217">Découvrez les [pratiques d'authentification](custom-functions-authentication.md)pour les fonctions personnalisées sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f91e4-217">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f91e4-218">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f91e4-218">See also</span></span>

* [<span data-ttu-id="f91e4-219">Résolution des problèmes des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="f91e4-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="f91e4-220">Gestion des erreurs liées aux fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="f91e4-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="f91e4-221">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="f91e4-221">Create custom functions in Excel</span></span>](custom-functions-overview.md)
