---
ms.date: 03/13/2019
description: DéBoguez vos fonctions personnalisées dans Excel.
title: Débogage des fonctions personnalisées (aperçu)
localization_priority: Normal
ms.openlocfilehash: 08563ef630ebc457219c4c622328b84d13e6acab
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448754"
---
# <a name="custom-functions-debugging-preview"></a><span data-ttu-id="da8f7-103">Débogage des fonctions personnalisées (aperçu)</span><span class="sxs-lookup"><span data-stu-id="da8f7-103">Custom functions debugging (preview)</span></span>

<span data-ttu-id="da8f7-104">Le débogage des fonctions personnalisées peut être réalisé de plusieurs manières, en fonction de la plateforme que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="da8f7-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

<span data-ttu-id="da8f7-105">Sur Windows:</span><span class="sxs-lookup"><span data-stu-id="da8f7-105">On Windows:</span></span>
- [<span data-ttu-id="da8f7-106">Débogueur de code Visual Studio et de bureau Excel (code VS)</span><span class="sxs-lookup"><span data-stu-id="da8f7-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="da8f7-107">Excel Online et le débogueur de code VS</span><span class="sxs-lookup"><span data-stu-id="da8f7-107">Excel Online and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [<span data-ttu-id="da8f7-108">Outils de navigation et Excel Online</span><span class="sxs-lookup"><span data-stu-id="da8f7-108">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="da8f7-109">Ligne de commande</span><span class="sxs-lookup"><span data-stu-id="da8f7-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="da8f7-110">Sur Mac:</span><span class="sxs-lookup"><span data-stu-id="da8f7-110">On Mac:</span></span>
- [<span data-ttu-id="da8f7-111">Outils de navigation et Excel Online</span><span class="sxs-lookup"><span data-stu-id="da8f7-111">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="da8f7-112">Ligne de commande</span><span class="sxs-lookup"><span data-stu-id="da8f7-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> <span data-ttu-id="da8f7-113">Par souci de simplicité, cet article présente le débogage dans le contexte de l'utilisation de Visual Studio code pour modifier, exécuter des tâches et, dans certains cas, utiliser l'affichage débogage.</span><span class="sxs-lookup"><span data-stu-id="da8f7-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="da8f7-114">Si vous utilisez un autre éditeur ou outil de ligne de commande, consultez les [instructions de ligne de commande](#use-the-command-line-tools-to-debug) à la fin de cet article.</span><span class="sxs-lookup"><span data-stu-id="da8f7-114">If you are using a different editor or command line tool, see the [command line instructions](#use-the-command-line-tools-to-debug) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="da8f7-115">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="da8f7-115">Requirements</span></span>

<span data-ttu-id="da8f7-116">Avant de commencer le débogage, vous devez créer un projet de complément de fonctions personnalisées à l'aide du générateur Yo Office et vous assurer que vous disposez de certificats auto-signés approuvés pour votre projet.</span><span class="sxs-lookup"><span data-stu-id="da8f7-116">Before starting to debug, you should create a custom functions add-in project using the Yo Office generator and ensured that you have trusted self-signed certificates for your project.</span></span> <span data-ttu-id="da8f7-117">Pour obtenir des instructions sur la création d'un projet, consultez le didacticiel sur les [fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="da8f7-117">For instructions to create a project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span> <span data-ttu-id="da8f7-118">Pour obtenir des instructions sur l'approbation des certificats, consultez la rubrique [Ajout de certificats auto-signés en tant que certificats racines approuvés](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="da8f7-118">For instructions on trusting certificates, see [Adding self-signed certificates as trusted root certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="da8f7-119">Utiliser le débogueur de code VS pour le bureau Excel</span><span class="sxs-lookup"><span data-stu-id="da8f7-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="da8f7-120">Vous pouvez utiliser le code VS pour déboguer des fonctions personnalisées dans Office Excel sur le bureau.</span><span class="sxs-lookup"><span data-stu-id="da8f7-120">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="da8f7-121">Le débogage de bureau pour Mac n'est pas disponible, mais peut être réalisé [à l'aide des outils de navigation pour déboguEr Excel Online](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online).</span><span class="sxs-lookup"><span data-stu-id="da8f7-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools to debug Excel Online](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="da8f7-122">Exécuter votre complément à partir du code VS</span><span class="sxs-lookup"><span data-stu-id="da8f7-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="da8f7-123">Ouvrez votre dossier de projet racine de fonctions personnalisées dans le [code vs](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="da8f7-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="da8f7-124">Choisissez **Terminal _GT_ exécuter la tâche** , puis tapez ou sélectionnez **Espion**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="da8f7-125">Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="da8f7-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="da8f7-126">Choisissez **Terminal _GT_ exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="da8f7-127">Démarrer le débogueur de code VS</span><span class="sxs-lookup"><span data-stu-id="da8f7-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="da8f7-128">Sélectionnez **Afficher >** déboguer ou **Appuyez sur Ctrl + Maj + D** pour basculer vers l'affichage débogage.</span><span class="sxs-lookup"><span data-stu-id="da8f7-128">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="da8f7-129">Dans les options de déBogage, choisissez **bureau Excel**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-129">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="da8f7-130">Sélectionnez **F5** (ou choisissez débogage **-> démarrer** le débogage dans le menu) pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="da8f7-130">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="da8f7-131">Un nouveau classeur Excel s'ouvre avec votre complément déjà versions test chargées et prêt à être utilisé.</span><span class="sxs-lookup"><span data-stu-id="da8f7-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="da8f7-132">Démarrer le débogage</span><span class="sxs-lookup"><span data-stu-id="da8f7-132">Start debugging</span></span>

1. <span data-ttu-id="da8f7-133">Dans le code VS, ouvrez votre fichier de script de code source (functions. js ou functions. TS).</span><span class="sxs-lookup"><span data-stu-id="da8f7-133">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="da8f7-134">[Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="da8f7-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="da8f7-135">Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="da8f7-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="da8f7-136">À ce stade, l'exécution s'arrêtera sur la ligne de code où vous définissez le point d'arrêt.</span><span class="sxs-lookup"><span data-stu-id="da8f7-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="da8f7-137">À présent, vous pouvez parcourir votre code, définir des montres et utiliser les fonctionnalités de débogage de code VS dont vous avez besoin.</span><span class="sxs-lookup"><span data-stu-id="da8f7-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a><span data-ttu-id="da8f7-138">Utiliser le débogueur de code VS pour Excel Online dans Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="da8f7-138">Use the VS Code debugger for Excel Online in Microsoft Edge</span></span>

<span data-ttu-id="da8f7-139">Vous pouvez utiliser le code VS pour déboguer des fonctions personnalisées dans Excel Online dans le navigateur Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="da8f7-139">You can use VS Code to debug custom functions in Excel Online in the Microsoft Edge browser.</span></span> <span data-ttu-id="da8f7-140">Pour utiliser le code VS avec Microsoft Edge, vous devez installer le débogueur pour l'extension [Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .</span><span class="sxs-lookup"><span data-stu-id="da8f7-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="da8f7-141">Exécuter votre complément à partir du code VS</span><span class="sxs-lookup"><span data-stu-id="da8f7-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="da8f7-142">Ouvrez votre dossier de projet racine de fonctions personnalisées dans le [code vs](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="da8f7-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="da8f7-143">Choisissez **Terminal _GT_ exécuter la tâche** , puis tapez ou sélectionnez **Espion**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="da8f7-144">Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="da8f7-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="da8f7-145">Choisissez **Terminal _GT_ exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="da8f7-146">Démarrer le débogueur de code VS</span><span class="sxs-lookup"><span data-stu-id="da8f7-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="da8f7-147">Sélectionnez **Afficher >** déboguer ou **Appuyez sur Ctrl + Maj + D** pour basculer vers l'affichage débogage.</span><span class="sxs-lookup"><span data-stu-id="da8f7-147">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="da8f7-148">Dans les options de déBogage, sélectionnez **Office Online (Edge)**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-148">From the Debug options, choose **Office Online (Edge)**.</span></span>
6. <span data-ttu-id="da8f7-149">Ouvrez Excel Online à l'aide du navigateur Microsoft Edge, ouvrez Excel Online, puis créez un classeur.</span><span class="sxs-lookup"><span data-stu-id="da8f7-149">Open Excel Online using the Microsoft Edge browser, open Excel Online, and create a new workbook.</span></span>
7. <span data-ttu-id="da8f7-150">Choisissez **partager** dans le ruban et copiez le lien de l'URL de ce nouveau classeur.</span><span class="sxs-lookup"><span data-stu-id="da8f7-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="da8f7-151">Sélectionnez **F5** (ou choisissez **Déboguer > démarrer** le débogage dans le menu) pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="da8f7-151">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="da8f7-152">Une invite s'affiche, qui vous demande l'URL de votre document.</span><span class="sxs-lookup"><span data-stu-id="da8f7-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="da8f7-153">Collez l'URL de votre classeur, puis appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="da8f7-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="da8f7-154">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="da8f7-154">Sideload your add-in</span></span>   

1. <span data-ttu-id="da8f7-155">Sélectionnez l'onglet **Insérer** dans le ruban, puis dans la section **compléments** , choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-155">Select the  **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="da8f7-156">Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MES COMPLÉMENTS**, choisissez **Gérer mes compléments**, puis **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-156">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

3.  <span data-ttu-id="da8f7-158">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="da8f7-160">Définir des points d'arrêt</span><span class="sxs-lookup"><span data-stu-id="da8f7-160">Set breakpoints</span></span>
1. <span data-ttu-id="da8f7-161">Dans le code VS, ouvrez votre fichier de script de code source (functions. js ou functions. TS).</span><span class="sxs-lookup"><span data-stu-id="da8f7-161">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="da8f7-162">[Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="da8f7-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="da8f7-163">Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="da8f7-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a><span data-ttu-id="da8f7-164">Utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées dans Excel Online</span><span class="sxs-lookup"><span data-stu-id="da8f7-164">Use the browser developer tools to debug custom functions in Excel Online</span></span>

<span data-ttu-id="da8f7-165">Vous pouvez utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées dans Excel online.</span><span class="sxs-lookup"><span data-stu-id="da8f7-165">You can use the browser developer tools to debug custom functions in Excel Online.</span></span> <span data-ttu-id="da8f7-166">Les étapes suivantes fonctionnent pour Windows et macOS.</span><span class="sxs-lookup"><span data-stu-id="da8f7-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="da8f7-167">Exécuter votre complément à partir de Visual Studio code</span><span class="sxs-lookup"><span data-stu-id="da8f7-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="da8f7-168">Ouvrez votre dossier de projet racine de fonctions personnalisées dans [Visual Studio code (vs code)](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="da8f7-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="da8f7-169">Choisissez **Terminal _GT_ exécuter la tâche** , puis tapez ou sélectionnez **Espion**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="da8f7-170">Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="da8f7-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="da8f7-171">Choisissez **Terminal _GT_ exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="sideload-your-add-in"></a><span data-ttu-id="da8f7-172">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="da8f7-172">Sideload your add-in</span></span>   

1. <span data-ttu-id="da8f7-173">Ouvrez [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="da8f7-173">Open [Microsoft Office Online](https://office.live.com/).</span></span>
2. <span data-ttu-id="da8f7-174">Ouvrez un nouveau classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="da8f7-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="da8f7-175">Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **Compléments**, choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-175">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="da8f7-176">Dans la boîte de dialogue **Compléments Office**, sélectionnez l’onglet **MES COMPLÉMENTS**, choisissez **Gérer mes compléments**, puis **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-176">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="da8f7-178">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="da8f7-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="da8f7-180">Une fois que vous avez versions test chargées dans le document, il reste versions test chargées chaque fois que vous ouvrez le document.</span><span class="sxs-lookup"><span data-stu-id="da8f7-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="da8f7-181">Démarrer le débogage</span><span class="sxs-lookup"><span data-stu-id="da8f7-181">Start debugging</span></span>

1. <span data-ttu-id="da8f7-182">Ouvrez outils de développement dans le navigateur.</span><span class="sxs-lookup"><span data-stu-id="da8f7-182">Open developer tools in the browser.</span></span> <span data-ttu-id="da8f7-183">Pour le chrome et la plupart des navigateurs F12 ouvre les outils de développement.</span><span class="sxs-lookup"><span data-stu-id="da8f7-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="da8f7-184">Dans outils de développement, ouvrez votre fichier de script de code source à l'aide de **cmd + p** ou **Ctrl + p** (functions. js ou functions. TS).</span><span class="sxs-lookup"><span data-stu-id="da8f7-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (functions.js or functions.ts).</span></span>
3. <span data-ttu-id="da8f7-185">[Définissez un point d'arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="da8f7-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="da8f7-186">Si vous devez modifier le code, vous pouvez effectuer des modifications dans le code VS et enregistrer les modifications.</span><span class="sxs-lookup"><span data-stu-id="da8f7-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="da8f7-187">Actualisez le navigateur pour voir les modifications chargées.</span><span class="sxs-lookup"><span data-stu-id="da8f7-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="da8f7-188">Utiliser les outils de ligne de commande pour déboguer</span><span class="sxs-lookup"><span data-stu-id="da8f7-188">Use the command line tools to debug</span></span>

<span data-ttu-id="da8f7-189">Si vous n'utilisez pas le code VS, vous pouvez utiliser la ligne de commande (par exemple, bash ou PowerShell) pour exécuter votre complément.</span><span class="sxs-lookup"><span data-stu-id="da8f7-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="da8f7-190">Vous devrez utiliser les outils de développement de navigateur pour déboguer votre code dans Excel online.</span><span class="sxs-lookup"><span data-stu-id="da8f7-190">You'll need to use the browser developer tools to debug your code in Excel Online.</span></span> <span data-ttu-id="da8f7-191">Vous ne pouvez pas déboguer la version de bureau d'Excel à l'aide de la ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="da8f7-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="da8f7-192">À partir de la ligne `npm run watch` de commande, exécutez le suivi et la régénération lorsque les modifications du code se produisent.</span><span class="sxs-lookup"><span data-stu-id="da8f7-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="da8f7-193">Ouvrir une deuxième fenêtre de ligne de commande (la première est bloquée lors de l'exécution de la fonction espion).</span><span class="sxs-lookup"><span data-stu-id="da8f7-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="da8f7-194">Si vous souhaitez démarrer votre complément dans la version de bureau d'Excel, exécutez la commande suivante:</span><span class="sxs-lookup"><span data-stu-id="da8f7-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start desktop`
    
    <span data-ttu-id="da8f7-195">Ou si vous préférez démarrer votre complément dans Excel Online, exécutez la commande suivante:</span><span class="sxs-lookup"><span data-stu-id="da8f7-195">Or if you prefer to start your add-in in Excel Online run the following command</span></span>
    
    `npm run start web`
    
    <span data-ttu-id="da8f7-196">Pour Excel Online, vous devez également chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="da8f7-196">For Excel Online you also need to sideload your add-in.</span></span> <span data-ttu-id="da8f7-197">Suivez les étapes décrites dans [chargement votre complément](#sideload-your-add-in) pour chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="da8f7-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="da8f7-198">Ensuite, passez à la section suivante pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="da8f7-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="da8f7-199">Ouvrez outils de développement dans le navigateur.</span><span class="sxs-lookup"><span data-stu-id="da8f7-199">Open developer tools in the browser.</span></span> <span data-ttu-id="da8f7-200">Pour le chrome et la plupart des navigateurs F12 ouvre les outils de développement.</span><span class="sxs-lookup"><span data-stu-id="da8f7-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="da8f7-201">Dans outils de développement, ouvrez votre fichier de script de code source (functions. js ou functions. TS).</span><span class="sxs-lookup"><span data-stu-id="da8f7-201">In developer tools, open your source code script file (functions.js or functions.ts).</span></span> <span data-ttu-id="da8f7-202">Votre code de fonctions personnalisées peut être situé à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="da8f7-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="da8f7-203">Dans le code source de la fonction personnalisée, appliquez un point d'arrêt en sélectionnant une ligne de code.</span><span class="sxs-lookup"><span data-stu-id="da8f7-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="da8f7-204">Si vous devez modifier le code, vous pouvez apporter des modifications dans Visual Studio et enregistrer les modifications.</span><span class="sxs-lookup"><span data-stu-id="da8f7-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="da8f7-205">Actualisez le navigateur pour voir les modifications chargées.</span><span class="sxs-lookup"><span data-stu-id="da8f7-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="da8f7-206">Commandes pour la création et l'exécution de votre complément</span><span class="sxs-lookup"><span data-stu-id="da8f7-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="da8f7-207">Plusieurs tâches de génération sont disponibles:</span><span class="sxs-lookup"><span data-stu-id="da8f7-207">There are several build tasks available:</span></span>
- <span data-ttu-id="da8f7-208">`npm run watch`: builds pour le développement et rebuilds automatiques lors de l'enregistrement d'un fichier source</span><span class="sxs-lookup"><span data-stu-id="da8f7-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="da8f7-209">`npm run build-dev`: builds pour le développement une seule fois</span><span class="sxs-lookup"><span data-stu-id="da8f7-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="da8f7-210">`npm run build`: builds pour la production</span><span class="sxs-lookup"><span data-stu-id="da8f7-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="da8f7-211">`npm run dev-server`: exécute le serveur Web utilisé pour le développement</span><span class="sxs-lookup"><span data-stu-id="da8f7-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="da8f7-212">Vous pouvez utiliser les tâches suivantes pour démarrer le débogage sur le bureau ou en ligne.</span><span class="sxs-lookup"><span data-stu-id="da8f7-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="da8f7-213">`npm run start desktop`: Démarre Excel sur le bureau et sideloads votre complément.</span><span class="sxs-lookup"><span data-stu-id="da8f7-213">`npm run start desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="da8f7-214">`npm run start web`: Démarre Excel Online et sideloads votre complément.</span><span class="sxs-lookup"><span data-stu-id="da8f7-214">`npm run start web`: Starts Excel Online and sideloads your add-in.</span></span>
- <span data-ttu-id="da8f7-215">`npm run stop`: Arrête Excel et le débogage.</span><span class="sxs-lookup"><span data-stu-id="da8f7-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="see-also"></a><span data-ttu-id="da8f7-216">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="da8f7-216">See also</span></span>

* [<span data-ttu-id="da8f7-217">Métadonnées fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="da8f7-217">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="da8f7-218">Exécution de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="da8f7-218">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="da8f7-219">Meilleures pratiques de fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="da8f7-219">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="da8f7-220">Fonctions personnalisées changelog</span><span class="sxs-lookup"><span data-stu-id="da8f7-220">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="da8f7-221">Didacticiel de fonctions personnalisées Excel</span><span class="sxs-lookup"><span data-stu-id="da8f7-221">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
