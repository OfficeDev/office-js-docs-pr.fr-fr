---
ms.date: 07/10/2019
description: Déboguez vos fonctions personnalisées dans Excel.
title: Débogage des fonctions personnalisées
localization_priority: Normal
ms.openlocfilehash: 4abd5f3da58c35485004b17f92b334b133cabd27
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719307"
---
# <a name="custom-functions-debugging"></a><span data-ttu-id="374bc-103">Débogage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="374bc-103">Custom functions debugging</span></span>

<span data-ttu-id="374bc-104">Le débogage des fonctions personnalisées peut être réalisé de plusieurs manières, en fonction de la plateforme que vous utilisez.</span><span class="sxs-lookup"><span data-stu-id="374bc-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="374bc-105">Sur Windows :</span><span class="sxs-lookup"><span data-stu-id="374bc-105">On Windows:</span></span>
- [<span data-ttu-id="374bc-106">Débogueur de code Visual Studio et de bureau Excel (code VS)</span><span class="sxs-lookup"><span data-stu-id="374bc-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="374bc-107">Excel sur le Web et le débogueur de code VS</span><span class="sxs-lookup"><span data-stu-id="374bc-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="374bc-108">Excel sur le Web et les outils de navigation</span><span class="sxs-lookup"><span data-stu-id="374bc-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="374bc-109">Ligne de commande</span><span class="sxs-lookup"><span data-stu-id="374bc-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="374bc-110">Sur Mac :</span><span class="sxs-lookup"><span data-stu-id="374bc-110">On Mac:</span></span>
- [<span data-ttu-id="374bc-111">Excel sur le Web et les outils de navigation</span><span class="sxs-lookup"><span data-stu-id="374bc-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="374bc-112">Ligne de commande</span><span class="sxs-lookup"><span data-stu-id="374bc-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="374bc-113">Par souci de simplicité, cet article présente le débogage dans le contexte de l’utilisation de Visual Studio code pour modifier, exécuter des tâches et, dans certains cas, utiliser l’affichage débogage.</span><span class="sxs-lookup"><span data-stu-id="374bc-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="374bc-114">Si vous utilisez un autre éditeur ou outil de ligne de commande, consultez les [instructions de ligne de commande](#commands-for-building-and-running-your-add-in) à la fin de cet article.</span><span class="sxs-lookup"><span data-stu-id="374bc-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="374bc-115">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="374bc-115">Requirements</span></span>

<span data-ttu-id="374bc-116">Avant de commencer le débogage, vous devez utiliser le [Générateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office) afin de créer un projet de fonctions personnalisées.</span><span class="sxs-lookup"><span data-stu-id="374bc-116">Before starting to debug, you should use the [Yeoman generator for Office add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project.</span></span> <span data-ttu-id="374bc-117">Pour obtenir des instructions sur la création d’un projet de fonctions personnalisées, consultez le didacticiel sur les [fonctions personnalisées](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="374bc-117">For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="374bc-118">Utiliser le débogueur de code VS pour le bureau Excel</span><span class="sxs-lookup"><span data-stu-id="374bc-118">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="374bc-119">Vous pouvez utiliser le code VS pour déboguer des fonctions personnalisées dans Office Excel sur le bureau.</span><span class="sxs-lookup"><span data-stu-id="374bc-119">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="374bc-120">Le débogage de bureau pour Mac n’est pas disponible, mais peut être réalisé [à l’aide des outils de navigation et de la ligne de commande pour déboguer Excel sur le Web](#use-the-command-line-tools-to-debug).</span><span class="sxs-lookup"><span data-stu-id="374bc-120">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="374bc-121">Exécuter votre complément à partir du code VS</span><span class="sxs-lookup"><span data-stu-id="374bc-121">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="374bc-122">Ouvrez votre dossier de projet racine de fonctions personnalisées dans le [code vs](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="374bc-122">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="374bc-123">Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **Espion**.</span><span class="sxs-lookup"><span data-stu-id="374bc-123">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="374bc-124">Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="374bc-124">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="374bc-125">Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.</span><span class="sxs-lookup"><span data-stu-id="374bc-125">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="374bc-126">Démarrer le débogueur de code VS</span><span class="sxs-lookup"><span data-stu-id="374bc-126">Start the VS Code debugger</span></span>

4. <span data-ttu-id="374bc-127">Sélectionnez **afficher > déboguer** ou **Appuyez sur Ctrl + Maj + D** pour basculer vers l’affichage débogage.</span><span class="sxs-lookup"><span data-stu-id="374bc-127">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="374bc-128">Dans les options de débogage, choisissez **bureau Excel**.</span><span class="sxs-lookup"><span data-stu-id="374bc-128">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="374bc-129">Sélectionnez **F5** (ou choisissez **Déboguer-> démarrer le débogage** dans le menu) pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="374bc-129">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="374bc-130">Un nouveau classeur Excel s’ouvre avec votre complément déjà versions test chargées et prêt à être utilisé.</span><span class="sxs-lookup"><span data-stu-id="374bc-130">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="374bc-131">Démarrer le débogage</span><span class="sxs-lookup"><span data-stu-id="374bc-131">Start debugging</span></span>

1. <span data-ttu-id="374bc-132">Dans le code VS, ouvrez votre fichier de script de code source (**functions. js** ou **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="374bc-132">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="374bc-133">[Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="374bc-133">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="374bc-134">Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="374bc-134">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="374bc-135">À ce stade, l’exécution s’arrêtera sur la ligne de code où vous définissez le point d’arrêt.</span><span class="sxs-lookup"><span data-stu-id="374bc-135">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="374bc-136">À présent, vous pouvez parcourir votre code, définir des montres et utiliser les fonctionnalités de débogage de code VS dont vous avez besoin.</span><span class="sxs-lookup"><span data-stu-id="374bc-136">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="374bc-137">Utiliser le débogueur de code VS pour Excel dans Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="374bc-137">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="374bc-138">Vous pouvez utiliser le code VS pour déboguer des fonctions personnalisées dans Excel dans le navigateur Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="374bc-138">You can use VS Code to debug custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="374bc-139">Pour utiliser le code VS avec Microsoft Edge, vous devez installer le [débogueur pour l’extension Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .</span><span class="sxs-lookup"><span data-stu-id="374bc-139">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="374bc-140">Exécuter votre complément à partir du code VS</span><span class="sxs-lookup"><span data-stu-id="374bc-140">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="374bc-141">Ouvrez votre dossier de projet racine de fonctions personnalisées dans le [code vs](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="374bc-141">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="374bc-142">Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **Espion**.</span><span class="sxs-lookup"><span data-stu-id="374bc-142">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="374bc-143">Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="374bc-143">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="374bc-144">Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.</span><span class="sxs-lookup"><span data-stu-id="374bc-144">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="374bc-145">Démarrer le débogueur de code VS</span><span class="sxs-lookup"><span data-stu-id="374bc-145">Start the VS Code debugger</span></span>

4. <span data-ttu-id="374bc-146">Sélectionnez **afficher > déboguer** ou **Appuyez sur Ctrl + Maj + D** pour basculer vers l’affichage débogage.</span><span class="sxs-lookup"><span data-stu-id="374bc-146">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="374bc-147">Dans les options de débogage, sélectionnez **Office Online (Microsoft Edge)**.</span><span class="sxs-lookup"><span data-stu-id="374bc-147">From the Debug options, choose **Office Online (Microsoft Edge)**.</span></span>
6. <span data-ttu-id="374bc-148">Ouvrez Excel dans le navigateur Microsoft Edge et créez un classeur.</span><span class="sxs-lookup"><span data-stu-id="374bc-148">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="374bc-149">Choisissez **partager** dans le ruban et copiez le lien de l’URL de ce nouveau classeur.</span><span class="sxs-lookup"><span data-stu-id="374bc-149">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="374bc-150">Sélectionnez **F5** (ou choisissez **déboguer > démarrer le débogage** dans le menu) pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="374bc-150">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="374bc-151">Une invite s’affiche, qui vous demande l’URL de votre document.</span><span class="sxs-lookup"><span data-stu-id="374bc-151">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="374bc-152">Collez l’URL de votre classeur, puis appuyez sur entrée.</span><span class="sxs-lookup"><span data-stu-id="374bc-152">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="374bc-153">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="374bc-153">Sideload your add-in</span></span>

1. <span data-ttu-id="374bc-154">Sélectionnez l’onglet **Insérer** dans le ruban, puis dans la section **compléments** , choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="374bc-154">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="374bc-155">Dans la boîte de dialogue **Compléments Office** , sélectionnez l’onglet **mes compléments** , choisissez **gérer mes compléments**, puis **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="374bc-155">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="374bc-157">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="374bc-157">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="374bc-159">Définir des points d’arrêt</span><span class="sxs-lookup"><span data-stu-id="374bc-159">Set breakpoints</span></span>
1. <span data-ttu-id="374bc-160">Dans le code VS, ouvrez votre fichier de script de code source (**functions. js** ou **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="374bc-160">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="374bc-161">[Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="374bc-161">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="374bc-162">Dans le classeur Excel, entrez une formule qui utilise votre fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="374bc-162">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="374bc-163">Utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées dans Excel sur le Web</span><span class="sxs-lookup"><span data-stu-id="374bc-163">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="374bc-164">Vous pouvez utiliser les outils de développement de navigateur pour déboguer des fonctions personnalisées dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="374bc-164">You can use the browser developer tools to debug custom functions in Excel on the web.</span></span> <span data-ttu-id="374bc-165">Les étapes suivantes fonctionnent pour Windows et macOS.</span><span class="sxs-lookup"><span data-stu-id="374bc-165">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="374bc-166">Exécuter votre complément à partir de Visual Studio code</span><span class="sxs-lookup"><span data-stu-id="374bc-166">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="374bc-167">Ouvrez votre dossier de projet racine de fonctions personnalisées dans [Visual Studio code (vs code)](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="374bc-167">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="374bc-168">Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **Espion**.</span><span class="sxs-lookup"><span data-stu-id="374bc-168">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="374bc-169">Cette opération permet de surveiller et de reconstruire les modifications apportées aux fichiers.</span><span class="sxs-lookup"><span data-stu-id="374bc-169">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="374bc-170">Choisissez **Terminal > exécuter la tâche** , puis tapez ou sélectionnez **serveur de développement**.</span><span class="sxs-lookup"><span data-stu-id="374bc-170">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="374bc-171">Charger une version test de votre complément</span><span class="sxs-lookup"><span data-stu-id="374bc-171">Sideload your add-in</span></span>

1. <span data-ttu-id="374bc-172">Ouvrez [Microsoft Office sur le web](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="374bc-172">Open [Microsoft Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="374bc-173">Ouvrez un nouveau classeur Excel.</span><span class="sxs-lookup"><span data-stu-id="374bc-173">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="374bc-174">Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **compléments** , choisissez **Compléments Office**.</span><span class="sxs-lookup"><span data-stu-id="374bc-174">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="374bc-175">Dans la boîte de dialogue **Compléments Office** , sélectionnez l’onglet **mes compléments** , choisissez **gérer mes compléments**, puis **Télécharger mon complément**.</span><span class="sxs-lookup"><span data-stu-id="374bc-175">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="374bc-177">**Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.</span><span class="sxs-lookup"><span data-stu-id="374bc-177">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="374bc-179">Une fois que vous avez versions test chargées dans le document, il reste versions test chargées chaque fois que vous ouvrez le document.</span><span class="sxs-lookup"><span data-stu-id="374bc-179">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="374bc-180">Démarrer le débogage</span><span class="sxs-lookup"><span data-stu-id="374bc-180">Start debugging</span></span>

1. <span data-ttu-id="374bc-181">Ouvrez outils de développement dans le navigateur.</span><span class="sxs-lookup"><span data-stu-id="374bc-181">Open developer tools in the browser.</span></span> <span data-ttu-id="374bc-182">Pour le chrome et la plupart des navigateurs F12 ouvre les outils de développement.</span><span class="sxs-lookup"><span data-stu-id="374bc-182">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="374bc-183">Dans outils de développement, ouvrez votre fichier de script de code source à l’aide de **cmd + p** ou **Ctrl + p** (**functions. js** ou **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="374bc-183">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="374bc-184">[Définissez un point d’arrêt](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) dans le code source de la fonction personnalisée.</span><span class="sxs-lookup"><span data-stu-id="374bc-184">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="374bc-185">Si vous devez modifier le code, vous pouvez effectuer des modifications dans le code VS et enregistrer les modifications.</span><span class="sxs-lookup"><span data-stu-id="374bc-185">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="374bc-186">Actualisez le navigateur pour voir les modifications chargées.</span><span class="sxs-lookup"><span data-stu-id="374bc-186">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="374bc-187">Utiliser les outils de ligne de commande pour déboguer</span><span class="sxs-lookup"><span data-stu-id="374bc-187">Use the command line tools to debug</span></span>

<span data-ttu-id="374bc-188">Si vous n’utilisez pas le code VS, vous pouvez utiliser la ligne de commande (par exemple, bash ou PowerShell) pour exécuter votre complément.</span><span class="sxs-lookup"><span data-stu-id="374bc-188">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="374bc-189">Vous devrez utiliser les outils de développement de navigateur pour déboguer votre code dans Excel sur le Web.</span><span class="sxs-lookup"><span data-stu-id="374bc-189">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="374bc-190">Vous ne pouvez pas déboguer la version de bureau d’Excel à l’aide de la ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="374bc-190">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="374bc-191">À partir de la ligne `npm run watch` de commande, exécutez le suivi et la régénération lorsque les modifications du code se produisent.</span><span class="sxs-lookup"><span data-stu-id="374bc-191">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="374bc-192">Ouvrir une deuxième fenêtre de ligne de commande (la première est bloquée lors de l’exécution de la fonction espion).</span><span class="sxs-lookup"><span data-stu-id="374bc-192">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="374bc-193">Si vous souhaitez démarrer votre complément dans la version de bureau d’Excel, exécutez la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="374bc-193">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="374bc-194">Ou si vous préférez démarrer votre complément dans Excel sur le Web, exécutez la commande suivante :</span><span class="sxs-lookup"><span data-stu-id="374bc-194">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="374bc-195">Pour Excel sur le Web, vous devez également chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="374bc-195">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="374bc-196">Suivez les étapes décrites dans [chargement votre complément](#sideload-your-add-in) pour chargement votre complément.</span><span class="sxs-lookup"><span data-stu-id="374bc-196">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="374bc-197">Ensuite, passez à la section suivante pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="374bc-197">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="374bc-198">Ouvrez outils de développement dans le navigateur.</span><span class="sxs-lookup"><span data-stu-id="374bc-198">Open developer tools in the browser.</span></span> <span data-ttu-id="374bc-199">Pour le chrome et la plupart des navigateurs F12 ouvre les outils de développement.</span><span class="sxs-lookup"><span data-stu-id="374bc-199">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="374bc-200">Dans outils de développement, ouvrez votre fichier de script de code source (**functions. js** ou **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="374bc-200">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="374bc-201">Votre code de fonctions personnalisées peut être situé à la fin du fichier.</span><span class="sxs-lookup"><span data-stu-id="374bc-201">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="374bc-202">Dans le code source de la fonction personnalisée, appliquez un point d’arrêt en sélectionnant une ligne de code.</span><span class="sxs-lookup"><span data-stu-id="374bc-202">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="374bc-203">Si vous devez modifier le code, vous pouvez apporter des modifications dans Visual Studio et enregistrer les modifications.</span><span class="sxs-lookup"><span data-stu-id="374bc-203">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="374bc-204">Actualisez le navigateur pour voir les modifications chargées.</span><span class="sxs-lookup"><span data-stu-id="374bc-204">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="374bc-205">Commandes pour la création et l’exécution de votre complément</span><span class="sxs-lookup"><span data-stu-id="374bc-205">Commands for building and running your add-in</span></span>

<span data-ttu-id="374bc-206">Plusieurs tâches de génération sont disponibles :</span><span class="sxs-lookup"><span data-stu-id="374bc-206">There are several build tasks available:</span></span>
- <span data-ttu-id="374bc-207">`npm run watch`: builds pour le développement et rebuilds automatiques lors de l’enregistrement d’un fichier source</span><span class="sxs-lookup"><span data-stu-id="374bc-207">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="374bc-208">`npm run build-dev`: builds pour le développement une seule fois</span><span class="sxs-lookup"><span data-stu-id="374bc-208">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="374bc-209">`npm run build`: builds pour la production</span><span class="sxs-lookup"><span data-stu-id="374bc-209">`npm run build`: builds for production</span></span>
- <span data-ttu-id="374bc-210">`npm run dev-server`: exécute le serveur Web utilisé pour le développement</span><span class="sxs-lookup"><span data-stu-id="374bc-210">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="374bc-211">Vous pouvez utiliser les tâches suivantes pour démarrer le débogage sur le bureau ou en ligne.</span><span class="sxs-lookup"><span data-stu-id="374bc-211">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="374bc-212">`npm run start:desktop`: Démarre Excel sur le bureau et sideloads votre complément.</span><span class="sxs-lookup"><span data-stu-id="374bc-212">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="374bc-213">`npm run start:web`: Démarre Excel sur le Web et sideloads votre complément.</span><span class="sxs-lookup"><span data-stu-id="374bc-213">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="374bc-214">`npm run stop`: Arrête Excel et le débogage.</span><span class="sxs-lookup"><span data-stu-id="374bc-214">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="374bc-215">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="374bc-215">Next steps</span></span>
<span data-ttu-id="374bc-216">Découvrez les [pratiques d’authentification dans les fonctions personnalisées](custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="374bc-216">Learn about [authentication practices in custom functions](custom-functions-authentication.md).</span></span> <span data-ttu-id="374bc-217">Ou, examinez [l’architecture unique de la fonction personnalisée](custom-functions-architecture.md).</span><span class="sxs-lookup"><span data-stu-id="374bc-217">Or, review [custom function's unique architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="374bc-218">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="374bc-218">See also</span></span>

* [<span data-ttu-id="374bc-219">Dépannage des fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="374bc-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="374bc-220">Gestion des erreurs liées aux fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="374bc-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="374bc-221">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="374bc-221">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="374bc-222">Créer des fonctions personnalisées dans Excel</span><span class="sxs-lookup"><span data-stu-id="374bc-222">Create custom functions in Excel</span></span>](custom-functions-overview.md)
