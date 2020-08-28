---
title: Complément Microsoft Office Extension de débogueur pour Visual Studio Code
description: Utilisez le débogueur de complément Microsoft Office de l’extension de code Visual Studio pour déboguer votre complément Office.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 1343014fa875509fd6f0c615c3504cc9ae50dc0d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293442"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="7ea63-103">Complément Microsoft Office Extension de débogueur pour Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="7ea63-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="7ea63-104">L’extension du débogueur de complément Microsoft Office pour Visual Studio code vous permet de déboguer votre complément Office par rapport au runtime Edge.</span><span class="sxs-lookup"><span data-stu-id="7ea63-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Edge runtime.</span></span>

<span data-ttu-id="7ea63-105">Ce mode de débogage est dynamique, ce qui vous permet de définir des points d’arrêt lors de l’exécution du code.</span><span class="sxs-lookup"><span data-stu-id="7ea63-105">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="7ea63-106">Vous pouvez voir les modifications apportées à votre code immédiatement lorsque le débogueur est attaché, tout cela sans perdre votre session de débogage.</span><span class="sxs-lookup"><span data-stu-id="7ea63-106">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="7ea63-107">Les modifications apportées au code sont également conservées, ce qui vous permet de voir les résultats de plusieurs modifications apportées à votre code.</span><span class="sxs-lookup"><span data-stu-id="7ea63-107">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="7ea63-108">L’image suivante illustre cette extension en action.</span><span class="sxs-lookup"><span data-stu-id="7ea63-108">The following image shows this extension in action.</span></span>

![Extension de débogage du complément Office AddIn débogage d’une section de compléments Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="7ea63-110">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ea63-110">Prerequisites</span></span>

- <span data-ttu-id="7ea63-111">[Visual Studio code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)</span><span class="sxs-lookup"><span data-stu-id="7ea63-111">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="7ea63-112">Node.js (version 10 +)</span><span class="sxs-lookup"><span data-stu-id="7ea63-112">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="7ea63-113">Windows 10</span><span class="sxs-lookup"><span data-stu-id="7ea63-113">Windows 10</span></span>
- [<span data-ttu-id="7ea63-114">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="7ea63-114">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="7ea63-115">Ces instructions supposent que vous avez une expérience en utilisant la ligne de commande, que vous compreniez JavaScript de base et que vous avez créé un projet de complément Office avant d’utiliser le générateur Yo Office.</span><span class="sxs-lookup"><span data-stu-id="7ea63-115">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="7ea63-116">Si vous ne l’avez pas encore fait, songez à consulter l’un de nos didacticiels, comme le [didacticiel sur les compléments Office Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="7ea63-116">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="7ea63-117">Installer et utiliser le débogueur</span><span class="sxs-lookup"><span data-stu-id="7ea63-117">Install and use the debugger</span></span>

1. <span data-ttu-id="7ea63-118">Si vous devez créer un projet de complément, [Utilisez le générateur Yo Office pour en créer un](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span><span class="sxs-lookup"><span data-stu-id="7ea63-118">If you need to create an add-in project, [use the Yo Office generator to create one](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span></span> <span data-ttu-id="7ea63-119">Suivez les invites de la ligne de commande pour configurer votre projet.</span><span class="sxs-lookup"><span data-stu-id="7ea63-119">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="7ea63-120">Vous pouvez choisir n’importe quelle langue ou type de projet en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="7ea63-120">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="7ea63-121">Si vous disposez déjà d’un projet, ignorez l’étape 1 et passez à l’étape 2.</span><span class="sxs-lookup"><span data-stu-id="7ea63-121">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="7ea63-122">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="7ea63-122">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="7ea63-123">![Options d’invite de commandes, y compris « exécuter en tant qu’administrateur » dans Windows 10](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="7ea63-123">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="7ea63-124">Naviguez jusqu’au répertoire de votre projet.</span><span class="sxs-lookup"><span data-stu-id="7ea63-124">Navigate to your project directory.</span></span>

4. <span data-ttu-id="7ea63-125">Exécutez la commande suivante pour ouvrir votre projet dans Visual Studio code en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="7ea63-125">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="7ea63-126">Une fois Visual Studio code ouvert, accédez manuellement au dossier du projet.</span><span class="sxs-lookup"><span data-stu-id="7ea63-126">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="7ea63-127">Pour ouvrir Visual Studio code en tant qu’administrateur, sélectionnez l’option **exécuter en tant qu’administrateur** lors de l’ouverture de Visual Studio code après avoir effectué une recherche dans Windows.</span><span class="sxs-lookup"><span data-stu-id="7ea63-127">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="7ea63-128">Dans le code VS, sélectionnez **Ctrl + Maj + X** pour ouvrir la barre extensions.</span><span class="sxs-lookup"><span data-stu-id="7ea63-128">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="7ea63-129">Recherchez l’extension « Microsoft Office Add-in Debugger » et installez-la.</span><span class="sxs-lookup"><span data-stu-id="7ea63-129">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="7ea63-130">Dans le dossier. vscode de votre projet, ouvrez le fichier **launch.js** .</span><span class="sxs-lookup"><span data-stu-id="7ea63-130">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="7ea63-131">Ajoutez le code suivant à la `configurations` section :</span><span class="sxs-lookup"><span data-stu-id="7ea63-131">Add the following code to the `configurations` section:</span></span>

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. <span data-ttu-id="7ea63-132">Dans la section de JSON que vous venez de copier, recherchez la section « URL ».</span><span class="sxs-lookup"><span data-stu-id="7ea63-132">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="7ea63-133">Dans cette URL, vous devrez remplacer le texte d’hôte en majuscules par l’application qui héberge votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="7ea63-133">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office add-in.</span></span> <span data-ttu-id="7ea63-134">Par exemple, si votre complément Office est destiné à Excel, la valeur de votre URL serait « https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $16.01 $ en-US $ \$ \$ \$ 0 ».</span><span class="sxs-lookup"><span data-stu-id="7ea63-134">For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="7ea63-135">Ouvrez l’invite de commandes et assurez-vous que vous vous trouvez dans le dossier racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="7ea63-135">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="7ea63-136">Exécutez la commande `npm start` pour démarrer le serveur de développement.</span><span class="sxs-lookup"><span data-stu-id="7ea63-136">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="7ea63-137">Lorsque votre complément est chargé dans le client Office, ouvrez le volet de tâches.</span><span class="sxs-lookup"><span data-stu-id="7ea63-137">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="7ea63-138">Revenez à Visual Studio code et choisissez **view > Debug** ou **Appuyez sur Ctrl + Maj + D** pour basculer vers le mode débogage.</span><span class="sxs-lookup"><span data-stu-id="7ea63-138">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="7ea63-139">Dans les options de débogage, choisissez **attacher aux compléments Office**. Sélectionnez **F5** ou choisissez **Déboguer-> démarrer le débogage** dans le menu pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="7ea63-139">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="7ea63-140">Définissez un point d’arrêt dans le fichier de volet Office de votre projet.</span><span class="sxs-lookup"><span data-stu-id="7ea63-140">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="7ea63-141">Vous pouvez définir des points d’arrêt dans le code VS en plaçant le curseur en regard d’une ligne de code et en sélectionnant le cercle rouge qui apparaît.</span><span class="sxs-lookup"><span data-stu-id="7ea63-141">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![Un cercle rouge apparaît sur une ligne de code dans un code VS](../images/set-breakpoint.jpg)

12. <span data-ttu-id="7ea63-143">Exécutez votre complément.</span><span class="sxs-lookup"><span data-stu-id="7ea63-143">Run your add-in.</span></span> <span data-ttu-id="7ea63-144">Vous verrez que des points d’arrêt ont été atteints et que vous pouvez inspecter les variables locales.</span><span class="sxs-lookup"><span data-stu-id="7ea63-144">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="7ea63-145">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="7ea63-145">See also</span></span>

* [<span data-ttu-id="7ea63-146">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="7ea63-146">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="7ea63-147">Débogage des compléments avec les outils de développement sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="7ea63-147">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="7ea63-148">Attacher un débogueur à partir du volet Office</span><span class="sxs-lookup"><span data-stu-id="7ea63-148">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
