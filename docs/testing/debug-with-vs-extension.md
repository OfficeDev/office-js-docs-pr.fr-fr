---
title: Complément Microsoft Office Extension de débogueur pour Visual Studio Code
description: Utilisez l’extension Visual Studio Code de Microsoft Office déboguer votre Office de débogage.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 3daedb48bdec5a17dfc220f049a8e2cdc86ac398
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349286"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="f50c8-103">Complément Microsoft Office Extension de débogueur pour Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f50c8-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="f50c8-104">L’extension de déboguer du Microsoft Office pour Visual Studio Code vous permet de déboguer votre Office Par rapport au Microsoft Edge avec le runtime WebView d’origine (EdgeHTML).</span><span class="sxs-lookup"><span data-stu-id="f50c8-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="f50c8-105">Pour obtenir des instructions sur le débogage Microsoft Edge WebView2 (basé sur Chromium web), consultez [cet article](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="f50c8-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="f50c8-106">Ce mode de débogage est dynamique, ce qui vous permet de définir des points d’arrêt pendant l’exécution du code.</span><span class="sxs-lookup"><span data-stu-id="f50c8-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="f50c8-107">Vous pouvez voir les modifications dans votre code immédiatement lorsque le déboguer est attaché, tout cela sans perdre votre session de débogage.</span><span class="sxs-lookup"><span data-stu-id="f50c8-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="f50c8-108">Vos modifications de code sont également persistantes, afin que vous pouvez voir les résultats de plusieurs modifications apportées à votre code.</span><span class="sxs-lookup"><span data-stu-id="f50c8-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="f50c8-109">L’image suivante illustre cette extension en action.</span><span class="sxs-lookup"><span data-stu-id="f50c8-109">The following image shows this extension in action.</span></span>

![Office Extension déboguer une section de l’extension déboguer Excel les autres.](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="f50c8-111">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f50c8-111">Prerequisites</span></span>

- <span data-ttu-id="f50c8-112">[Visual Studio Code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)</span><span class="sxs-lookup"><span data-stu-id="f50c8-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="f50c8-113">Node.js (version 10+)</span><span class="sxs-lookup"><span data-stu-id="f50c8-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="f50c8-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="f50c8-114">Windows 10</span></span>
- [<span data-ttu-id="f50c8-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="f50c8-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="f50c8-116">Ces instructions supposent que vous avez de l’expérience en utilisant la ligne de commande, que vous comprenez javaScript de base et que vous avez créé un projet de Office avant d’utiliser le générateur Yo Office.</span><span class="sxs-lookup"><span data-stu-id="f50c8-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="f50c8-117">Si vous ne l’avez pas encore fait, envisagez de consulter l’un de nos didacticiels, comme Excel Office [didacticiel sur le add-in.](../tutorials/excel-tutorial.md)</span><span class="sxs-lookup"><span data-stu-id="f50c8-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="f50c8-118">Installer et utiliser le débogueur</span><span class="sxs-lookup"><span data-stu-id="f50c8-118">Install and use the debugger</span></span>

1. <span data-ttu-id="f50c8-119">Si vous avez besoin de créer un projet de Office, utilisez le générateur [yo-Office pour en créer un.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)</span><span class="sxs-lookup"><span data-stu-id="f50c8-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="f50c8-120">Suivez les invites de la ligne de commande pour configurer votre projet.</span><span class="sxs-lookup"><span data-stu-id="f50c8-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="f50c8-121">Vous pouvez choisir n’importe quelle langue ou type de projet en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="f50c8-121">You can choose any language or type of project to suit your needs.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f50c8-122">Si vous avez déjà un projet, ignorez l’étape 1 et passez à l’étape 2.</span><span class="sxs-lookup"><span data-stu-id="f50c8-122">If you already have a project, skip step 1 and move to step 2.</span></span>

1. <span data-ttu-id="f50c8-123">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="f50c8-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="f50c8-124">![Options d’invite de commandes, y compris « Exécuter en tant qu’administrateur » Windows 10.](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="f50c8-124">![Command prompt options, including "run as administrator" in Windows 10.](../images/run-as-administrator-vs-code.jpg)</span></span>

1. <span data-ttu-id="f50c8-125">Accédez au répertoire de votre projet.</span><span class="sxs-lookup"><span data-stu-id="f50c8-125">Navigate to your project directory.</span></span>

1. <span data-ttu-id="f50c8-126">Exécutez la commande suivante pour ouvrir votre projet dans Visual Studio Code en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="f50c8-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

    ```command&nbsp;line
    code .
    ```

  <span data-ttu-id="f50c8-127">Une Visual Studio Code est ouverte, accédez manuellement au dossier du projet.</span><span class="sxs-lookup"><span data-stu-id="f50c8-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

  > [!TIP]
  > <span data-ttu-id="f50c8-128">Pour ouvrir Visual Studio Code en tant qu’administrateur, sélectionnez **l’option** Exécuter en tant qu’administrateur lors de l’ouverture Visual Studio Code après l’avoir recherché dans Windows.</span><span class="sxs-lookup"><span data-stu-id="f50c8-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

1. <span data-ttu-id="f50c8-129">Dans VS Code, sélectionnez **Ctrl + Maj + X** pour ouvrir la barre Extensions.</span><span class="sxs-lookup"><span data-stu-id="f50c8-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="f50c8-130">Recherchez l’extension « Microsoft Office débompeur de add-in » et installez-la.</span><span class="sxs-lookup"><span data-stu-id="f50c8-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

1. <span data-ttu-id="f50c8-131">Dans le dossier .vscode de votre projet, ouvrez le fichier **launch.json**.</span><span class="sxs-lookup"><span data-stu-id="f50c8-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="f50c8-132">Ajoutez le code suivant à la `configurations` section.</span><span class="sxs-lookup"><span data-stu-id="f50c8-132">Add the following code to the `configurations` section.</span></span>

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

1. <span data-ttu-id="f50c8-133">Dans la section JSON que vous avez copiée, recherchez la section « url ».</span><span class="sxs-lookup"><span data-stu-id="f50c8-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="f50c8-134">Dans cette URL, vous devez remplacer le texte HOST en minuscules par l’application qui héberge votre Office de messagerie.</span><span class="sxs-lookup"><span data-stu-id="f50c8-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="f50c8-135">Par exemple, si votre Office est pour Excel, la valeur de votre URL est « https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0 ».</span><span class="sxs-lookup"><span data-stu-id="f50c8-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

1. <span data-ttu-id="f50c8-136">Ouvrez l’invite de commandes et assurez-vous que vous êtes dans le dossier racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="f50c8-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="f50c8-137">Exécutez la commande `npm start` pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="f50c8-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="f50c8-138">Lorsque votre add-in se charge dans le client Office client, ouvrez le volet Des tâches.</span><span class="sxs-lookup"><span data-stu-id="f50c8-138">When your add-in loads in the Office client, open the task pane.</span></span>

1. <span data-ttu-id="f50c8-139">Revenir à Visual Studio Code et choisissez **Afficher >** déboguer ou entrez **Ctrl + Shift + D** pour basculer en mode débogage.</span><span class="sxs-lookup"><span data-stu-id="f50c8-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

1. <span data-ttu-id="f50c8-140">Dans les options de débogage, sélectionnez **Attacher aux Office de travail.** Sélectionnez **F5** ou **choisissez Déboguer -> démarrer le** débogage à partir du menu pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="f50c8-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

1. <span data-ttu-id="f50c8-141">Définissez un point d’arrêt dans le fichier du volet Des tâches de votre projet.</span><span class="sxs-lookup"><span data-stu-id="f50c8-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="f50c8-142">Vous pouvez définir des points d’arrêt Visual Studio Code en pointant à côté d’une ligne de code et en sélectionnant le cercle rouge qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="f50c8-142">You can set breakpoints in Visual Studio Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

    ![Un cercle rouge apparaît sur une ligne de code Visual Studio Code.](../images/set-breakpoint.jpg)

1. <span data-ttu-id="f50c8-144">Exécutez votre add-in.</span><span class="sxs-lookup"><span data-stu-id="f50c8-144">Run your add-in.</span></span> <span data-ttu-id="f50c8-145">Vous verrez que les points d’arrêt ont été atteints et que vous pouvez inspecter les variables locales.</span><span class="sxs-lookup"><span data-stu-id="f50c8-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="f50c8-146">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f50c8-146">See also</span></span>

- [<span data-ttu-id="f50c8-147">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="f50c8-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

- [<span data-ttu-id="f50c8-148">Débogage des compléments avec les outils de développement sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="f50c8-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [<span data-ttu-id="f50c8-149">Déboguer des compléments à l’aide de Microsoft Edge WebView2 (avec Chromium)</span><span class="sxs-lookup"><span data-stu-id="f50c8-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
