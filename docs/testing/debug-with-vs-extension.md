---
title: Complément Microsoft Office Extension de débogueur pour Visual Studio Code
description: Utilisez l’extension Visual Studio code Microsoft Office déboguer votre module de déboguer votre add-in Office.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237930"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="f997a-103">Complément Microsoft Office Extension de débogueur pour Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f997a-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="f997a-104">L’extension Microsoft Office déboguer de l’application pour Visual Studio Code vous permet de déboguer votre application Office par rapport à Microsoft Edge avec le runtime WebView d’origine (EdgeHTML).</span><span class="sxs-lookup"><span data-stu-id="f997a-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="f997a-105">Pour obtenir des instructions sur le débogage sur Microsoft Edge WebView2 (basé sur Chromium), [consultez cet article.](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="f997a-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="f997a-106">Ce mode de débogage est dynamique, ce qui vous permet de définir des points d’arrêt pendant l’exécution du code.</span><span class="sxs-lookup"><span data-stu-id="f997a-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="f997a-107">Vous pouvez voir les modifications dans votre code immédiatement lorsque le déboguer est attaché, tout cela sans perdre votre session de débogage.</span><span class="sxs-lookup"><span data-stu-id="f997a-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="f997a-108">Vos modifications de code sont également persistantes, afin que vous pouvez voir les résultats de plusieurs modifications apportées à votre code.</span><span class="sxs-lookup"><span data-stu-id="f997a-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="f997a-109">L’image suivante illustre cette extension en action.</span><span class="sxs-lookup"><span data-stu-id="f997a-109">The following image shows this extension in action.</span></span>

![Extension de déboguer du débogage d’une section de modules de débogage de l’extension de débogage de l’extension de débogage d’un addin Office](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="f997a-111">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f997a-111">Prerequisites</span></span>

- <span data-ttu-id="f997a-112">[Visual Studio code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)</span><span class="sxs-lookup"><span data-stu-id="f997a-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="f997a-113">Node.js (version 10+)</span><span class="sxs-lookup"><span data-stu-id="f997a-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="f997a-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="f997a-114">Windows 10</span></span>
- [<span data-ttu-id="f997a-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="f997a-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="f997a-116">Ces instructions supposent que vous avez de l’expérience en utilisant la ligne de commande, que vous comprenez javaScript de base et que vous avez créé un projet de add-in Office avant d’utiliser le générateur Yo Office.</span><span class="sxs-lookup"><span data-stu-id="f997a-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="f997a-117">Si vous ne l’avez pas encore fait, envisagez de consulter l’un de nos didacticiels, comme ce didacticiel sur les [modules de 2013 excel.](../tutorials/excel-tutorial.md)</span><span class="sxs-lookup"><span data-stu-id="f997a-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="f997a-118">Installer et utiliser le débogger</span><span class="sxs-lookup"><span data-stu-id="f997a-118">Install and use the debugger</span></span>

1. <span data-ttu-id="f997a-119">Si vous devez créer un projet de add-in, [utilisez le générateur Yo Office pour en créer un.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)</span><span class="sxs-lookup"><span data-stu-id="f997a-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="f997a-120">Suivez les invites de la ligne de commande pour configurer votre projet.</span><span class="sxs-lookup"><span data-stu-id="f997a-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="f997a-121">Vous pouvez choisir n’importe quelle langue ou type de projet en fonction de vos besoins.</span><span class="sxs-lookup"><span data-stu-id="f997a-121">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="f997a-122">Si vous avez déjà un projet, ignorez l’étape 1 et passez à l’étape 2.</span><span class="sxs-lookup"><span data-stu-id="f997a-122">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="f997a-123">Ouvrez une invite de commandes en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="f997a-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="f997a-124">![Options d’invite de commandes, y compris « Exécuter en tant qu’administrateur » dans Windows 10](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="f997a-124">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="f997a-125">Accédez au répertoire de votre projet.</span><span class="sxs-lookup"><span data-stu-id="f997a-125">Navigate to your project directory.</span></span>

4. <span data-ttu-id="f997a-126">Exécutez la commande suivante pour ouvrir votre projet dans Visual Studio Code en tant qu’administrateur.</span><span class="sxs-lookup"><span data-stu-id="f997a-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="f997a-127">Une Visual Studio code est ouvert, accédez manuellement au dossier du projet.</span><span class="sxs-lookup"><span data-stu-id="f997a-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="f997a-128">Pour ouvrir Visual Studio code en tant qu’administrateur, sélectionnez **l’option** Exécuter en tant qu’administrateur lors de l’ouverture Visual Studio Code après l’avoir recherché dans Windows.</span><span class="sxs-lookup"><span data-stu-id="f997a-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="f997a-129">Dans VS Code, sélectionnez **Ctrl + Shift + X** pour ouvrir la barre Extensions.</span><span class="sxs-lookup"><span data-stu-id="f997a-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="f997a-130">Recherchez l’extension « Microsoft Office débompeur de l’extension de module de 2013</span><span class="sxs-lookup"><span data-stu-id="f997a-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="f997a-131">Dans le dossier .vscode de votre projet, ouvrez le **fichierlaunch.jssur.**</span><span class="sxs-lookup"><span data-stu-id="f997a-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="f997a-132">Ajoutez le code suivant à la `configurations` section :</span><span class="sxs-lookup"><span data-stu-id="f997a-132">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="f997a-133">Dans la section JSON que vous avez copiée, recherchez la section « url ».</span><span class="sxs-lookup"><span data-stu-id="f997a-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="f997a-134">Dans cette URL, vous devez remplacer le texte HOST en minuscules par l’application qui héberge votre application Office.</span><span class="sxs-lookup"><span data-stu-id="f997a-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="f997a-135">Par exemple, si votre add-in Office est pour Excel, la valeur de votre URL sera « https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0 ».</span><span class="sxs-lookup"><span data-stu-id="f997a-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="f997a-136">Ouvrez l’invite de commandes et assurez-vous que vous êtes dans le dossier racine de votre projet.</span><span class="sxs-lookup"><span data-stu-id="f997a-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="f997a-137">Exécutez la commande `npm start` pour démarrer le serveur dev.</span><span class="sxs-lookup"><span data-stu-id="f997a-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="f997a-138">Lorsque votre add-in se charge dans le client Office, ouvrez le volet Office.</span><span class="sxs-lookup"><span data-stu-id="f997a-138">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="f997a-139">Revenir à Visual Studio Code et choisissez Afficher **>** Déboguer ou entrez **Ctrl + Shift + D** pour basculer en mode débogage.</span><span class="sxs-lookup"><span data-stu-id="f997a-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="f997a-140">Dans les options de débogage, choisissez **Attacher aux add-ins Office.** Sélectionnez **F5** ou **choisissez Déboguer -> démarrer le** débogage à partir du menu pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="f997a-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="f997a-141">Définissez un point d’arrêt dans le fichier du volet Des tâches de votre projet.</span><span class="sxs-lookup"><span data-stu-id="f997a-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="f997a-142">Vous pouvez définir des points d’arrêt dans VS Code en pointant sur une ligne de code et en sélectionnant le cercle rouge qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="f997a-142">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![Un cercle rouge apparaît sur une ligne de code dans VS Code](../images/set-breakpoint.jpg)

12. <span data-ttu-id="f997a-144">Exécutez votre add-in.</span><span class="sxs-lookup"><span data-stu-id="f997a-144">Run your add-in.</span></span> <span data-ttu-id="f997a-145">Vous verrez que les points d’arrêt ont été atteints et que vous pouvez inspecter les variables locales.</span><span class="sxs-lookup"><span data-stu-id="f997a-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="f997a-146">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f997a-146">See also</span></span>

* [<span data-ttu-id="f997a-147">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="f997a-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="f997a-148">Débogage des compléments avec les outils de développement sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="f997a-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="f997a-149">Déboguer des applications sur Windows à l’aide de Microsoft Edge WebView2 (basé sur Chromium)</span><span class="sxs-lookup"><span data-stu-id="f997a-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
