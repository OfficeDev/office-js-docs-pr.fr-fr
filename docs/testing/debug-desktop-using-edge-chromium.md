---
title: Déboguer des compléments à l’aide de Microsoft Edge WebView2 (avec Chromium)
description: Découvrez comment déboguer un complément Office qui utilise Microsoft Edge WebView2 (avec Chromium) à l’aide du débogueur pour l’extension Microsoft Edge dans VS Code.
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 0908bb5040b49568006324600acacb5e36dbd1a5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238113"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a><span data-ttu-id="aeea7-103">Déboguer un complément à l’aide de Microsoft Edge WebView2</span><span class="sxs-lookup"><span data-stu-id="aeea7-103">Debug add-ins on Windows using Edge Chromium WebView2</span></span>

<span data-ttu-id="aeea7-104">L’exécution d’un complément Office sur Windows peut utiliser le débogueur pour l’extension Microsoft Edge dans VS Code pour déboguer sur le runtime d’Edge Chromium WebView2.</span><span class="sxs-lookup"><span data-stu-id="aeea7-104">Office Add-ins running on Windows can use the Debugger for Microsoft Edge extension in VS Code to debug against the Edge Chromium WebView2 runtime.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="aeea7-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="aeea7-105">Prerequisites</span></span>

- <span data-ttu-id="aeea7-106">[Visual Studio Code](https://code.visualstudio.com/) (doit être exécuté en tant qu’administrateur)</span><span class="sxs-lookup"><span data-stu-id="aeea7-106">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="aeea7-107">Node.js (version 10+)</span><span class="sxs-lookup"><span data-stu-id="aeea7-107">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="aeea7-108">Windows 10</span><span class="sxs-lookup"><span data-stu-id="aeea7-108">Windows 10</span></span>
- [<span data-ttu-id="aeea7-109">Microsoft Edge Chromium à la disposition des participants au programme Insider de Windows</span><span class="sxs-lookup"><span data-stu-id="aeea7-109">Microsoft Edge Chromium available to Windows Insiders</span></span>](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="aeea7-110">Installer et utiliser le débogueur</span><span class="sxs-lookup"><span data-stu-id="aeea7-110">Install and use the debugger</span></span>

1. <span data-ttu-id="aeea7-111">Créez un projet à l’aide du [générateur Yoman pour complément Office](https://github.com/OfficeDev/generator-office). Vous pouvez utiliser l’un de nos guides de démarrage rapide, tels que le [Démarrage rapide du complément Outlook](../quickstarts/outlook-quickstart.md) pour pouvoir exécuter cette opération.</span><span class="sxs-lookup"><span data-stu-id="aeea7-111">Create a project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). You can use any one of our quick start guides, such as the [Outlook add-in quickstart](../quickstarts/outlook-quickstart.md), in order to do this.</span></span>

> [!TIP]
> <span data-ttu-id="aeea7-112">Si vous n’utilisez pas de générateur Yeoman basé sur un complément, vous devez régler une clé de Registre.</span><span class="sxs-lookup"><span data-stu-id="aeea7-112">If you aren't using a Yeoman generator based add-in, you need to adjust a registry key.</span></span> <span data-ttu-id="aeea7-113">Lorsque vous êtes dans le dossier racine de votre projet, exécutez ce qui suit dans la ligne de commande : `office-add-in-debugging start <your manifest path>`.</span><span class="sxs-lookup"><span data-stu-id="aeea7-113">While in the root folder of your project, run the following in the command line: `office-add-in-debugging start <your manifest path>`.</span></span>

2. <span data-ttu-id="aeea7-114">Ouvrez le projet dans VS Code.</span><span class="sxs-lookup"><span data-stu-id="aeea7-114">Open your project in VS Code.</span></span> <span data-ttu-id="aeea7-115">Dans VS Code, sélectionnez **Ctrl + Maj + X** pour ouvrir la barre Extensions.</span><span class="sxs-lookup"><span data-stu-id="aeea7-115">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="aeea7-116">Recherchez l’extension « Débogueur pour Microsoft Edge », puis installez-la.</span><span class="sxs-lookup"><span data-stu-id="aeea7-116">Search for the "Debugger for Microsoft Edge" extension and install it.</span></span>

3. <span data-ttu-id="aeea7-117">Dans le dossier **.vscode** de votre projet, ouvrez le fichier **launch.json**.</span><span class="sxs-lookup"><span data-stu-id="aeea7-117">In the **.vscode** folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="aeea7-118">Ajoutez le code suivant à la section de configuration :</span><span class="sxs-lookup"><span data-stu-id="aeea7-118">Add the following code to the configurations section:</span></span>

```JSON
  {
      "name": "Debug Office Add-in (Edge Chromium)",
      "type": "edge",
      "request": "attach",
      "useWebView": "advanced",
      "port": 9229,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
    },
```

4. <span data-ttu-id="aeea7-119">Ensuite, choisissez **Afficher > Débogage** ou entrez **Ctrl + Maj + D** pour passer à l’affichage Débogage.</span><span class="sxs-lookup"><span data-stu-id="aeea7-119">Next, choose  **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

5. <span data-ttu-id="aeea7-120">À partir des options Débogage, choisissez l’option Edge Chromium pour votre application hôte, telle que la **version de bureau d’Excel (Edge Chromium)**</span><span class="sxs-lookup"><span data-stu-id="aeea7-120">From the Debug options, choose the Edge Chromium option for your host application, such as **Excel Desktop (Edge Chromium)**.</span></span> <span data-ttu-id="aeea7-121">Sélectionnez **F5** ou choisissez **Déboguer > Démarrer le débogage** à partir du menu pour commencer le débogage.</span><span class="sxs-lookup"><span data-stu-id="aeea7-121">Select **F5** or choose **Debug > Start Debugging** from the menu to begin debugging.</span></span>

6. <span data-ttu-id="aeea7-122">Dans l’application hôte, telle qu’Excel, votre complément est désormais prêt à être utilisé.</span><span class="sxs-lookup"><span data-stu-id="aeea7-122">In the host application, such as Excel, your add-in is now ready to use.</span></span> <span data-ttu-id="aeea7-123">Sélectionnez **Afficher le volet de tâches** ou exécutez toute autre commande de complément.</span><span class="sxs-lookup"><span data-stu-id="aeea7-123">Select **Show Taskpane** or run any other add-in command.</span></span> <span data-ttu-id="aeea7-124">Une boîte de dialogue s'affiche, indiquant :</span><span class="sxs-lookup"><span data-stu-id="aeea7-124">A dialog box will appear, reading:</span></span>

> <span data-ttu-id="aeea7-125">Arrêter sur chargement WebView.</span><span class="sxs-lookup"><span data-stu-id="aeea7-125">WebView Stop On Load.</span></span> 
> <span data-ttu-id="aeea7-126">Pour déboguer l’affichage web, attachez VS Code dans l’instance d’affichage web à l’aide du débogueur Microsoft pour l’extension Edge, puis cliquez sur OK pour continuer.</span><span class="sxs-lookup"><span data-stu-id="aeea7-126">To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue.</span></span> <span data-ttu-id="aeea7-127">Pour empêcher l’affichage de cette boîte de dialogue dans le futur, cliquez sur « Annuler ».</span><span class="sxs-lookup"><span data-stu-id="aeea7-127">To prevent this dialog from appearing in the future, click Cancel."</span></span>

<span data-ttu-id="aeea7-128">Sélectionnez **OK**.</span><span class="sxs-lookup"><span data-stu-id="aeea7-128">Select **OK**.</span></span>

> [!NOTE]
> <span data-ttu-id="aeea7-129">Si vous sélectionnez **Annuler**, la boîte de dialogue ne s’affiche plus lors de l’exécution de cette instance du complément.</span><span class="sxs-lookup"><span data-stu-id="aeea7-129">If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running.</span></span> <span data-ttu-id="aeea7-130">Toutefois, si vous redémarrez votre complément, la boîte de dialogue s’affichera à nouveau.</span><span class="sxs-lookup"><span data-stu-id="aeea7-130">However, if you restart your add-in, you'll see the dialog again.</span></span>

7. <span data-ttu-id="aeea7-131">Vous pourrez définir des points d’arrêt dans le code de votre projet, puis déboguer.</span><span class="sxs-lookup"><span data-stu-id="aeea7-131">You're now able to set breakpoints in your project's code and debug.</span></span>

## <a name="see-also"></a><span data-ttu-id="aeea7-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="aeea7-132">See also</span></span>

* [<span data-ttu-id="aeea7-133">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="aeea7-133">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="aeea7-134">Complément Microsoft Office Extension de débogueur pour Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="aeea7-134">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
* [<span data-ttu-id="aeea7-135">Attacher un débogueur à partir du volet Office</span><span class="sxs-lookup"><span data-stu-id="aeea7-135">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)