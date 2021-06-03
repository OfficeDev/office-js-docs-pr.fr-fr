---
title: Test d’Internet Explorer 11
description: Testez votre Office sur Internet Explorer 11.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: de256ee8b0633f18d3188c5bbfae52cb24ff2c35
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727933"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a><span data-ttu-id="2a563-103">Tester votre Office sur Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="2a563-103">Test your Office Add-in on Internet Explorer 11</span></span>

<span data-ttu-id="2a563-104">Si vous envisagez de commercialiser votre application via AppSource ou si vous prévoyez de prendre en charge des versions antérieures de Windows et Office, votre application doit fonctionner dans le contrôle de navigateur in incorporer basé sur Internet Explorer 11 (IE11).</span><span class="sxs-lookup"><span data-stu-id="2a563-104">If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11).</span></span> <span data-ttu-id="2a563-105">Vous pouvez utiliser une ligne de commande pour passer de runtimes plus modernes utilisés par les modules de mise à l’essai à Internet Explorer 11 pour ce test.</span><span class="sxs-lookup"><span data-stu-id="2a563-105">You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span> <span data-ttu-id="2a563-106">Pour plus d’informations sur les versions de Windows et Office utiliser le contrôle d’affichage web Internet Explorer 11, voir Navigateurs utilisés par les Office des [applications.](../concepts/browsers-used-by-office-web-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="2a563-106">For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2a563-107">Internet Explorer 11 ne prend pas en charge les versions de JavaScript ultérieures à la version ES5.</span><span class="sxs-lookup"><span data-stu-id="2a563-107">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="2a563-108">Si vous souhaitez utiliser la syntaxe et les fonctionnalités d’ECMAScript 2015 ou une ultérieure, vous disposez de deux options :</span><span class="sxs-lookup"><span data-stu-id="2a563-108">If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:</span></span>
>
> - <span data-ttu-id="2a563-109">Écrivez votre code dans ECMAScript 2015 (également appelé ES6) ou version ultérieure JavaScript, ou dans TypeScript, puis compilez votre code en JavaScript ES5 à l’aide d’un compilateur tel que [celui-ci ou](https://babeljs.io/) [tsc.](https://www.typescriptlang.org/index.html)</span><span class="sxs-lookup"><span data-stu-id="2a563-109">Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).</span></span>
> - <span data-ttu-id="2a563-110">Écrivez en JavaScript ECMAScript 2015 ou version ultérieure, mais chargez également une [bibliothèque polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) telle que [core-js](https://github.com/zloirock/core-js) qui permet à IE d’exécuter votre code.</span><span class="sxs-lookup"><span data-stu-id="2a563-110">Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.</span></span>
>
> <span data-ttu-id="2a563-111">Pour plus d’informations sur ces options, voir [Support Internet Explorer 11](../develop/support-ie-11.md).</span><span class="sxs-lookup"><span data-stu-id="2a563-111">For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).</span></span>
>
> <span data-ttu-id="2a563-112">Par ailleurs, Internet Explorer 11 ne prend pas en charge certaines fonctionnalités HTML5 telles que les éléments multimédias, l’enregistrement et l’emplacement.</span><span class="sxs-lookup"><span data-stu-id="2a563-112">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="2a563-113">Pour tester votre add-in sur le navigateur Internet Explorer 11, ouvrez Office sur le Web dans Internet Explorer et chargez une version test [du module.](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="2a563-113">To test your add-in on the Internet Explorer 11 browser, open Office on the web in Internet Explorer and [sideload the add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2a563-114">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="2a563-114">Prerequisites</span></span>

- <span data-ttu-id="2a563-115">[Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="2a563-115">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

<span data-ttu-id="2a563-116">Ces instructions supposent que vous avez déjà installé un projet Office Yo.</span><span class="sxs-lookup"><span data-stu-id="2a563-116">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="2a563-117">Si vous ne l’avez pas encore fait, envisagez de lire un démarrage rapide, tel que [celui-ci pour Excel de recherche.](../quickstarts/excel-quickstart-jquery.md)</span><span class="sxs-lookup"><span data-stu-id="2a563-117">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="switching-to-the-internet-explorer-11-webview"></a><span data-ttu-id="2a563-118">Basculement vers le webview Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="2a563-118">Switching to the Internet Explorer 11 webview</span></span>

1. <span data-ttu-id="2a563-119">Créez un projet de générateur Office Yo.</span><span class="sxs-lookup"><span data-stu-id="2a563-119">Create a Yo Office generator project.</span></span> <span data-ttu-id="2a563-120">Peu importe le type de projet que vous sélectionnez, cet outil fonctionne avec tous les types de projets.</span><span class="sxs-lookup"><span data-stu-id="2a563-120">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

    > [!NOTE]
    > <span data-ttu-id="2a563-121">Si vous avez un projet existant et que vous souhaitez ajouter cet outil sans créer de nouveau projet, ignorez cette étape et passez à l’étape suivante.</span><span class="sxs-lookup"><span data-stu-id="2a563-121">If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

1. <span data-ttu-id="2a563-122">Dans le dossier racine de votre projet, exécutez la commande suivante dans la ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="2a563-122">In the root folder of your project, run the following in the command line.</span></span> <span data-ttu-id="2a563-123">Cet exemple suppose que le fichier manifeste de votre projet se trouve à la racine.</span><span class="sxs-lookup"><span data-stu-id="2a563-123">This example assumes that your project's manifest file is in the root.</span></span> <span data-ttu-id="2a563-124">Si ce n’est pas le cas, spécifiez le chemin d’accès relatif au fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="2a563-124">If it isn't, specify the relative path to the manifest file.</span></span> <span data-ttu-id="2a563-125">Un message doit s’afficher dans la ligne de commande pour vous dire que le type d’affichage web est désormais définie sur IE.</span><span class="sxs-lookup"><span data-stu-id="2a563-125">You should see a message in the command line that the web view type is now set to IE.</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> <span data-ttu-id="2a563-126">Il n’est pas nécessaire d’utiliser cette commande, mais elle doit aider à déboguer la plupart des problèmes liés au runtime d’Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="2a563-126">It isn't necessary to use this command, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="2a563-127">Pour une robustesse totale, vous devez tester l’utilisation d’ordinateurs avec différentes combinaisons de Windows 7, 8.1 et 10, ainsi que différentes versions de Office.</span><span class="sxs-lookup"><span data-stu-id="2a563-127">For complete robustness, you should test using computers with various combinations of Windows 7, 8.1, and 10 and various versions of Office.</span></span> <span data-ttu-id="2a563-128">Pour plus d’informations, voir [Navigateurs](../concepts/browsers-used-by-office-web-add-ins.md) utilisés par les Office et comment revenir à une version antérieure de [Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span><span class="sxs-lookup"><span data-stu-id="2a563-128">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span></span>

### <a name="command-options"></a><span data-ttu-id="2a563-129">Options de commande</span><span class="sxs-lookup"><span data-stu-id="2a563-129">Command options</span></span>

<span data-ttu-id="2a563-130">La `office-addin-dev-settings webview` commande peut également prendre un certain nombre d’runtimes comme arguments :</span><span class="sxs-lookup"><span data-stu-id="2a563-130">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="2a563-131">ie</span><span class="sxs-lookup"><span data-stu-id="2a563-131">ie</span></span>
- <span data-ttu-id="2a563-132">edge</span><span class="sxs-lookup"><span data-stu-id="2a563-132">edge</span></span>
- <span data-ttu-id="2a563-133">Valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="2a563-133">default</span></span>

## <a name="see-also"></a><span data-ttu-id="2a563-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="2a563-134">See also</span></span>

* [<span data-ttu-id="2a563-135">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="2a563-135">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="2a563-136">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="2a563-136">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="2a563-137">Débogage des compléments avec les outils de développement sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="2a563-137">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="2a563-138">Attacher un débogueur à partir du volet Office</span><span class="sxs-lookup"><span data-stu-id="2a563-138">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
