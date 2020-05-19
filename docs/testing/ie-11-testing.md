---
ms.date: 05/16/2020
description: Testez votre complément Office à l’aide d’Internet Explorer 11.
title: Test Internet Explorer 11
localization_priority: Normal
ms.openlocfilehash: 697c87d90df9aa70a7b20da5cd4c91d4445fb850
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275945"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a><span data-ttu-id="a9dc0-103">Tester votre complément Office à l’aide d’Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="a9dc0-103">Test your Office Add-in using Internet Explorer 11</span></span>

<span data-ttu-id="a9dc0-104">En fonction des spécifications de votre complément, vous pouvez envisager de prendre en charge des versions antérieures de Windows et d’Office, qui nécessitent des tests sur Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-104">Depending on the specifications of your add-in, you may plan to support older versions of Windows and Office, which require testing on Internet Explorer 11.</span></span> <span data-ttu-id="a9dc0-105">Cela est souvent nécessaire dans le cadre de l’envoi de votre complément à AppSource.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-105">This is often necessary as part of submitting your add-in to AppSource.</span></span> <span data-ttu-id="a9dc0-106">Vous pouvez utiliser les outils de ligne de commande suivants pour basculer d’autres runtimes modernes utilisés par les compléments vers le runtime Internet Explorer 11 pour ce test.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-106">You can use the following command line tooling to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span>

## <a name="pre-requisites"></a><span data-ttu-id="a9dc0-107">Conditions préalables</span><span class="sxs-lookup"><span data-stu-id="a9dc0-107">Pre-requisites</span></span>

- <span data-ttu-id="a9dc0-108">[Node.js](https://nodejs.org/) (la dernière version [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="a9dc0-108">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>
- <span data-ttu-id="a9dc0-109">Éditeur de code.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-109">A code editor.</span></span> <span data-ttu-id="a9dc0-110">Nous recommandons [Visual Studio code](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="a9dc0-110">We recommend [Visual Studio Code](https://code.visualstudio.com/)</span></span>
- [<span data-ttu-id="a9dc0-111">Faire partie du programme Office Insider</span><span class="sxs-lookup"><span data-stu-id="a9dc0-111">Be part of the Office Insider program</span></span>](https://insider.office.com)

<span data-ttu-id="a9dc0-112">Ces instructions supposent que vous avez configuré un projet de générateur Office Yo avant.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-112">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="a9dc0-113">Si vous n’avez pas encore fait cela, envisagez de lire un démarrage rapide, tel que celui- [ci pour les compléments Excel](../quickstarts/excel-quickstart-jquery.md).</span><span class="sxs-lookup"><span data-stu-id="a9dc0-113">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="using-ie11-tooling"></a><span data-ttu-id="a9dc0-114">Utilisation des outils IE11</span><span class="sxs-lookup"><span data-stu-id="a9dc0-114">Using IE11 tooling</span></span>

1. <span data-ttu-id="a9dc0-115">Créez un projet de générateur Office Yo.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-115">Create a Yo Office generator project.</span></span> <span data-ttu-id="a9dc0-116">Quel que soit le type de projet que vous sélectionnez, ces outils fonctionnent avec tous les types de projets.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-116">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

> <span data-ttu-id="a9dc0-117">! Note Si vous disposez d’un projet existant et que vous souhaitez ajouter cet outil sans créer de nouveau projet, ignorez cette étape et passez à l’étape suivante.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-117">![NOTE] If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

2. <span data-ttu-id="a9dc0-118">Dans le dossier racine de votre nouveau projet, exécutez la commande suivante dans la ligne de commande :</span><span class="sxs-lookup"><span data-stu-id="a9dc0-118">In the root folder of your new project, run the following in the command line:</span></span>

```command&nbsp;line
office-add-dev-settings webview manifest.xml ie
```
<span data-ttu-id="a9dc0-119">Vous devriez voir une remarque dans la ligne de commande que le type d’affichage Web est maintenant défini sur Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-119">You should see a note in the command line that the web view type is now set to IE.</span></span>

> <span data-ttu-id="a9dc0-120">! TETE Il n’est pas nécessaire d’utiliser cet outil, mais cela devrait vous aider à déboguer la majorité des problèmes liés à Internet Explorer 11 Runtime.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-120">![TIP] It isn't necessary to use this tooling, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="a9dc0-121">Pour une robustesse totale, vous devez tester à l’aide d’un ordinateur sur lequel une copie de Windows 7 et Office 2013 est installée.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-121">For complete robustness, you should test using a computer with a copy of Windows 7 and Office 2013 installed.</span></span>

## <a name="command-settings"></a><span data-ttu-id="a9dc0-122">Paramètres de la commande</span><span class="sxs-lookup"><span data-stu-id="a9dc0-122">Command settings</span></span>

<span data-ttu-id="a9dc0-123">Si vous avez un chemin d’accès de manifeste différent, spécifiez-le dans la commande, comme indiqué dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="a9dc0-123">Should you have a different manifest path, specify this in the command, as shown in the following:</span></span>

`office-add-dev-settings webview [path to your manifest] ie`

<span data-ttu-id="a9dc0-124">La `office-addin-dev-settings webview` commande peut également prendre un certain nombre d’exécutions en tant qu’arguments :</span><span class="sxs-lookup"><span data-stu-id="a9dc0-124">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="a9dc0-125">échange</span><span class="sxs-lookup"><span data-stu-id="a9dc0-125">ie</span></span>
- <span data-ttu-id="a9dc0-126">cadre</span><span class="sxs-lookup"><span data-stu-id="a9dc0-126">edge</span></span>
- <span data-ttu-id="a9dc0-127">Valeur par défaut.</span><span class="sxs-lookup"><span data-stu-id="a9dc0-127">default</span></span>

## <a name="see-also"></a><span data-ttu-id="a9dc0-128">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a9dc0-128">See also</span></span>
* [<span data-ttu-id="a9dc0-129">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="a9dc0-129">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="a9dc0-130">Chargement de la version test des compléments Office</span><span class="sxs-lookup"><span data-stu-id="a9dc0-130">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="a9dc0-131">Débogage des compléments avec les outils de développement sur Windows 10</span><span class="sxs-lookup"><span data-stu-id="a9dc0-131">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="a9dc0-132">Attacher un débogueur à partir du volet Office</span><span class="sxs-lookup"><span data-stu-id="a9dc0-132">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)