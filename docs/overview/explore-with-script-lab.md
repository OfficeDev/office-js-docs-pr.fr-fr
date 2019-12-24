---
title: Explorer l’API JavaScript Office à l’aide de Script Lab
description: Utilisez script Lab pour explorer l’API Office JS et pour prototyper les fonctionnalités.
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Normal
ms.openlocfilehash: fbefd205ac929579cea1120b8398a53146bca19c
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851515"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="c481d-103">Explorer l’API JavaScript Office à l’aide de Script Lab</span><span class="sxs-lookup"><span data-stu-id="c481d-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="c481d-104">Le [complément script Lab](https://appsource.microsoft.com/product/office/WA104380862), qui est disponible gratuitement à partir de AppSource, vous permet d’explorer l’API JavaScript Office pendant que vous travaillez dans un programme Office tel qu’Excel ou Word.</span><span class="sxs-lookup"><span data-stu-id="c481d-104">The [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862), which is available free from AppSource, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="c481d-105">Script Lab est un outil pratique à ajouter à votre boîte à outils de développement lorsque vous prototypez et vérifiez les fonctionnalités souhaitées dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="c481d-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="c481d-106">Qu’est-ce que script Lab ?</span><span class="sxs-lookup"><span data-stu-id="c481d-106">What is Script Lab?</span></span>

<span data-ttu-id="c481d-107">Script Lab est un outil destiné aux utilisateurs qui souhaitent apprendre à développer des compléments Office à l’aide de l’API JavaScript Office dans Excel, Word ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="c481d-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="c481d-108">Il fournit IntelliSense afin que vous puissiez voir ce qui est disponible et repose sur l’infrastructure Monaco, la même infrastructure utilisée par Visual Studio code.</span><span class="sxs-lookup"><span data-stu-id="c481d-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="c481d-109">Grâce à script Lab, vous pouvez accéder à une bibliothèque d’exemples pour essayer rapidement des fonctionnalités ou vous pouvez utiliser un exemple comme point de départ pour votre propre code.</span><span class="sxs-lookup"><span data-stu-id="c481d-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="c481d-110">Vous pouvez même utiliser l’atelier de script pour essayer les API d’aperçu.</span><span class="sxs-lookup"><span data-stu-id="c481d-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="c481d-111">Le bruit est-il bien fait ?</span><span class="sxs-lookup"><span data-stu-id="c481d-111">Sounds good so far?</span></span> <span data-ttu-id="c481d-112">Jetez un œil à cette vidéo d’une minute pour voir script Lab en action.</span><span class="sxs-lookup"><span data-stu-id="c481d-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="c481d-113">[![Aperçu de la vidéo avec script Lab en cours d’exécution dans Excel, Word et PowerPoint.](../images/screenshot-wide-youtube.png 'Vidéo de l’aperçu de script Lab')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="c481d-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="c481d-114">Principales fonctionnalités</span><span class="sxs-lookup"><span data-stu-id="c481d-114">Key features</span></span>

<span data-ttu-id="c481d-115">Script Lab offre un certain nombre de fonctionnalités pour vous aider à explorer l’API JavaScript Office et la fonctionnalité de complément prototype.</span><span class="sxs-lookup"><span data-stu-id="c481d-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="c481d-116">Explorer les exemples</span><span class="sxs-lookup"><span data-stu-id="c481d-116">Explore samples</span></span>

<span data-ttu-id="c481d-117">Prise en main rapide avec une collection d’extraits de code intégrés qui montrent comment effectuer des tâches avec l’API.</span><span class="sxs-lookup"><span data-stu-id="c481d-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="c481d-118">Vous pouvez exécuter les exemples pour voir instantanément le résultat dans le volet Office ou le document, examiner les exemples pour savoir comment fonctionne l’API, et même utiliser des exemples pour prototyper votre propre complément.</span><span class="sxs-lookup"><span data-stu-id="c481d-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![Exemples](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="c481d-120">Code et style</span><span class="sxs-lookup"><span data-stu-id="c481d-120">Code and style</span></span>

<span data-ttu-id="c481d-121">En plus du code JavaScript ou de la machine à écrire qui appelle l’API Office JS, chaque extrait de code contient également un balisage HTML qui définit le contenu du volet de tâches et CSS qui définit l’apparence du volet Office.</span><span class="sxs-lookup"><span data-stu-id="c481d-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="c481d-122">Vous pouvez personnaliser les balises HTML et CSS pour tester le positionnement et le style des éléments lorsque vous prototypez la conception de volet des tâches pour votre propre complément.</span><span class="sxs-lookup"><span data-stu-id="c481d-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="c481d-123">Pour appeler les API d’aperçu dans un extrait de code, vous devez mettre à jour les bibliothèques de l’extrait de`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`code afin d’utiliser la version `@types/office-js-preview`bêta de CDN () et les définitions des types d’aperçu.</span><span class="sxs-lookup"><span data-stu-id="c481d-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="c481d-124">En outre, certaines API d’aperçu ne sont accessibles que si vous vous êtes inscrit au [programme Office Insider](https://products.office.com/office-insider) et si vous exécutez une version Insider d’Office.</span><span class="sxs-lookup"><span data-stu-id="c481d-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://products.office.com/office-insider) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="c481d-125">Enregistrer et partager des extraits de code</span><span class="sxs-lookup"><span data-stu-id="c481d-125">Save and share snippets</span></span>

<span data-ttu-id="c481d-126">Par défaut, les extraits de code que vous ouvrez dans script Lab seront enregistrés dans le cache de votre navigateur.</span><span class="sxs-lookup"><span data-stu-id="c481d-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="c481d-127">Pour enregistrer un extrait de code de manière permanente, vous pouvez l’exporter vers un [GitHub](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="c481d-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="c481d-128">Créez un annuaire secret pour enregistrer un extrait de code exclusivement pour votre propre usage, ou créez un annuaire public si vous envisagez de le partager avec d’autres personnes.</span><span class="sxs-lookup"><span data-stu-id="c481d-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![Options de partage](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="c481d-130">Importer des extraits de code</span><span class="sxs-lookup"><span data-stu-id="c481d-130">Import snippets</span></span>

<span data-ttu-id="c481d-131">Vous pouvez importer un extrait de code dans script Lab en spécifiant l’URL du [GitHub](https://gist.github.com) public où l’extrait de code YAML est stocké ou en collant dans le YAML complet pour l’extrait de code.</span><span class="sxs-lookup"><span data-stu-id="c481d-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="c481d-132">Cette fonctionnalité peut être utile dans les scénarios où quelqu’un d’autre a partagé son extrait de code avec vous en le publiant dans un GitHub ou en fournissant les YAML de son extrait de code.</span><span class="sxs-lookup"><span data-stu-id="c481d-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![Option Importer un extrait](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="c481d-134">Clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="c481d-134">Supported clients</span></span>

<span data-ttu-id="c481d-135">Le script Lab est pris en charge pour Excel, Word et PowerPoint sur les clients suivants.</span><span class="sxs-lookup"><span data-stu-id="c481d-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="c481d-136">Office 2013 ou version ultérieure sur Windows</span><span class="sxs-lookup"><span data-stu-id="c481d-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="c481d-137">Office 2016 ou version ultérieure sur Mac</span><span class="sxs-lookup"><span data-stu-id="c481d-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="c481d-138">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="c481d-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="c481d-139">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="c481d-139">Next steps</span></span>

<span data-ttu-id="c481d-140">Pour utiliser script Lab dans Excel, Word ou PowerPoint, installez le [complément script Lab](https://appsource.microsoft.com/product/office/WA104380862) à partir de AppSource.</span><span class="sxs-lookup"><span data-stu-id="c481d-140">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="c481d-141">Vous pouvez développer l’exemple de bibliothèque dans script Lab en apposant de nouveaux extraits de code dans le référentiel GitHub [Office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) .</span><span class="sxs-lookup"><span data-stu-id="c481d-141">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="c481d-142">Lorsque vous êtes prêt à créer votre premier complément Office, essayez le démarrage rapide pour [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md)ou [Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="c481d-142">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c481d-143">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c481d-143">See also</span></span>

- [<span data-ttu-id="c481d-144">Obtenir un laboratoire de script</span><span class="sxs-lookup"><span data-stu-id="c481d-144">Get Script Lab</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="c481d-145">En savoir plus sur script Lab</span><span class="sxs-lookup"><span data-stu-id="c481d-145">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="c481d-146">S’inscrire au programme de développement</span><span class="sxs-lookup"><span data-stu-id="c481d-146">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="c481d-147">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="c481d-147">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
