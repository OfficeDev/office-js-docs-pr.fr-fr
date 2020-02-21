---
title: Explorer l’API JavaScript Office à l’aide de Script Lab
description: Utilisez Script Lab pour explorer l’API JS Office et pour prototyper les fonctionnalités.
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: e36aeaeb7313be2c5b797c7ce8f3b7d8a0dedc10
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163884"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="03a83-103">Explorer l’API JavaScript Office à l’aide de Script Lab</span><span class="sxs-lookup"><span data-stu-id="03a83-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="03a83-104">Le [complément Script Lab](https://appsource.microsoft.com/product/office/WA104380862), disponible gratuitement à partir d’AppSource, vous permet d’explorer l’API JavaScript Office lorsque vous travaillez dans un programme Office tel qu’Excel ou Word.</span><span class="sxs-lookup"><span data-stu-id="03a83-104">The [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862), which is available free from AppSource, enables you to explore the Office JavaScript API while you're working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="03a83-105">Script Lab est un outil pratique à ajouter à votre kit de ressources de développement lorsque vous prototypez et vérifiez les fonctionnalités souhaitées dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="03a83-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="03a83-106">Qu’est-ce que script Lab ?</span><span class="sxs-lookup"><span data-stu-id="03a83-106">What is Script Lab?</span></span>

<span data-ttu-id="03a83-107">Script Lab est un outil destiné à toute personne souhaitant en savoir plus sur la manière de développer des compléments Office à l’aide de l’API JavaScript Office dans Excel, Word ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="03a83-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="03a83-108">Il fournit IntelliSense, si bien que vous pouvez voir ce qui est disponible et qui repose sur l’infrastructure de Monaco, l’infrastructure utilisée par Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="03a83-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="03a83-109">Via Script Lab, vous pouvez accéder à une bibliothèque d'exemples pour essayer rapidement des fonctionnalités ou utiliser un exemple comme point de départ pour votre propre code.</span><span class="sxs-lookup"><span data-stu-id="03a83-109">Through Script Lab, you can access a library of samples to quickly try out features or you can use a sample as the starting point for your own code.</span></span> <span data-ttu-id="03a83-110">Vous pouvez même utiliser Script Lab pour essayer les API d’aperçu.</span><span class="sxs-lookup"><span data-stu-id="03a83-110">You can even use Script Lab to try preview APIs.</span></span>

<span data-ttu-id="03a83-111">C’est bien pour l’instant ?</span><span class="sxs-lookup"><span data-stu-id="03a83-111">Sounds good so far?</span></span> <span data-ttu-id="03a83-112">Visionnez cette vidéo d’une minute pour découvrir Script Lab en action.</span><span class="sxs-lookup"><span data-stu-id="03a83-112">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="03a83-113">[![Vidéo d’aperçu montrant l’exécution d’un Script Lab dans Excel, Word et PowerPoint.](../images/screenshot-wide-youtube.png 'Vidéo de la version préliminaire de Script Lab')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="03a83-113">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="key-features"></a><span data-ttu-id="03a83-114">Principales fonctionnalités</span><span class="sxs-lookup"><span data-stu-id="03a83-114">Key features</span></span>

<span data-ttu-id="03a83-115">Script Lab propose de nombreuses fonctionnalités pour vous aider à explorer l’API JavaScript Office et la fonctionnalité de complément prototype.</span><span class="sxs-lookup"><span data-stu-id="03a83-115">Script Lab offers a number of features to help you explore the Office JavaScript API and prototype add-in functionality.</span></span>

### <a name="explore-samples"></a><span data-ttu-id="03a83-116">Explorer des exemples</span><span class="sxs-lookup"><span data-stu-id="03a83-116">Explore samples</span></span>

<span data-ttu-id="03a83-117">Commencez rapidement avec une collection d’exemples d’extraits de code intégrés qui montrent comment effectuer des tâches avec l’API.</span><span class="sxs-lookup"><span data-stu-id="03a83-117">Get started quickly with a collection of built-in sample snippets that show how to complete tasks with the API.</span></span> <span data-ttu-id="03a83-118">Vous pouvez exécuter les exemples pour afficher instantanément le résultat dans le volet des tâches ou le document, examiner les exemples pour découvrir le fonctionnement de l’API, voire utiliser les exemples pour prototyper votre propre complément.</span><span class="sxs-lookup"><span data-stu-id="03a83-118">You can run the samples to instantly see the result in the task pane or document, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

![Exemples](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a><span data-ttu-id="03a83-120">Code et style</span><span class="sxs-lookup"><span data-stu-id="03a83-120">Code and style</span></span>

<span data-ttu-id="03a83-121">En plus du code JavaScript ou TypeScript qui appelle l’API Office JS, chaque extrait de code contient également une balise HTML qui définit le contenu du volet des tâches et CSS qui définit l’apparence de ce dernier.</span><span class="sxs-lookup"><span data-stu-id="03a83-121">In addition to JavaScript or TypeScript code that calls the Office JS API, each snippet also contains HTML markup that defines content of the task pane and CSS that defines the appearance of the task pane.</span></span> <span data-ttu-id="03a83-122">Vous pouvez personnaliser la balise HTML et CSS pour tester le placement des éléments et les styles lorsque vous prototypez la conception du volet des tâches pour votre propre complément.</span><span class="sxs-lookup"><span data-stu-id="03a83-122">You can customize the HTML markup and CSS to experiment with element placement and styling as you prototype task pane design for your own add-in.</span></span>

> [!TIP]
> <span data-ttu-id="03a83-123">Pour appeler les API d’aperçu dans un extrait de code, vous devez mettre à jour les bibliothèques de l’extrait de code de façon à utiliser le CDN bêta (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) et les définitions de type d’aperçu `@types/office-js-preview`.</span><span class="sxs-lookup"><span data-stu-id="03a83-123">To call preview APIs within a snippet, you'll need to update the snippet's libraries to use the beta CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) and the preview type definitions `@types/office-js-preview`.</span></span> <span data-ttu-id="03a83-124">De plus, certaines API d’aperçu sont accessibles uniquement si vous êtes inscrit au [programme Office Insider](https://products.office.com/office-insider) et que vous exécutez une version Insider d’Office.</span><span class="sxs-lookup"><span data-stu-id="03a83-124">Additionally, some preview APIs are only accessible if you've signed up for the [Office Insider program](https://products.office.com/office-insider) and are running an Insider build of Office.</span></span>

### <a name="save-and-share-snippets"></a><span data-ttu-id="03a83-125">Enregistrer et partager des extraits de code</span><span class="sxs-lookup"><span data-stu-id="03a83-125">Save and share snippets</span></span>

<span data-ttu-id="03a83-126">Par défaut, les extraits de code que vous ouvrez dans Script Lab sont enregistrés dans le cache de votre navigateur.</span><span class="sxs-lookup"><span data-stu-id="03a83-126">By default, snippets that you open in Script Lab will be saved to your browser cache.</span></span> <span data-ttu-id="03a83-127">Pour enregistrer définitivement un extrait de code, vous pouvez l’exporter dans un contenu [Gist GitHub](https://gist.github.com).</span><span class="sxs-lookup"><span data-stu-id="03a83-127">To save a snippet permanently, you can export it to a [GitHub gist](https://gist.github.com).</span></span> <span data-ttu-id="03a83-128">Créez un contenu Gist secret pour enregistrer un extrait de code exclusivement pour votre usage personnel ou créez un contenu Gist public si vous envisagez de le partager avec d’autres personnes.</span><span class="sxs-lookup"><span data-stu-id="03a83-128">Create a secret gist to save a snippet exclusively for your own use, or create a public gist if you plan to share it with others.</span></span>

![Options de partage](../images/script-lab-share.jpg)

### <a name="import-snippets"></a><span data-ttu-id="03a83-130">Importer des extraits de code</span><span class="sxs-lookup"><span data-stu-id="03a83-130">Import snippets</span></span>

<span data-ttu-id="03a83-131">Vous pouvez importer un extrait de code dans Script Lab en spécifiant l’URL du [contenu Gist GitHub](https://gist.github.com) public où le YAML de l’extrait de code est stocké ou en collant dans le YAML complet de l’extrait de code.</span><span class="sxs-lookup"><span data-stu-id="03a83-131">You can import a snippet into Script Lab either by specifying the URL to the public [GitHub gist](https://gist.github.com) where the snippet YAML is stored or by pasting in the complete YAML for the snippet.</span></span> <span data-ttu-id="03a83-132">Cette fonctionnalité peut être utile dans les cas où quelqu’un d’autre a partagé son extrait de code avec vous, soit en le publiant dans un contenu Gist GitHub, soit en fournissant le YAML de son extrait de code.</span><span class="sxs-lookup"><span data-stu-id="03a83-132">This feature may be useful in scenarios where someone else has shared their snippet with you by either publishing it to a GitHub gist or providing their snippet's YAML.</span></span>

![Option Importer un extrait](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a><span data-ttu-id="03a83-134">Clients pris en charge</span><span class="sxs-lookup"><span data-stu-id="03a83-134">Supported clients</span></span>

<span data-ttu-id="03a83-135">Script Lab est pris en charge pour Excel, Word et PowerPoint sur les clients suivants.</span><span class="sxs-lookup"><span data-stu-id="03a83-135">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="03a83-136">Office 2013 ou version ultérieure sous Windows</span><span class="sxs-lookup"><span data-stu-id="03a83-136">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="03a83-137">Office 2016 ou version ultérieure sous Mac</span><span class="sxs-lookup"><span data-stu-id="03a83-137">Office 2016 or later on Mac</span></span>
- <span data-ttu-id="03a83-138">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="03a83-138">Office on the web</span></span>

## <a name="next-steps"></a><span data-ttu-id="03a83-139">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="03a83-139">Next steps</span></span>

<span data-ttu-id="03a83-140">Pour utiliser Script Lab dans Excel, Word ou PowerPoint, installez le [complément Script Lab](https://appsource.microsoft.com/product/office/WA104380862) à partir d’AppSource.</span><span class="sxs-lookup"><span data-stu-id="03a83-140">To use Script Lab in Excel, Word, or PowerPoint, install the [Script Lab add-in](https://appsource.microsoft.com/product/office/WA104380862) from AppSource.</span></span> 

<span data-ttu-id="03a83-141">Nous vous invitons à développer l’exemple de bibliothèque dans Script Lab en apportant de nouveaux extraits de code dans le référentiel GitHub [Office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span><span class="sxs-lookup"><span data-stu-id="03a83-141">You're welcome to expand the sample library in Script Lab by contributing new snippets to the [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub repository.</span></span>

<span data-ttu-id="03a83-142">Lorsque vous êtes prêt à créer votre premier complément Office, essayez le guide de démarrage rapide pour [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md) ou [Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="03a83-142">When you're ready to create your first Office Add-in, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](../quickstarts/outlook-quickstart.md), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="03a83-143">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="03a83-143">See also</span></span>

- [<span data-ttu-id="03a83-144">Obtenir Script Lab</span><span class="sxs-lookup"><span data-stu-id="03a83-144">Get Script Lab</span></span>](https://appsource.microsoft.com/product/office/WA104380862)
- [<span data-ttu-id="03a83-145">En savoir plus sur Script Lab</span><span class="sxs-lookup"><span data-stu-id="03a83-145">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="03a83-146">Rejoindre le programme pour les développeurs Office 365</span><span class="sxs-lookup"><span data-stu-id="03a83-146">Join the Office 365 Developer Program</span></span>](https://developer.microsoft.com/office/dev-program)
- [<span data-ttu-id="03a83-147">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="03a83-147">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
