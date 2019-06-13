---
title: Explorer l’API JavaScript pour Office à l’aide de script Lab
description: Utilisez script Lab pour explorer l’API Office JS et pour prototyper les fonctionnalités.
ms.topic: article
ms.date: 06/07/2019
localization_priority: Normal
ms.openlocfilehash: 0bab566b08ba25dd3c01cff72f331b2dc9ce304d
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910189"
---
# <a name="explore-office-javascript-api-using-script-lab"></a><span data-ttu-id="780cc-103">Explorer l’API JavaScript pour Office à l’aide de script Lab</span><span class="sxs-lookup"><span data-stu-id="780cc-103">Explore Office JavaScript API using Script Lab</span></span>

<span data-ttu-id="780cc-104">Le [complément script Lab](https://store.office.com/app.aspx?assetid=WA104380862), qui est disponible gratuitement à partir de l’Office Store, vous permet d’explorer l’API JavaScript Office pendant que vous travaillez dans un programme Office tel qu’Excel ou Word.</span><span class="sxs-lookup"><span data-stu-id="780cc-104">The [Script Lab add-in](https://store.office.com/app.aspx?assetid=WA104380862), which is available free from the Office store, enables you to explore the Office JavaScript API while you are working in an Office program such as Excel or Word.</span></span> <span data-ttu-id="780cc-105">Script Lab est un outil pratique à ajouter à votre boîte à outils de développement lorsque vous prototypez et vérifiez les fonctionnalités souhaitées dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="780cc-105">Script Lab is a convenient tool to add to your development toolkit as you prototype and verify functionality you want in your add-in.</span></span>

## <a name="what-is-script-lab"></a><span data-ttu-id="780cc-106">Qu’est-ce que script Lab?</span><span class="sxs-lookup"><span data-stu-id="780cc-106">What is Script Lab?</span></span>

<span data-ttu-id="780cc-107">Script Lab est un outil destiné aux utilisateurs qui souhaitent apprendre à développer des compléments Office à l’aide de l’API JavaScript Office dans Excel, Word ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="780cc-107">Script Lab is a tool for anyone who wants to learn how to develop Office Add-ins using the Office JavaScript API in Excel, Word, or PowerPoint.</span></span> <span data-ttu-id="780cc-108">Il fournit IntelliSense afin que vous puissiez voir ce qui est disponible et repose sur l’infrastructure Monaco, la même infrastructure utilisée par Visual Studio code.</span><span class="sxs-lookup"><span data-stu-id="780cc-108">It provides IntelliSense so you can see what's available and is built on the Monaco framework, the same framework used by Visual Studio Code.</span></span> <span data-ttu-id="780cc-109">Grâce à script Lab, vous pouvez accéder à une bibliothèque d’exemples pour essayer rapidement des fonctionnalités ou vous pouvez choisir un exemple comme base pour votre propre code.</span><span class="sxs-lookup"><span data-stu-id="780cc-109">Through Script Lab, you can access a library of samples to quickly try out features or you can choose a sample as the base for your own code.</span></span> <span data-ttu-id="780cc-110">Vous pouvez également développer l’exemple de bibliothèque en ajoutant des extraits de code dans la [référentiel Office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span><span class="sxs-lookup"><span data-stu-id="780cc-110">You are also welcome to expand the sample library by adding snippets to the [office-js-snippets repo](https://github.com/OfficeDev/office-js-snippets#office-js-snippets).</span></span> <span data-ttu-id="780cc-111">Une autre fonctionnalité intéressante de script Lab est une fonctionnalité bêta ou d’aperçu que vous pouvez essayer.</span><span class="sxs-lookup"><span data-stu-id="780cc-111">Another exciting feature of Script Lab is beta or preview functionality is available for you to try.</span></span>

> [!TIP]
> <span data-ttu-id="780cc-112">Pour participer à la version bêta ou à l’aperçu, vous devrez peut-être vous inscrire au [programme Office](https://products.office.com/office-insider)Insider.</span><span class="sxs-lookup"><span data-stu-id="780cc-112">To participate in beta or preview, you may have to sign up for the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="780cc-113">Le bruit est-il bien fait?</span><span class="sxs-lookup"><span data-stu-id="780cc-113">Sounds good so far?</span></span> <span data-ttu-id="780cc-114">Jetez un œil à cette vidéo d’une minute pour voir script Lab en action.</span><span class="sxs-lookup"><span data-stu-id="780cc-114">Take a look at this one-minute video to see Script Lab in action.</span></span>

<span data-ttu-id="780cc-115">[![Aperçu de la vidéo avec script Lab en cours d’exécution dans Excel, Word et PowerPoint online.] (../images/screenshot-wide-youtube.png 'Vidéo de l’aperçu de script Lab')](https://aka.ms/scriptlabvideo)</span><span class="sxs-lookup"><span data-stu-id="780cc-115">[![Preview video showing Script Lab running in Excel, Word, and PowerPoint Online.](../images/screenshot-wide-youtube.png 'Script Lab preview video')](https://aka.ms/scriptlabvideo)</span></span>

## <a name="script-lab-supported-clients"></a><span data-ttu-id="780cc-116">Clients de script Lab pris en charge</span><span class="sxs-lookup"><span data-stu-id="780cc-116">Script Lab supported clients</span></span>

<span data-ttu-id="780cc-117">Le script Lab est pris en charge pour Excel, Word et PowerPoint sur les clients suivants.</span><span class="sxs-lookup"><span data-stu-id="780cc-117">Script Lab is supported for Excel, Word, and PowerPoint on the following clients.</span></span>

- <span data-ttu-id="780cc-118">Office sur Windows (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="780cc-118">Office on Windows (connected to Office 365)</span></span>
- <span data-ttu-id="780cc-119">Office pour Mac (connecté à Office 365)</span><span class="sxs-lookup"><span data-stu-id="780cc-119">Office for Mac (connected to Office 365)</span></span>
- <span data-ttu-id="780cc-120">Office Online</span><span class="sxs-lookup"><span data-stu-id="780cc-120">Office Online</span></span>
- <span data-ttu-id="780cc-121">Office 2013 ou version ultérieure sur Windows</span><span class="sxs-lookup"><span data-stu-id="780cc-121">Office 2013 or later on Windows</span></span>
- <span data-ttu-id="780cc-122">Office 2016 ou version ultérieure pour Mac</span><span class="sxs-lookup"><span data-stu-id="780cc-122">Office 2016 or later for Mac</span></span>

## <a name="next-steps"></a><span data-ttu-id="780cc-123">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="780cc-123">Next steps</span></span>

<span data-ttu-id="780cc-124">Lorsque vous êtes prêt à créer votre complément Office, reportez-vous au [démarrage rapide de 5 minutes](/office/dev/add-ins/#5-minute-quick-starts) pour votre application Office par défaut.</span><span class="sxs-lookup"><span data-stu-id="780cc-124">When you're ready to create your Office Add-in, see the [5-minute quick start](/office/dev/add-ins/#5-minute-quick-starts) for your preferred Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="780cc-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="780cc-125">See also</span></span>

- [<span data-ttu-id="780cc-126">Obtenir un laboratoire de script</span><span class="sxs-lookup"><span data-stu-id="780cc-126">Get Script Lab</span></span>](https://store.office.com/app.aspx?assetid=WA104380862)
- [<span data-ttu-id="780cc-127">En savoir plus sur script Lab</span><span class="sxs-lookup"><span data-stu-id="780cc-127">Learn more about Script Lab</span></span>](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [<span data-ttu-id="780cc-128">S’inscrire au programme de développement</span><span class="sxs-lookup"><span data-stu-id="780cc-128">Sign up for the dev program</span></span>](https://developer.microsoft.com/office/dev-program)
