---
title: Configuration de votre environnement de développement
description: Configurez votre environnement de développement pour créer des compléments Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 644194d7d0da479b13ac09d7e830af53e9a9838e
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740832"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="c87f3-103">Configuration de votre environnement de développement</span><span class="sxs-lookup"><span data-stu-id="c87f3-103">Set up your development environment</span></span>

<span data-ttu-id="c87f3-104">Ce guide vous aide à configurer les outils de manière à pouvoir créer des compléments Office en suivant nos guides de démarrage rapide ou nos didacticiels.</span><span class="sxs-lookup"><span data-stu-id="c87f3-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="c87f3-105">Vous devrez installer les outils à partir de la liste ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="c87f3-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="c87f3-106">Si ces éléments sont déjà installés, vous êtes prêt à commencer un démarrage rapide, tel que le [démarrage rapide de Microsoft Excel REACT](../quickstarts/excel-quickstart-react.md).</span><span class="sxs-lookup"><span data-stu-id="c87f3-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="c87f3-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="c87f3-107">Node.js</span></span>
- <span data-ttu-id="c87f3-108">npm</span><span class="sxs-lookup"><span data-stu-id="c87f3-108">npm</span></span>
- <span data-ttu-id="c87f3-109">Un compte Microsoft 365 qui inclut la version d’abonnement d’Office</span><span class="sxs-lookup"><span data-stu-id="c87f3-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="c87f3-110">Un éditeur de code de votre choix</span><span class="sxs-lookup"><span data-stu-id="c87f3-110">A code editor of your choice</span></span>

<span data-ttu-id="c87f3-111">Ce guide suppose que vous sachiez comment utiliser un outil de ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="c87f3-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="c87f3-112">Installer Node.js.</span><span class="sxs-lookup"><span data-stu-id="c87f3-112">Install Node.js</span></span>

<span data-ttu-id="c87f3-113">Node.js est un Runtime JavaScript dont vous aurez besoin pour développer des compléments Office modernes.</span><span class="sxs-lookup"><span data-stu-id="c87f3-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="c87f3-114">Installez Node.js en [téléchargeant la version recommandée la plus récente à partir de leur site Web](https://nodejs.org).</span><span class="sxs-lookup"><span data-stu-id="c87f3-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="c87f3-115">Suivez les instructions d’installation pour votre système d’exploitation.</span><span class="sxs-lookup"><span data-stu-id="c87f3-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="c87f3-116">Installer NPM</span><span class="sxs-lookup"><span data-stu-id="c87f3-116">Install npm</span></span>

<span data-ttu-id="c87f3-117">NPM est un registre de logiciels open source à partir duquel télécharger les packages utilisés dans le développement des compléments Office.</span><span class="sxs-lookup"><span data-stu-id="c87f3-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="c87f3-118">Pour installer NPM, exécutez ce qui suit dans la ligne de commande :</span><span class="sxs-lookup"><span data-stu-id="c87f3-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="c87f3-119">Pour vérifier si NPM est déjà installé et voir la version installée, exécutez la commande suivante dans la ligne de commande :</span><span class="sxs-lookup"><span data-stu-id="c87f3-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="c87f3-120">Vous souhaiterez peut-être utiliser un gestionnaire de version de nœud pour vous permettre de basculer entre plusieurs versions de Node.js et NPM, mais ce n’est pas obligatoire.</span><span class="sxs-lookup"><span data-stu-id="c87f3-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="c87f3-121">Pour plus d’informations sur la procédure à suivre, [consultez les instructions de NPM](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span><span class="sxs-lookup"><span data-stu-id="c87f3-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="c87f3-122">Obtenir Office 365</span><span class="sxs-lookup"><span data-stu-id="c87f3-122">Get Office 365</span></span>

<span data-ttu-id="c87f3-123">Si vous n’avez pas déjà un compte Office 365, vous pouvez obtenir gratuitement un abonnement de 90 jours renouvelable de Microsoft 365 en rejoignant le [Programme pour les développeurs Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="c87f3-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="c87f3-124">Installer un éditeur de code</span><span class="sxs-lookup"><span data-stu-id="c87f3-124">Install a code editor</span></span>

<span data-ttu-id="c87f3-125">Vous pouvez utiliser n’importe quel éditeur de code ou IDE qui prend en charge le développement côté client pour créer votre composant WebPart, par exemple :</span><span class="sxs-lookup"><span data-stu-id="c87f3-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="c87f3-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="c87f3-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="c87f3-127">Atom</span><span class="sxs-lookup"><span data-stu-id="c87f3-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="c87f3-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="c87f3-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="c87f3-129">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="c87f3-129">Next steps</span></span>

<span data-ttu-id="c87f3-130">Essayez de créer votre propre complément ou utilisez script Lab pour essayer des exemples intégrés.</span><span class="sxs-lookup"><span data-stu-id="c87f3-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="c87f3-131">Créer un complément Office</span><span class="sxs-lookup"><span data-stu-id="c87f3-131">Create an Office add-in</span></span>

<span data-ttu-id="c87f3-132">Vous pouvez créer rapidement un complément de base pour Excel, OneNote, Outlook, PowerPoint, Project ou Word en effectuant un [démarrage rapide de 5 minutes](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="c87f3-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](/office/dev/add-ins/).</span></span> <span data-ttu-id="c87f3-133">Si vous avez déjà effectué un démarrage rapide et que vous voulez créer un complément légèrement plus complexe, vous devez essayer le [Didacticiel](/office/dev/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="c87f3-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](/office/dev/add-ins/).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="c87f3-134">Explorez des API avec Script Lab</span><span class="sxs-lookup"><span data-stu-id="c87f3-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="c87f3-135">Explorez la bibliothèque d’exemples intégrés dans [Script Lab](explore-with-script-lab.md) pour avoir une idée des capacités des API JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="c87f3-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="c87f3-136">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c87f3-136">See also</span></span>

- [<span data-ttu-id="c87f3-137">Concepts de base pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="c87f3-137">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="c87f3-138">Développement de compléments Office</span><span class="sxs-lookup"><span data-stu-id="c87f3-138">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="c87f3-139">Concevoir des compléments Office</span><span class="sxs-lookup"><span data-stu-id="c87f3-139">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="c87f3-140">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="c87f3-140">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="c87f3-141">Publier des compléments Office</span><span class="sxs-lookup"><span data-stu-id="c87f3-141">Publish Office Add-ins</span></span>](../publish/publish.md)
- [<span data-ttu-id="c87f3-142">En savoir plus sur le programme de développement Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="c87f3-142">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
