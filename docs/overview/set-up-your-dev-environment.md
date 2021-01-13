---
title: Configuration de votre environnement de développement
description: Configurer votre environnement de développement pour créer des add-ins Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: eddf8bdf7b20a54667e6f8eb38bdace801ea1813
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839711"
---
# <a name="set-up-your-development-environment"></a><span data-ttu-id="62ce3-103">Configuration de votre environnement de développement</span><span class="sxs-lookup"><span data-stu-id="62ce3-103">Set up your development environment</span></span>

<span data-ttu-id="62ce3-104">Ce guide vous aide à configurer des outils pour vous aider à créer des add-ins Office en suivant nos démarrages rapides ou didacticiels.</span><span class="sxs-lookup"><span data-stu-id="62ce3-104">This guide helps you set up tools so you can create Office Add-ins by following our quick starts or tutorials.</span></span> <span data-ttu-id="62ce3-105">Vous devez installer les outils à partir de la liste ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="62ce3-105">You'll need to install the tools from the list below.</span></span> <span data-ttu-id="62ce3-106">Si vous avez déjà installé ces éléments, vous êtes prêt à commencer un démarrage rapide, tel que ce démarrage rapide [Excel React.](../quickstarts/excel-quickstart-react.md)</span><span class="sxs-lookup"><span data-stu-id="62ce3-106">If you already have these installed, you are ready to begin a quick start, such as this [Excel React quick start](../quickstarts/excel-quickstart-react.md).</span></span>

- <span data-ttu-id="62ce3-107">Node.js</span><span class="sxs-lookup"><span data-stu-id="62ce3-107">Node.js</span></span>
- <span data-ttu-id="62ce3-108">npm</span><span class="sxs-lookup"><span data-stu-id="62ce3-108">npm</span></span>
- <span data-ttu-id="62ce3-109">Un compte Microsoft 365 qui inclut la version d’abonnement d’Office</span><span class="sxs-lookup"><span data-stu-id="62ce3-109">A Microsoft 365 account which includes the subscription version of Office</span></span>
- <span data-ttu-id="62ce3-110">Éditeur de code de votre choix</span><span class="sxs-lookup"><span data-stu-id="62ce3-110">A code editor of your choice</span></span>

<span data-ttu-id="62ce3-111">Ce guide suppose que vous savez utiliser un outil de ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="62ce3-111">This guide assumes that you know how to use a command line tool.</span></span> 

## <a name="install-nodejs"></a><span data-ttu-id="62ce3-112">Installer Node.js.</span><span class="sxs-lookup"><span data-stu-id="62ce3-112">Install Node.js</span></span>

<span data-ttu-id="62ce3-113">Node.js est un runtime JavaScript dont vous aurez besoin pour développer des add-ins Office modernes.</span><span class="sxs-lookup"><span data-stu-id="62ce3-113">Node.js is a JavaScript runtime you will need to develop modern Office Add-ins.</span></span>

<span data-ttu-id="62ce3-114">Installez Node.js en [téléchargeant la dernière version recommandée à partir de leur site web.](https://nodejs.org)</span><span class="sxs-lookup"><span data-stu-id="62ce3-114">Install Node.js by [downloading the latest recommended version from their website](https://nodejs.org).</span></span> <span data-ttu-id="62ce3-115">Suivez les instructions d’installation de votre système d’exploitation.</span><span class="sxs-lookup"><span data-stu-id="62ce3-115">Follow the installation instructions for your operating system.</span></span>

## <a name="install-npm"></a><span data-ttu-id="62ce3-116">Installer npm</span><span class="sxs-lookup"><span data-stu-id="62ce3-116">Install npm</span></span>

<span data-ttu-id="62ce3-117">npm est un registre de logiciel open source à partir duquel télécharger les packages utilisés dans le développement de modules office.</span><span class="sxs-lookup"><span data-stu-id="62ce3-117">npm is an open source software registry from which to download the packages used in developing Office Add-ins.</span></span>

<span data-ttu-id="62ce3-118">Pour installer npm, exécutez la commande suivante dans la ligne de commande :</span><span class="sxs-lookup"><span data-stu-id="62ce3-118">To install npm, run the following in the command line:</span></span>

```command&nbsp;line
    npm install npm -g
```

<span data-ttu-id="62ce3-119">Pour vérifier si npm est déjà installé et voir la version installée, exécutez la commande suivante dans la ligne de commande :</span><span class="sxs-lookup"><span data-stu-id="62ce3-119">To check if you already have npm installed and see the installed version, run the following in the command line:</span></span>

```command&nbsp;line
npm -v
```

<span data-ttu-id="62ce3-120">Vous pouvez utiliser un gestionnaire de version Node pour vous permettre de basculer entre plusieurs versions de Node.js et npm, mais cela n’est pas strictement nécessaire.</span><span class="sxs-lookup"><span data-stu-id="62ce3-120">You may wish to use a Node version manager to allow you to switch between multiple versions of Node.js and npm, but this is not strictly necessary.</span></span> <span data-ttu-id="62ce3-121">Pour plus d’informations sur la façon de faire, voir [les instructions de npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)</span><span class="sxs-lookup"><span data-stu-id="62ce3-121">For details on how to do this, [see npm's instructions](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).</span></span>

## <a name="get-office-365"></a><span data-ttu-id="62ce3-122">Obtenir Office 365</span><span class="sxs-lookup"><span data-stu-id="62ce3-122">Get Office 365</span></span>

<span data-ttu-id="62ce3-123">Si vous n’avez pas déjà un compte Office 365, vous pouvez obtenir gratuitement un abonnement de 90 jours renouvelable de Microsoft 365 en rejoignant le [Programme pour les développeurs Microsoft 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="62ce3-123">If you don't already have a Microsoft 365 account, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="install-a-code-editor"></a><span data-ttu-id="62ce3-124">Installer un éditeur de code</span><span class="sxs-lookup"><span data-stu-id="62ce3-124">Install a code editor</span></span>

<span data-ttu-id="62ce3-125">Vous pouvez utiliser n’importe quel éditeur de code ou IDE qui prend en charge le développement côté client pour créer votre composant WebPart, par exemple :</span><span class="sxs-lookup"><span data-stu-id="62ce3-125">You can use any code editor or IDE that supports client-side development to build your web part, such as:</span></span>

- [<span data-ttu-id="62ce3-126">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="62ce3-126">Visual Studio Code</span></span>](https://code.visualstudio.com/)
- [<span data-ttu-id="62ce3-127">Atom</span><span class="sxs-lookup"><span data-stu-id="62ce3-127">Atom</span></span>](https://atom.io)
- [<span data-ttu-id="62ce3-128">Webstorm</span><span class="sxs-lookup"><span data-stu-id="62ce3-128">Webstorm</span></span>](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a><span data-ttu-id="62ce3-129">Étapes suivantes</span><span class="sxs-lookup"><span data-stu-id="62ce3-129">Next steps</span></span>

<span data-ttu-id="62ce3-130">Essayez de créer votre propre add-in ou utilisez Script Lab pour essayer des exemples intégrés.</span><span class="sxs-lookup"><span data-stu-id="62ce3-130">Try creating your own add-in or use Script Lab to try built-in samples.</span></span>

### <a name="create-an-office-add-in"></a><span data-ttu-id="62ce3-131">Créer un complément Office</span><span class="sxs-lookup"><span data-stu-id="62ce3-131">Create an Office add-in</span></span>

<span data-ttu-id="62ce3-132">Vous pouvez créer rapidement un complément de base pour Excel, OneNote, Outlook, PowerPoint, Project ou Word en effectuant un [démarrage rapide de 5 minutes](../index.yml).</span><span class="sxs-lookup"><span data-stu-id="62ce3-132">You can quickly create a basic add-in for Excel, OneNote, Outlook, PowerPoint, Project, or Word by completing a [5-minute quick start](../index.yml).</span></span> <span data-ttu-id="62ce3-133">Si vous avez déjà effectué un démarrage rapide et que vous voulez créer un complément légèrement plus complexe, vous devez essayer le [Didacticiel](../index.yml).</span><span class="sxs-lookup"><span data-stu-id="62ce3-133">If you've previously completed a quick start and want to create a slightly more complex add-in, you should try the [tutorial](../index.yml).</span></span>

### <a name="explore-the-apis-with-script-lab"></a><span data-ttu-id="62ce3-134">Explorez des API avec Script Lab</span><span class="sxs-lookup"><span data-stu-id="62ce3-134">Explore the APIs with Script Lab</span></span>

<span data-ttu-id="62ce3-135">Explorez la bibliothèque d’exemples intégrés dans [Script Lab](explore-with-script-lab.md) pour avoir une idée des capacités des API JavaScript Office.</span><span class="sxs-lookup"><span data-stu-id="62ce3-135">Explore the library of built-in samples in [Script Lab](explore-with-script-lab.md) to get a sense for the capabilities of the Office JavaScript APIs.</span></span>

## <a name="see-also"></a><span data-ttu-id="62ce3-136">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="62ce3-136">See also</span></span>

- [<span data-ttu-id="62ce3-137">Concepts de base pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="62ce3-137">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="62ce3-138">Développement de add-ins Office</span><span class="sxs-lookup"><span data-stu-id="62ce3-138">Developing Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="62ce3-139">Concevoir des compléments Office</span><span class="sxs-lookup"><span data-stu-id="62ce3-139">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="62ce3-140">Test et débogage de compléments Office</span><span class="sxs-lookup"><span data-stu-id="62ce3-140">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="62ce3-141">Publier des compléments Office</span><span class="sxs-lookup"><span data-stu-id="62ce3-141">Publish Office Add-ins</span></span>](../publish/publish.md)
- [<span data-ttu-id="62ce3-142">Découvrez le programme pour les développeurs Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="62ce3-142">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)