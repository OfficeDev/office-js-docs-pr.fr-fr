---
title: Fabric Core dans les Office de base
description: Obtenez une vue d’ensemble de l’utilisation de Fabric Core et des composants de l’interface utilisateur fabric dans Office des composants.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: e93efaea55841cc3bb6fa79ea1d1bbcaa76a4d05
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330199"
---
# <a name="fabric-core-in-office-add-ins"></a><span data-ttu-id="b12dc-103">Fabric Core dans les Office de base</span><span class="sxs-lookup"><span data-stu-id="b12dc-103">Fabric Core in Office Add-ins</span></span>

<span data-ttu-id="b12dc-104">Fabric Core est une collection open source de classes CSS et de mixins SASS conçus pour être utilisés dans des React *Office* non utilisés. Fabric Core contient des éléments de base du langage de conception de l’interface utilisateur Fluent, tels que les icônes, les couleurs, les polices et les grilles.</span><span class="sxs-lookup"><span data-stu-id="b12dc-104">Fabric Core is an open-source collection of CSS classes and SASS mixins that's *intended for use in non-React* Office Add-ins. Fabric Core contains basic elements of the Fluent UI design language such as icons, colors, typefaces, and grids.</span></span> <span data-ttu-id="b12dc-105">Fabric Core est indépendant de l’infrastructure, il peut donc être utilisé avec n’importe quelle application à page unique ou n’importe quelle infrastructure d’interface utilisateur web côté serveur.</span><span class="sxs-lookup"><span data-stu-id="b12dc-105">Fabric Core is framework independent, so it can be used with any single-page application or any server-side web UI framework.</span></span> <span data-ttu-id="b12dc-106">(Il est appelé « Fabric Core » au lieu de « Fluent Core » pour des raisons historiques.)</span><span class="sxs-lookup"><span data-stu-id="b12dc-106">(It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)</span></span>

<span data-ttu-id="b12dc-107">Si l’interface utilisateur de votre React n’est pas basée sur React, vous pouvez également utiliser un ensemble de composants React non utilisés.</span><span class="sxs-lookup"><span data-stu-id="b12dc-107">If your add-in's UI is not React-based, you can also make use of a set of non-React components.</span></span> <span data-ttu-id="b12dc-108">Voir [Utiliser Office composants JS UI Fabric.](#use-office-ui-fabric-js-components)</span><span class="sxs-lookup"><span data-stu-id="b12dc-108">See [Use Office UI Fabric JS components](#use-office-ui-fabric-js-components).</span></span>

> [!NOTE]
> <span data-ttu-id="b12dc-109">Cet article décrit l’utilisation de Fabric Core dans le contexte de Office des modules. Mais il est également utilisé dans un large éventail d’applications Microsoft 365 et d’extensions.</span><span class="sxs-lookup"><span data-stu-id="b12dc-109">This article describes the use of Fabric Core in the context of Office Add-ins. But it's also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="b12dc-110">Pour plus d’informations, [voir Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) et le repo open source Office [UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).</span><span class="sxs-lookup"><span data-stu-id="b12dc-110">For more information, see [Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="b12dc-111">Utiliser Fabric Core : icônes, polices, couleurs</span><span class="sxs-lookup"><span data-stu-id="b12dc-111">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="b12dc-112">Pour commencer à utiliser Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="b12dc-112">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="b12dc-113">Ajoutez la référence CDN au code HTML sur votre page.</span><span class="sxs-lookup"><span data-stu-id="b12dc-113">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="b12dc-114">Utilisez les polices et les icônes Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="b12dc-114">Use Fabric Core icons and fonts.</span></span>

    <span data-ttu-id="b12dc-115">Pour utiliser une icône Fabric Core, incluez l’élément « i » sur votre page, puis référencez les classes appropriées.</span><span class="sxs-lookup"><span data-stu-id="b12dc-115">To use a Fabric Core icon, include the "i" element on your page, and then reference the appropriate classes.</span></span> <span data-ttu-id="b12dc-116">Vous pouvez contrôler la taille de l’icône en modifiant la taille de police.</span><span class="sxs-lookup"><span data-stu-id="b12dc-116">You can control the size of the icon by changing the font size.</span></span> <span data-ttu-id="b12dc-117">Par exemple, le code suivant montre comment créer une icône de tableau extra large qui utilise la couleur themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="b12dc-117">For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="b12dc-118">Pour obtenir des instructions plus détaillées, voir [Icônes de l’interface utilisateur Fluent.](https://developer.microsoft.com/fluentui#/styles/web/icons)</span><span class="sxs-lookup"><span data-stu-id="b12dc-118">For more detailed instructions, see [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons).</span></span> <span data-ttu-id="b12dc-119">Pour trouver d’autres icônes disponibles dans Fabric Core, utilisez la fonctionnalité de recherche sur cette page.</span><span class="sxs-lookup"><span data-stu-id="b12dc-119">To find more icons that are available in Fabric Core, use the search feature on that page.</span></span> <span data-ttu-id="b12dc-120">Lorsque vous trouvez une icône à utiliser dans votre complément, veillez à précéder le nom de l’icône de `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="b12dc-120">When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="b12dc-121">Pour plus d’informations sur les tailles de police et les couleurs disponibles dans Fabric Core, voir [Typographie](https://developer.microsoft.com/fluentui#/styles/web/typography) et la table des matières **Couleurs** dans [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).</span><span class="sxs-lookup"><span data-stu-id="b12dc-121">For information about font sizes and colors that are available in Fabric Core, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).</span></span>

<span data-ttu-id="b12dc-122">Des exemples sont inclus dans [les exemples plus](#samples) loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="b12dc-122">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="use-office-ui-fabric-js-components"></a><span data-ttu-id="b12dc-123">Utiliser Office composants JS UI Fabric</span><span class="sxs-lookup"><span data-stu-id="b12dc-123">Use Office UI Fabric JS components</span></span>

<span data-ttu-id="b12dc-124">Les applications avec des interfaces utilisateur non React peuvent également utiliser l’un des nombreux composants de [Office UI Fabric JS,](https://github.com/OfficeDev/office-ui-fabric-js)y compris les boutons, les boîtes de dialogue, les suceurs et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="b12dc-124">Add-ins with non-React UIs can also use any of the many components from [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), including buttons, dialogs, pickers, and more.</span></span> <span data-ttu-id="b12dc-125">Consultez le lisez-moi du repo pour obtenir des instructions.</span><span class="sxs-lookup"><span data-stu-id="b12dc-125">See the readme of the repo for instructions.</span></span>

<span data-ttu-id="b12dc-126">Des exemples sont inclus dans [les exemples plus](#samples) loin dans cet article.</span><span class="sxs-lookup"><span data-stu-id="b12dc-126">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="samples"></a><span data-ttu-id="b12dc-127">Exemples</span><span class="sxs-lookup"><span data-stu-id="b12dc-127">Samples</span></span>

<span data-ttu-id="b12dc-128">Les exemples de composants suivants utilisent Fabric Core et/ou Office composants JS UI Fabric.</span><span class="sxs-lookup"><span data-stu-id="b12dc-128">The following sample add-ins use Fabric Core and/or Office UI Fabric JS components.</span></span> <span data-ttu-id="b12dc-129">Certains de ces dépôts sont archivés, ce qui signifie qu’ils ne sont plus mis à jour avec des correctifs de bogue ou de sécurité, mais vous pouvez toujours les utiliser pour apprendre à utiliser les composants d’interface utilisateur Fabric Core et Fabric.</span><span class="sxs-lookup"><span data-stu-id="b12dc-129">Some of these repos are archived, meaning that they are no longer being updated with bug or security fixes, but you can still use them to learn how to use Fabric Core and Fabric UI components.</span></span>

- [<span data-ttu-id="b12dc-130">Excel Add-in JavaScript SalesTracker</span><span class="sxs-lookup"><span data-stu-id="b12dc-130">Excel Add-in JavaScript SalesTracker</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [<span data-ttu-id="b12dc-131">Excel Add-in SalesLeads</span><span class="sxs-lookup"><span data-stu-id="b12dc-131">Excel Add-in SalesLeads</span></span>](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [<span data-ttu-id="b12dc-132">Excel Tendances des dépenses de WoodGrove du add-in</span><span class="sxs-lookup"><span data-stu-id="b12dc-132">Excel Add-in WoodGrove Expense Trends</span></span>](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [<span data-ttu-id="b12dc-133">Excel Content Add-in Humongous Insurance</span><span class="sxs-lookup"><span data-stu-id="b12dc-133">Excel Content Add-in Humongous Insurance</span></span>](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [<span data-ttu-id="b12dc-134">Office Exemple d’interface utilisateur de la structure de la structure de la add-in</span><span class="sxs-lookup"><span data-stu-id="b12dc-134">Office Add-in Fabric UI Sample</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="b12dc-135">Office-Add-in-UX-Design-Patterns-Code</span><span class="sxs-lookup"><span data-stu-id="b12dc-135">Office-Add-in-UX-Design-Patterns-Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="b12dc-136">Outlook Add-in GifMe</span><span class="sxs-lookup"><span data-stu-id="b12dc-136">Outlook Add-in GifMe</span></span>](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [<span data-ttu-id="b12dc-137">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span><span class="sxs-lookup"><span data-stu-id="b12dc-137">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="b12dc-138">Word Add-in Angular2 StyleChecker</span><span class="sxs-lookup"><span data-stu-id="b12dc-138">Word Add-in Angular2 StyleChecker</span></span>](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [<span data-ttu-id="b12dc-139">Word Add-in JS Redact</span><span class="sxs-lookup"><span data-stu-id="b12dc-139">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="b12dc-140">Word Add-in MarkdownConversion</span><span class="sxs-lookup"><span data-stu-id="b12dc-140">Word Add-in MarkdownConversion</span></span>](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
