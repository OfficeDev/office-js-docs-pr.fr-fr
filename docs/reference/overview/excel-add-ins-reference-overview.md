---
title: Vue d’ensemble de l’API JavaScript pour Excel
description: En savoir plus sur l’API JavaScript pour Excel
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 80340b4990b56b2ba4d51f2a028480af3e267828
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650808"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="bf518-103">Vue d’ensemble de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="bf518-103">Excel JavaScript API overview</span></span>

<span data-ttu-id="bf518-104">Un complément Excel interagit avec des objets dans Excel en utilisant l’API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="bf518-104">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="bf518-105">**API JavaScript Excel** : il s’agit des [applications propres aux API](../../develop/application-specific-api-model.md) pour Excel.</span><span class="sxs-lookup"><span data-stu-id="bf518-105">**Excel JavaScript API**: These are the [application-specific APIs](../../develop/application-specific-api-model.md) for Excel.</span></span> <span data-ttu-id="bf518-106">Inclut dans Office 2016, l’[API JavaScript Excel](/javascript/api/excel) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="bf518-106">Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="bf518-107">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="bf518-107">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="bf518-108">Cette section de la documentation traite de l’API JavaScript pour Excel, que vous allez utiliser pour développer la majorité des fonctionnalités des compléments utilisés dans Excel sur le web ou dans Excel 2016 ou versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="bf518-108">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="bf518-109">Pour plus d’informations sur les API communes, voir le [Modèle objet des API JavaScript communes](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="bf518-109">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span>

## <a name="learn-object-model-concepts"></a><span data-ttu-id="bf518-110">Découvrir les concepts du modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="bf518-110">Learn object model concepts</span></span>

<span data-ttu-id="bf518-111">Voir [Modèle d’objet JavaScript Excel dans les compléments Office](../../excel/excel-add-ins-core-concepts.md) pour plus d’informations sur les concepts importants du modèle d’objet.</span><span class="sxs-lookup"><span data-stu-id="bf518-111">See [Excel JavaScript object model in Office Add-ins](../../excel/excel-add-ins-core-concepts.md) for information about important object model concepts.</span></span>

<span data-ttu-id="bf518-112">Pour apprendre à utiliser l’API JavaScript pour Excel afin d’accéder à des objets dans Excel, suivez le [didacticiel sur les compléments Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="bf518-112">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span>

## <a name="learn-api-capabilities"></a><span data-ttu-id="bf518-113">En savoir plus sur les fonctionnalités des API</span><span class="sxs-lookup"><span data-stu-id="bf518-113">Learn API capabilities</span></span>

<span data-ttu-id="bf518-114">Chaque fonctionnalité principale de l’API Excel inclut un article ou un ensemble d’articles sur la façon dont cette fonctionnalité et le modèle d’objet approprié sont utilisés.</span><span class="sxs-lookup"><span data-stu-id="bf518-114">Each major Excel API feature has an article or set of articles exploring what that feature can do and the relevant object model.</span></span>

* [<span data-ttu-id="bf518-115">Graphiques</span><span class="sxs-lookup"><span data-stu-id="bf518-115">Charts</span></span>](../../excel/excel-add-ins-charts.md)
* [<span data-ttu-id="bf518-116">Commentaires</span><span class="sxs-lookup"><span data-stu-id="bf518-116">Comments</span></span>](../../excel/excel-add-ins-comments.md)
* [<span data-ttu-id="bf518-117">Mise en forme conditionnelle</span><span class="sxs-lookup"><span data-stu-id="bf518-117">Conditional formatting</span></span>](../../excel/excel-add-ins-conditional-formatting.md)
* [<span data-ttu-id="bf518-118">Fonctions personnalisées</span><span class="sxs-lookup"><span data-stu-id="bf518-118">Custom functions</span></span>](../../excel/custom-functions-overview.md)
* [<span data-ttu-id="bf518-119">Validation des données</span><span class="sxs-lookup"><span data-stu-id="bf518-119">Data validation</span></span>](../../excel/excel-add-ins-data-validation.md)
* [<span data-ttu-id="bf518-120">Événements</span><span class="sxs-lookup"><span data-stu-id="bf518-120">Events</span></span>](../../excel/excel-add-ins-events.md)
* [<span data-ttu-id="bf518-121">PivotTables</span><span class="sxs-lookup"><span data-stu-id="bf518-121">PivotTables</span></span>](../../excel/excel-add-ins-pivottables.md)
* <span data-ttu-id="bf518-122">[Plages](../../excel/excel-add-ins-ranges-get.md) et [Cellules](../../excel/excel-add-ins-cells.md)</span><span class="sxs-lookup"><span data-stu-id="bf518-122">[Ranges](../../excel/excel-add-ins-ranges-get.md) and [Cells](../../excel/excel-add-ins-cells.md)</span></span>
* [<span data-ttu-id="bf518-123">RangeAreas (Plages multiples)</span><span class="sxs-lookup"><span data-stu-id="bf518-123">RangeAreas (Multiple ranges)</span></span>](../../excel/excel-add-ins-multiple-ranges.md)
* [<span data-ttu-id="bf518-124">Formes</span><span class="sxs-lookup"><span data-stu-id="bf518-124">Shapes</span></span>](../../excel/excel-add-ins-shapes.md)
* [<span data-ttu-id="bf518-125">Tableaux</span><span class="sxs-lookup"><span data-stu-id="bf518-125">Tables</span></span>](../../excel/excel-add-ins-tables.md)
* [<span data-ttu-id="bf518-126">Classeurs et API au niveau de l’application</span><span class="sxs-lookup"><span data-stu-id="bf518-126">Workbooks and Application-level APIs</span></span>](../../excel/excel-add-ins-workbooks.md)
* [<span data-ttu-id="bf518-127">Feuilles de calcul</span><span class="sxs-lookup"><span data-stu-id="bf518-127">Worksheets</span></span>](../../excel/excel-add-ins-worksheets.md)

<span data-ttu-id="bf518-128">Pour en savoir plus sur le modèle objet de l’API JavaScript pour Excel, consultez la [documentation de référence sur l’API JavaScript pour Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="bf518-128">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="bf518-129">Tester les exemples de code dans Script Lab</span><span class="sxs-lookup"><span data-stu-id="bf518-129">Try out code samples in Script Lab</span></span>

<span data-ttu-id="bf518-130">Utilisez [Script Lab](../../overview/explore-with-script-lab.md) pour commencer rapidement avec une collection d’exemples intégrés qui vous explique comment accomplir des tâches avec l’API.</span><span class="sxs-lookup"><span data-stu-id="bf518-130">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="bf518-131">Vous pouvez exécuter les exemples dans Script Lab de manière à afficher instantanément le résultat dans le volet Office ou la feuille de calcul, examiner les exemples pour découvrir le fonctionnement de l’API, voire utiliser les exemples pour prototyper votre propre complément.</span><span class="sxs-lookup"><span data-stu-id="bf518-131">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="bf518-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="bf518-132">See also</span></span>

* [<span data-ttu-id="bf518-133">Documentation sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="bf518-133">Excel add-ins documentation</span></span>](../../excel/index.yml)
* [<span data-ttu-id="bf518-134">Présentation des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="bf518-134">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
* [<span data-ttu-id="bf518-135">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="bf518-135">Excel JavaScript API reference</span></span>](/javascript/api/excel)
* [<span data-ttu-id="bf518-136">Application cliente Office et disponibilité de la plateforme pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="bf518-136">Office client application and platform availability for Office Add-ins</span></span>](../../overview/office-add-in-availability.md)
* [<span data-ttu-id="bf518-137">Utilisation du modèle API propre à l’application</span><span class="sxs-lookup"><span data-stu-id="bf518-137">Using the application-specific API model</span></span>](../../develop/application-specific-api-model.md)
