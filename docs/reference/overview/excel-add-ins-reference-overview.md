---
title: Vue d’ensemble de l’API JavaScript pour Excel
description: ''
ms.date: 02/19/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: aa0e18025e539f747851f5dc9f5a25197761c5c8
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163618"
---
# <a name="excel-javascript-api-overview"></a><span data-ttu-id="f2831-102">Vue d’ensemble de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="f2831-102">Excel JavaScript API overview</span></span>

<span data-ttu-id="f2831-103">Un complément Excel interagit avec des objets dans Excel à l’aide de l’API JavaScript pour Office, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="f2831-103">An Excel add-in interacts with objects in Excel by using the JavaScript API for Office, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="f2831-104">**API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](/javascript/api/excel) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="f2831-104">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](/javascript/api/excel) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="f2831-105">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) permettent d’accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="f2831-105">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="f2831-106">Cette section de la documentation traite de l’API JavaScript pour Excel, que vous allez utiliser pour développer la majorité des fonctionnalités des compléments utilisés dans Excel sur le web ou dans Excel 2016 ou versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="f2831-106">This section of the documentation focuses on the Excel JavaScript API, which you'll use to develop the majority of functionality in add-ins that target Excel on the web or Excel 2016 or later.</span></span> <span data-ttu-id="f2831-107">Pour plus d’informations sur les API communes, voir le [Modèle objet des API JavaScript communes](../../develop/office-javascript-api-object-model.md).</span><span class="sxs-lookup"><span data-stu-id="f2831-107">For information about the Common API, see [Common JavaScript API object model](../../develop/office-javascript-api-object-model.md).</span></span> 

## <a name="learn-programming-concepts"></a><span data-ttu-id="f2831-108">Découvrir les concepts de programmation</span><span class="sxs-lookup"><span data-stu-id="f2831-108">Learn programming concepts</span></span>

<span data-ttu-id="f2831-109">Pour plus d’informations sur les concepts de programmation essentiels, consultez les articles suivants :</span><span class="sxs-lookup"><span data-stu-id="f2831-109">See the following articles for information about important programming concepts:</span></span>
 
- [<span data-ttu-id="f2831-110">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="f2831-110">Fundamental programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-core-concepts.md)

- [<span data-ttu-id="f2831-111">Concepts avancés de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="f2831-111">Advanced programming concepts with the Excel JavaScript API</span></span>](../../excel/excel-add-ins-advanced-concepts.md)

## <a name="learn-about-api-capabilities"></a><span data-ttu-id="f2831-112">En savoir plus sur les fonctionnalités des API</span><span class="sxs-lookup"><span data-stu-id="f2831-112">Learn about API capabilities</span></span>

<span data-ttu-id="f2831-113">Reportez-vous aux autres articles présents dans cette section de la documentation pour apprendre à utiliser les [événements](../../excel/excel-add-ins-events.md), les [graphiques](../../excel/excel-add-ins-charts.md), les [plages](../../excel/excel-add-ins-ranges.md), les [tableaux](../../excel/excel-add-ins-tables.md), les [feuilles de calcul](../../excel/excel-add-ins-worksheets.md), etc.</span><span class="sxs-lookup"><span data-stu-id="f2831-113">Use other articles in this section of the documentation to learn about working with [events](../../excel/excel-add-ins-events.md), [charts](../../excel/excel-add-ins-charts.md), [ranges](../../excel/excel-add-ins-ranges.md), [tables](../../excel/excel-add-ins-tables.md), [worksheets](../../excel/excel-add-ins-worksheets.md), and more.</span></span> <span data-ttu-id="f2831-114">Vous trouverez également dans cette section des conseils sur les concepts relatifs à l’API JavaScript pour Excel, tels que la [co-édition dans les compléments Excel](../../excel/co-authoring-in-excel-add-ins.md), la [validation des données](../../excel/excel-add-ins-data-validation.md), la [gestion des erreurs](../../excel/excel-add-ins-error-handling.md) et l’[optimisation des performances](../../excel/performance.md).</span><span class="sxs-lookup"><span data-stu-id="f2831-114">Also in this section, you'll find guidance about Excel JavaScript API concepts such as [coauthoring in Excel add-ins](../../excel/co-authoring-in-excel-add-ins.md), [data validation](../../excel/excel-add-ins-data-validation.md), [error handling](../../excel/excel-add-ins-error-handling.md), and [performance optimization](../../excel/performance.md).</span></span> <span data-ttu-id="f2831-115">Reportez-vous à la table des matières pour obtenir la liste complète des articles disponibles.</span><span class="sxs-lookup"><span data-stu-id="f2831-115">See the table of contents for the complete list of available articles.</span></span>

<span data-ttu-id="f2831-116">Pour apprendre à utiliser l’API JavaScript pour Excel afin d’accéder à des objets dans Excel, suivez le [didacticiel sur les compléments Excel](../../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="f2831-116">For hands-on experience using the Excel JavaScript API to access objects in Excel, complete the [Excel add-in tutorial](../../tutorials/excel-tutorial.md).</span></span> 

<span data-ttu-id="f2831-117">Pour en savoir plus sur le modèle objet de l’API JavaScript pour Excel, consultez la [documentation de référence sur l’API JavaScript pour Excel](/javascript/api/excel).</span><span class="sxs-lookup"><span data-stu-id="f2831-117">For detailed information about the Excel JavaScript API object model, see the [Excel JavaScript API reference documentation](/javascript/api/excel).</span></span>

## <a name="try-out-code-samples-in-script-lab"></a><span data-ttu-id="f2831-118">Tester les exemples de code dans Script Lab</span><span class="sxs-lookup"><span data-stu-id="f2831-118">Try out code samples in Script Lab</span></span>

<span data-ttu-id="f2831-119">Utilisez [Script Lab](../../overview/explore-with-script-lab.md) pour commencer rapidement avec une collection d’exemples intégrés qui vous explique comment accomplir des tâches avec l’API.</span><span class="sxs-lookup"><span data-stu-id="f2831-119">Use [Script Lab](../../overview/explore-with-script-lab.md) to get started quickly with a collection of built-in samples that show how to complete tasks with the API.</span></span> <span data-ttu-id="f2831-120">Vous pouvez exécuter les exemples dans Script Lab de manière à afficher instantanément le résultat dans le volet Office ou la feuille de calcul, examiner les exemples pour découvrir le fonctionnement de l’API, voire utiliser les exemples pour prototyper votre propre complément.</span><span class="sxs-lookup"><span data-stu-id="f2831-120">You can run the samples in Script Lab to instantly see the result in the task pane or worksheet, examine the samples to learn how the API works, and even use samples to prototype your own add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="f2831-121">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f2831-121">See also</span></span>

- [<span data-ttu-id="f2831-122">Documentation sur les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="f2831-122">Excel add-ins documentation</span></span>](../../excel/index.md)
- [<span data-ttu-id="f2831-123">Présentation des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="f2831-123">Excel add-ins overview</span></span>](../../excel/excel-add-ins-overview.md)
- [<span data-ttu-id="f2831-124">Référence sur l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="f2831-124">Excel JavaScript API reference</span></span>](/javascript/api/excel)
- [<span data-ttu-id="f2831-125">Disponibilité des compléments Office sur les plateformes et les hôtes</span><span class="sxs-lookup"><span data-stu-id="f2831-125">Office Add-in host and platform availability</span></span>](../../overview/office-add-in-availability.md)
- [<span data-ttu-id="f2831-126">Spécifications ouvertes des API</span><span class="sxs-lookup"><span data-stu-id="f2831-126">API open specifications</span></span>](../openspec/openspec.md)
