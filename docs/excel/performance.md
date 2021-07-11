---
title: Optimisation des performances API JavaScript Excel
description: Optimisez Excel de votre application à l’aide de l’API JavaScript.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 5313bb3fe25d165e49cc0508e81d58294db48798
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349384"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="c995d-103">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="c995d-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="c995d-104">Il existe plusieurs façons d’effectuer des tâches courantes avec l’API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="c995d-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="c995d-105">Vous trouverez des différences de performances significatives entre les différentes approches.</span><span class="sxs-lookup"><span data-stu-id="c995d-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="c995d-106">Cet article fournit des instructions et exemples de code pour vous montrer comment effectuer des tâches courantes efficacement à l’aide des API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="c995d-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c995d-107">De nombreux problèmes de performances peuvent être résolus par le biais d’une utilisation recommandée `load` et `sync` d’appels.</span><span class="sxs-lookup"><span data-stu-id="c995d-107">Many performance issues can be addressed through recommended usage of `load` and `sync` calls.</span></span> <span data-ttu-id="c995d-108">Consultez la section « Améliorations des performances avec les API propres à l’application » des limites de ressources et de l’optimisation des performances pour les Office Pour obtenir des [conseils](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) sur l’utilisation efficace des API propres à l’application.</span><span class="sxs-lookup"><span data-stu-id="c995d-108">See the "Performance improvements with the application-specific APIs" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) for advice on working with the application-specific APIs in an efficient way.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="c995d-109">Suspendre temporairement les processus Excel</span><span class="sxs-lookup"><span data-stu-id="c995d-109">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="c995d-110">Excel a des tâches en arrière-plan qui réagissent à l’entrée des utilisateurs et de votre complément.</span><span class="sxs-lookup"><span data-stu-id="c995d-110">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="c995d-111">Certains de ces processus Excel peuvent être contrôlés pour accroître les performances.</span><span class="sxs-lookup"><span data-stu-id="c995d-111">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="c995d-112">Ceci est particulièrement utile lorsque votre complément utilise de grands ensembles de données.</span><span class="sxs-lookup"><span data-stu-id="c995d-112">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="c995d-113">Suspendre temporairement les calculs</span><span class="sxs-lookup"><span data-stu-id="c995d-113">Suspend calculation temporarily</span></span>

<span data-ttu-id="c995d-114">Si vous essayez d’effectuer une opération sur un grand nombre de cellules (par exemple, la définition de la valeur d’un objet plage énorme) et que vous n’avez rien contre la suspension de calcul dans Excel temporairement le temps que votre opération se termine, nous vous recommandons de suspendre le calcul jusqu’à ce que le prochain `context.sync()` soit appelé.</span><span class="sxs-lookup"><span data-stu-id="c995d-114">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="c995d-115">Reportez-vous à la documentation de référence [Objet Application](/javascript/api/excel/excel.application) pour plus d’informations sur l’utilisation de l’API `suspendApiCalculationUntilNextSync()` pour suspendre et réactiver les calculs de manière très pratique.</span><span class="sxs-lookup"><span data-stu-id="c995d-115">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="c995d-116">Le code suivant montre comment suspendre temporairement le calcul.</span><span class="sxs-lookup"><span data-stu-id="c995d-116">The following code demonstrates how to suspend calculation temporarily.</span></span>

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

<span data-ttu-id="c995d-117">Notez que seuls les calculs de formule sont suspendus.</span><span class="sxs-lookup"><span data-stu-id="c995d-117">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="c995d-118">Toutes les références modifiées sont toujours reconstruites.</span><span class="sxs-lookup"><span data-stu-id="c995d-118">Any altered references are still rebuilt.</span></span> <span data-ttu-id="c995d-119">Par exemple, le fait de renommer une feuille de calcul met toujours à jour les références dans les formules de cette feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="c995d-119">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="c995d-120">Suspendre la mise à jour de l’écran</span><span class="sxs-lookup"><span data-stu-id="c995d-120">Suspend screen updating</span></span>

<span data-ttu-id="c995d-121">Excel affiche les modifications effectuées par votre complément à peu près au moment où elles ont lieu dans le code.</span><span class="sxs-lookup"><span data-stu-id="c995d-121">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="c995d-122">Dans le cas de grands ensembles de données itératifs, il se peut que vous ne deviez pas afficher cette progression sur l’écran en temps réel.</span><span class="sxs-lookup"><span data-stu-id="c995d-122">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="c995d-123">`Application.suspendScreenUpdatingUntilNextSync()` interrompt les mises à jour visuelles vers Excel tant que le complément n’appelle pas `context.sync()`, ou tant que `Excel.run` ne se termine pas (appelant implicitement `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="c995d-123">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="c995d-124">N’oubliez pas qu'Excel n’affiche aucun signe d’activité jusqu'à la synchronisation suivante. Votre complément doit donner des conseils aux utilisateurs pour les préparer à ce délai ou fournir une barre d’état pour démontrer l’activité.</span><span class="sxs-lookup"><span data-stu-id="c995d-124">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="c995d-125">N’appelez `suspendScreenUpdatingUntilNextSync` pas à plusieurs reprises (par exemple, dans une boucle).</span><span class="sxs-lookup"><span data-stu-id="c995d-125">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="c995d-126">Les appels répétés entraînent le clignotement Excel fenêtre.</span><span class="sxs-lookup"><span data-stu-id="c995d-126">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="c995d-127">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="c995d-127">Enable and disable events</span></span>

<span data-ttu-id="c995d-128">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="c995d-128">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="c995d-129">Un exemple de code montrant comment activer et désactiver les événements dans l’article[manipuler les événements](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="c995d-129">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="c995d-130">Importation de données dans des tableaux</span><span class="sxs-lookup"><span data-stu-id="c995d-130">Importing data into tables</span></span>

<span data-ttu-id="c995d-131">Lorsque vous tentez d’importer une quantité considérable de données directement dans un objet[tableau](/javascript/api/excel/excel.table) (par exemple, à l’aide de `TableRowCollection.add()`), vous risquez de rencontrer une dégradation des performances.</span><span class="sxs-lookup"><span data-stu-id="c995d-131">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="c995d-132">Si vous essayez d’ajouter un nouveau tableau, vous devez remplir les données d’abord en définissant `range.values`, puis appelez `worksheet.tables.add()` pour créer un tableau sur la plage.</span><span class="sxs-lookup"><span data-stu-id="c995d-132">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="c995d-133">Si vous essayez d’écrire des données dans un tableau existant, écrivez les données dans un objet plage via `table.getDataBodyRange()`, et le tableau s’agrandit automatiquement.</span><span class="sxs-lookup"><span data-stu-id="c995d-133">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span>

<span data-ttu-id="c995d-134">Voici un exemple de cette approche :</span><span class="sxs-lookup"><span data-stu-id="c995d-134">Here is an example of this approach:</span></span>

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> <span data-ttu-id="c995d-135">Vous pouvez convertir un objet de Tableau en objet de Plage à l’aide de la méthode[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="c995d-135">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="c995d-136">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c995d-136">See also</span></span>

* [<span data-ttu-id="c995d-137">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="c995d-137">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
* [<span data-ttu-id="c995d-138">Limites des ressources et optimisation des performances pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="c995d-138">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
* [<span data-ttu-id="c995d-139">Objet de fonctions de feuille de calcul (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="c995d-139">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
