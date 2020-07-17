---
title: Optimisation des performances API JavaScript Excel
description: Optimisation des performances à l’aide de l’API JavaScript d’Excel
ms.date: 07/14/2020
localization_priority: Normal
ms.openlocfilehash: 193cbe8c8cd1a432c6567401ed645990cb93e5e9
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159093"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="1c98b-103">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="1c98b-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="1c98b-104">Il existe plusieurs façons d’effectuer des tâches courantes avec l’API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="1c98b-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="1c98b-105">Vous trouverez des différences de performances significatives entre les différentes approches.</span><span class="sxs-lookup"><span data-stu-id="1c98b-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="1c98b-106">Cet article fournit des instructions et exemples de code pour vous montrer comment effectuer des tâches courantes efficacement à l’aide des API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="1c98b-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="1c98b-107">Limitez le nombre d’appels sync()</span><span class="sxs-lookup"><span data-stu-id="1c98b-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="1c98b-108">Dans l’API JavaScript Excel, `sync()` est la seule opération asynchrone et elle peut être lente dans certaines circonstances, en particulier pour Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="1c98b-108">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="1c98b-109">Pour optimiser les performances, vous devez limiter le nombre de fois que vous appelez `sync()` et mettre en file d’attente autant de modifications que possible avant d’appeler.</span><span class="sxs-lookup"><span data-stu-id="1c98b-109">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="1c98b-110">Voir [Concepts principaux - sync()](excel-add-ins-core-concepts.md#sync) pour des exemples de code qui suivent cette pratique.</span><span class="sxs-lookup"><span data-stu-id="1c98b-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="1c98b-111">Réduire le nombre d’objets proxy créés</span><span class="sxs-lookup"><span data-stu-id="1c98b-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="1c98b-112">Éviter de créer le même objet proxy à plusieurs reprises.</span><span class="sxs-lookup"><span data-stu-id="1c98b-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="1c98b-113">Au lieu de cela, si vous avez besoin du même objet proxy pour plus d’une opération, créez-le une seule fois et affectez-le à une variable, puis utilisez cette variable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="1c98b-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
// BAD: repeated calls to .getRange() to create the same proxy object
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: create the range proxy object once and assign to a variable
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

## <a name="load-necessary-properties-only"></a><span data-ttu-id="1c98b-114">Charger les propriétés nécessaires uniquement</span><span class="sxs-lookup"><span data-stu-id="1c98b-114">Load necessary properties only</span></span>

<span data-ttu-id="1c98b-115">Dans l’API JavaScript Excel, vous devez charger explicitement les propriétés d’un objet proxy.</span><span class="sxs-lookup"><span data-stu-id="1c98b-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="1c98b-116">Bien que vous soyez en mesure de charger les propriétés en une fois avec un appel vide`load()`, cette approche peut causer une surcharge significative des performances.</span><span class="sxs-lookup"><span data-stu-id="1c98b-116">Although you're able to load all the properties at once with an empty `load()` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="1c98b-117">Au lieu de cela, nous vous conseillons de charger uniquement les propriétés nécessaires, en particulier pour ces objets qui ont un grand nombre de propriétés.</span><span class="sxs-lookup"><span data-stu-id="1c98b-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="1c98b-118">Par exemple, si vous avez uniquement l’intention de lire la `address` propriété d’un objet Range, spécifiez uniquement cette propriété lorsque vous appelez la `load()` méthode :</span><span class="sxs-lookup"><span data-stu-id="1c98b-118">For example, if you only intend to read the `address` property of a range object, specify only that property when you call the `load()` method:</span></span>

```js
range.load('address');
```

<span data-ttu-id="1c98b-119">Vous pouvez appeler `load()` la méthode de l’une des manières suivantes :</span><span class="sxs-lookup"><span data-stu-id="1c98b-119">You can call `load()` method in any of the following ways:</span></span>

<span data-ttu-id="1c98b-120">_Syntaxe :_</span><span class="sxs-lookup"><span data-stu-id="1c98b-120">_Syntax:_</span></span>

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

<span data-ttu-id="1c98b-121">_Où :_</span><span class="sxs-lookup"><span data-stu-id="1c98b-121">_Where:_</span></span>

* <span data-ttu-id="1c98b-122">`properties` est la liste des propriétés à charger, fournie sous forme de chaînes séparées par des virgules ou de tableau de noms.</span><span class="sxs-lookup"><span data-stu-id="1c98b-122">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names.</span></span> <span data-ttu-id="1c98b-123">Pour plus d’informations, consultez les `load()` méthodes définies pour les objets dans la référence de l' [API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md).</span><span class="sxs-lookup"><span data-stu-id="1c98b-123">For more information, see the `load()` methods defined for objects in [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md).</span></span>
* <span data-ttu-id="1c98b-p106">`loadOption` spécifie un objet qui décrit les options select, expand, top et skip. Pour plus d’informations, reportez-vous aux [options](/javascript/api/office/officeextension.loadoption) de chargement d’objet.</span><span class="sxs-lookup"><span data-stu-id="1c98b-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="1c98b-126">N’oubliez pas que certaines des « propriétés » sous un objet peuvent avoir le même nom qu’un autre objet.</span><span class="sxs-lookup"><span data-stu-id="1c98b-126">Please be aware that some of the "properties" under an object may have the same name as another object.</span></span> <span data-ttu-id="1c98b-127">Par exemple, `format` est une propriété sous plage d’objet, mais `format` lui-même est également un objet.</span><span class="sxs-lookup"><span data-stu-id="1c98b-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="1c98b-128">Par conséquent, si vous passez un appel comme `range.load("format")`, cela équivaut à `range.format.load()`, c'est-à-dire, un appel load() vide pouvant entraîner des problèmes de performances comme indiqué précédemment.</span><span class="sxs-lookup"><span data-stu-id="1c98b-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="1c98b-129">Pour éviter cela, votre code doit uniquement charger les « nœuds feuille » dans une arborescence d’objets.</span><span class="sxs-lookup"><span data-stu-id="1c98b-129">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="1c98b-130">Suspendre temporairement les processus Excel</span><span class="sxs-lookup"><span data-stu-id="1c98b-130">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="1c98b-131">Excel a des tâches en arrière-plan qui réagissent à l’entrée des utilisateurs et de votre complément.</span><span class="sxs-lookup"><span data-stu-id="1c98b-131">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="1c98b-132">Certains de ces processus Excel peuvent être contrôlés pour accroître les performances.</span><span class="sxs-lookup"><span data-stu-id="1c98b-132">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="1c98b-133">Ceci est particulièrement utile lorsque votre complément utilise de grands ensembles de données.</span><span class="sxs-lookup"><span data-stu-id="1c98b-133">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="1c98b-134">Suspendre temporairement les calculs</span><span class="sxs-lookup"><span data-stu-id="1c98b-134">Suspend calculation temporarily</span></span>

<span data-ttu-id="1c98b-135">Si vous essayez d’effectuer une opération sur un grand nombre de cellules (par exemple, la définition de la valeur d’un objet plage énorme) et que vous n’avez rien contre la suspension de calcul dans Excel temporairement le temps que votre opération se termine, nous vous recommandons de suspendre le calcul jusqu’à ce que le prochain `context.sync()` soit appelé.</span><span class="sxs-lookup"><span data-stu-id="1c98b-135">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="1c98b-136">Reportez-vous à la documentation de référence [Objet Application](/javascript/api/excel/excel.application) pour plus d’informations sur l’utilisation de l’API `suspendApiCalculationUntilNextSync()` pour suspendre et réactiver les calculs de manière très pratique.</span><span class="sxs-lookup"><span data-stu-id="1c98b-136">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="1c98b-137">Le code suivant montre comment suspendre temporairement le calcul :</span><span class="sxs-lookup"><span data-stu-id="1c98b-137">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

<span data-ttu-id="1c98b-138">Veuillez noter que seuls les calculs de formule sont suspendus.</span><span class="sxs-lookup"><span data-stu-id="1c98b-138">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="1c98b-139">Toutes les références modifiées sont toujours recréées.</span><span class="sxs-lookup"><span data-stu-id="1c98b-139">Any altered references are still rebuilt.</span></span> <span data-ttu-id="1c98b-140">Par exemple, si vous renommez une feuille de calcul, les références des formules sont toujours mises à jour dans cette feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="1c98b-140">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="1c98b-141">Suspendre la mise à jour de l’écran</span><span class="sxs-lookup"><span data-stu-id="1c98b-141">Suspend screen updating</span></span>

<span data-ttu-id="1c98b-142">Excel affiche les modifications effectuées par votre complément à peu près au moment où elles ont lieu dans le code.</span><span class="sxs-lookup"><span data-stu-id="1c98b-142">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="1c98b-143">Dans le cas de grands ensembles de données itératifs, il se peut que vous ne deviez pas afficher cette progression sur l’écran en temps réel.</span><span class="sxs-lookup"><span data-stu-id="1c98b-143">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="1c98b-144">`Application.suspendScreenUpdatingUntilNextSync()` interrompt les mises à jour visuelles vers Excel tant que le complément n’appelle pas `context.sync()`, ou tant que `Excel.run` ne se termine pas (appelant implicitement `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="1c98b-144">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="1c98b-145">N’oubliez pas qu'Excel n’affiche aucun signe d’activité jusqu'à la synchronisation suivante. Votre complément doit donner des conseils aux utilisateurs pour les préparer à ce délai ou fournir une barre d’état pour démontrer l’activité.</span><span class="sxs-lookup"><span data-stu-id="1c98b-145">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="1c98b-146">Ne pas appeler `suspendScreenUpdatingUntilNextSync` de manière répétée (comme dans une boucle).</span><span class="sxs-lookup"><span data-stu-id="1c98b-146">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="1c98b-147">Les appels répétés entraînent le scintillement de la fenêtre Excel.</span><span class="sxs-lookup"><span data-stu-id="1c98b-147">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="1c98b-148">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="1c98b-148">Enable and disable events</span></span>

<span data-ttu-id="1c98b-149">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="1c98b-149">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="1c98b-150">Un exemple de code montrant comment activer et désactiver les événements dans l’article[manipuler les événements](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="1c98b-150">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="1c98b-151">Importation de données dans des tableaux</span><span class="sxs-lookup"><span data-stu-id="1c98b-151">Importing data into tables</span></span>

<span data-ttu-id="1c98b-152">Lorsque vous tentez d’importer une quantité considérable de données directement dans un objet[tableau](/javascript/api/excel/excel.table) (par exemple, à l’aide de `TableRowCollection.add()`), vous risquez de rencontrer une dégradation des performances.</span><span class="sxs-lookup"><span data-stu-id="1c98b-152">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="1c98b-153">Si vous essayez d’ajouter un nouveau tableau, vous devez remplir les données d’abord en définissant `range.values`, puis appelez `worksheet.tables.add()` pour créer un tableau sur la plage.</span><span class="sxs-lookup"><span data-stu-id="1c98b-153">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="1c98b-154">Si vous essayez d’écrire des données dans un tableau existant, écrivez les données dans un objet plage via `table.getDataBodyRange()`, et le tableau s’agrandit automatiquement.</span><span class="sxs-lookup"><span data-stu-id="1c98b-154">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="1c98b-155">Voici un exemple de cette approche :</span><span class="sxs-lookup"><span data-stu-id="1c98b-155">Here is an example of this approach:</span></span>

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first 
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
> <span data-ttu-id="1c98b-156">Vous pouvez convertir un objet de Tableau en objet de Plage à l’aide de la méthode[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="1c98b-156">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="1c98b-157">Annuler le suivi des plages inutiles</span><span class="sxs-lookup"><span data-stu-id="1c98b-157">Untrack unneeded ranges</span></span>

<span data-ttu-id="1c98b-158">La couche JavaScript crée des objets proxy pour votre complément pour interagir avec le classeur Excel et les sous-jacentes.</span><span class="sxs-lookup"><span data-stu-id="1c98b-158">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="1c98b-159">Ces objets sont conservés en mémoire jusqu'à `context.sync()` soit appelé.</span><span class="sxs-lookup"><span data-stu-id="1c98b-159">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="1c98b-160">Les opérations par lots volumineux peuvent générer un grand nombre d’objets proxy qui sont uniquement utiles une fois pour le complément et peuvent être publiés à partir de la mémoire avant l’exécution du lot.</span><span class="sxs-lookup"><span data-stu-id="1c98b-160">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="1c98b-161">La méthode [Range.untrack()](/javascript/api/excel/excel.range#untrack--) libère un objet plage Excel à partir de la mémoire.</span><span class="sxs-lookup"><span data-stu-id="1c98b-161">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="1c98b-162">Appeler cette méthode une fois que votre complément a terminé avec la plage doit créer une amélioration notable des performances lors de l’utilisation d’un grand nombre d’objets de plage.</span><span class="sxs-lookup"><span data-stu-id="1c98b-162">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span>

> [!NOTE]
> <span data-ttu-id="1c98b-163">`Range.untrack()` est un raccourci pour [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span><span class="sxs-lookup"><span data-stu-id="1c98b-163">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="1c98b-164">N’importe quel objet proxy peut être non suivi en le supprimant de la liste d’objets suivis dans le contexte.</span><span class="sxs-lookup"><span data-stu-id="1c98b-164">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="1c98b-165">En règle générale, les objets Plage sont les seuls objets Excel utilisés dans une quantité suffisante pour justifier le non suivi.</span><span class="sxs-lookup"><span data-stu-id="1c98b-165">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="1c98b-166">L’exemple de code suivant remplit une plage sélectionnée avec des données, une cellule à la fois.</span><span class="sxs-lookup"><span data-stu-id="1c98b-166">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="1c98b-167">Une fois que la valeur est ajoutée à la cellule, la plage représentant cette cellule est non suivie.</span><span class="sxs-lookup"><span data-stu-id="1c98b-167">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="1c98b-168">Exécuter tout d’abord ce code avec une plage sélectionnée de 10 000 à 20 000 cellules, avec la `cell.untrack()` ligne et puis sans.</span><span class="sxs-lookup"><span data-stu-id="1c98b-168">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="1c98b-169">Vous devez remarquer que le code est exécuté plus rapidement avec la `cell.untrack()` ligne que sans elle.</span><span class="sxs-lookup"><span data-stu-id="1c98b-169">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="1c98b-170">Vous pouvez également remarquer un temps de réponse plus rapide par la suite, étant donné que l’étape de nettoyage prend moins de temps.</span><span class="sxs-lookup"><span data-stu-id="1c98b-170">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // call untrack() to release the range from memory
            cell.untrack();
        }
    }

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="1c98b-171">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1c98b-171">See also</span></span>

- [<span data-ttu-id="1c98b-172">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="1c98b-172">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="1c98b-173">Concepts avancés de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="1c98b-173">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="1c98b-174">Limites des ressources et optimisation des performances pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="1c98b-174">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- [<span data-ttu-id="1c98b-175">Objet de fonctions de feuille de calcul (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="1c98b-175">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
