---
title: Optimisation des performances API JavaScript Excel
description: Optimisation des performances à l’aide de l’API JavaScript d’Excel
ms.date: 11/29/2018
ms.openlocfilehash: fb0f81b79d2eac847a91a7b2a4fab92362330a10
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156578"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="6f97e-103">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="6f97e-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="6f97e-104">Il existe plusieurs façons d’effectuer des tâches courantes avec l’API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="6f97e-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="6f97e-105">Vous trouverez des différences de performances significatives entre les différentes approches.</span><span class="sxs-lookup"><span data-stu-id="6f97e-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="6f97e-106">Cet article fournit des instructions et exemples de code pour vous montrer comment effectuer des tâches courantes efficacement à l’aide des API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="6f97e-106">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="6f97e-107">Limitez le nombre d’appels sync()</span><span class="sxs-lookup"><span data-stu-id="6f97e-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="6f97e-108">Dans l’API JavaScript Excel, ```sync()``` est la seule opération asynchrone et elle peut être lente dans certaines circonstances, en particulier pour Excel Online.</span><span class="sxs-lookup"><span data-stu-id="6f97e-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="6f97e-109">Pour optimiser les performances, vous devez limiter le nombre de fois que vous appelez ```sync()``` et mettre en file d’attente autant de modifications que possible avant d’appeler.</span><span class="sxs-lookup"><span data-stu-id="6f97e-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="6f97e-110">Voir [Concepts principaux - sync()](excel-add-ins-core-concepts.md#sync) pour des exemples de code qui suivent cette pratique.</span><span class="sxs-lookup"><span data-stu-id="6f97e-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="6f97e-111">Réduire le nombre d’objets proxy créés</span><span class="sxs-lookup"><span data-stu-id="6f97e-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="6f97e-112">Éviter de créer le même objet proxy à plusieurs reprises.</span><span class="sxs-lookup"><span data-stu-id="6f97e-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="6f97e-113">Au lieu de cela, si vous avez besoin du même objet proxy pour plus d’une opération, créez-le une seule fois et affectez-le à une variable, puis utilisez cette variable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="6f97e-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```javascript
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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="6f97e-114">Charger les propriétés nécessaires uniquement</span><span class="sxs-lookup"><span data-stu-id="6f97e-114">Load necessary properties only</span></span>

<span data-ttu-id="6f97e-115">Dans l’API JavaScript Excel, vous devez charger explicitement les propriétés d’un objet proxy.</span><span class="sxs-lookup"><span data-stu-id="6f97e-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="6f97e-116">Bien que vous soyez en mesure de charger les propriétés en une fois avec un appel vide```load()```, cette approche peut causer une surcharge significative des performances.</span><span class="sxs-lookup"><span data-stu-id="6f97e-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="6f97e-117">Au lieu de cela, nous vous conseillons de charger uniquement les propriétés nécessaires, en particulier pour ces objets qui ont un grand nombre de propriétés.</span><span class="sxs-lookup"><span data-stu-id="6f97e-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="6f97e-118">Par exemple, si vous souhaitez uniquement lire la propriété **adresse** d’un objet de la plage, spécifiez uniquement cette propriété lorsque vous appelez la méthode **load()**  :</span><span class="sxs-lookup"><span data-stu-id="6f97e-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="6f97e-119">Vous pouvez appeler la méthode **load()** de l’une des façons suivantes :</span><span class="sxs-lookup"><span data-stu-id="6f97e-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="6f97e-120">_Syntaxe :_</span><span class="sxs-lookup"><span data-stu-id="6f97e-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="6f97e-121">_Où :_</span><span class="sxs-lookup"><span data-stu-id="6f97e-121">_Where:_</span></span>
 
* <span data-ttu-id="6f97e-122">`properties` est la liste des propriétés à charger, fournie sous forme de chaînes séparées par des virgules ou de tableau de noms.</span><span class="sxs-lookup"><span data-stu-id="6f97e-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="6f97e-123">Pour plus d’informations, reportez-vous aux méthodes **load()** définies pour les objets dans la rubrique [Référence de l’API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="6f97e-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="6f97e-p106">`loadOption` spécifie un objet qui décrit les options select, expand, top et skip. Pour plus d’informations, reportez-vous aux [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) de chargement d’objet.</span><span class="sxs-lookup"><span data-stu-id="6f97e-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="6f97e-126">N’oubliez pas que certaines des « propriétés » sous un objet peuvent avoir le même nom qu’un autre objet.</span><span class="sxs-lookup"><span data-stu-id="6f97e-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="6f97e-127">Par exemple, `format` est une propriété sous plage d’objet, mais `format` lui-même est également un objet.</span><span class="sxs-lookup"><span data-stu-id="6f97e-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="6f97e-128">Par conséquent, si vous passez un appel comme `range.load("format")`, cela équivaut à `range.format.load()`, c'est-à-dire, un appel load() vide pouvant entraîner des problèmes de performances comme indiqué précédemment.</span><span class="sxs-lookup"><span data-stu-id="6f97e-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="6f97e-129">Pour éviter cela, votre code devrait charger uniquement les nœuds « terminaux » dans une arborescence d’objets.</span><span class="sxs-lookup"><span data-stu-id="6f97e-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="6f97e-130">Suspendre temporairement les calculs</span><span class="sxs-lookup"><span data-stu-id="6f97e-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="6f97e-131">Si vous essayez d’effectuer une opération sur un grand nombre de cellules (par exemple, la définition de la valeur d’un objet plage énorme) et que vous n’avez rien contre la suspension de calcul dans Excel temporairement le temps que votre opération se termine, nous vous recommandons de suspendre le calcul jusqu’à ce que le prochain ```context.sync()``` soit appelé.</span><span class="sxs-lookup"><span data-stu-id="6f97e-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="6f97e-132">Voir la documentation[objet Application](https://docs.microsoft.com/javascript/api/excel/excel.application) pour plus d’informations sur l’utilisation de l’```suspendApiCalculationUntilNextSync()``` API pour suspendre et réactiver les calculs de manière très pratique.</span><span class="sxs-lookup"><span data-stu-id="6f97e-132">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="6f97e-133">Le code suivant montre comment suspendre temporairement le calcul :</span><span class="sxs-lookup"><span data-stu-id="6f97e-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

    // Suspending recalc
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with supend recalc
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="6f97e-134">Mettre à jour toutes les cellules d’une plage</span><span class="sxs-lookup"><span data-stu-id="6f97e-134">Update all cells in a range</span></span> 

<span data-ttu-id="6f97e-135">Lorsque vous devez mettre à jour toutes les cellules dans une plage avec la même valeur ou propriété, il peut être lent de le faire via une matrice 2 dimensions indiquant à plusieurs reprises la même valeur étant donné que cette approche nécessite qu’Excel le répète sur toutes les cellules dans la plage pour définir chacune séparément.</span><span class="sxs-lookup"><span data-stu-id="6f97e-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="6f97e-136">Excel propose une méthode plus efficace pour mettre à jour toutes les cellules dans une plage avec la même valeur ou propriété.</span><span class="sxs-lookup"><span data-stu-id="6f97e-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="6f97e-137">Si vous voulez appliquer la même valeur, le même format de nombre ou la même formule à une plage de cellules, il est plus efficace de spécifier une valeur unique au lieu d’une matrice de valeurs.</span><span class="sxs-lookup"><span data-stu-id="6f97e-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="6f97e-138">Cette opération va améliorer sensiblement les performances.</span><span class="sxs-lookup"><span data-stu-id="6f97e-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="6f97e-139">Pour voir un exemple de code indiquant cette approche en action, [principaux concepts - mettre à jour toutes les cellules dans une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="6f97e-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="6f97e-140">Un scénario classique où vous pouvez appliquer cette approche est lors de la définition de différents formats numériques différents sur différentes colonnes dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="6f97e-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="6f97e-141">Dans ce cas, vous pouvez simplement itérer dans les colonnes et définir le format de nombre dans chaque colonne avec une valeur unique.</span><span class="sxs-lookup"><span data-stu-id="6f97e-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="6f97e-142">Traiter chaque colonne comme une plage, comme illustré dans l’exemple de code[mettre à jour toutes les cellules dans une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="6f97e-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="6f97e-143">Si vous utilisez TypeScript, vous remarquerez une erreur de compilation indiquant qu’une seule valeur ne peut pas être définie à une matrice 2D.</span><span class="sxs-lookup"><span data-stu-id="6f97e-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="6f97e-144">Ceci est inévitable puisque les valeurs *sont* un tableau 2D qui extrait les propriétés et TypeScript n’autorise pas de types différents pour configurer et récolter.</span><span class="sxs-lookup"><span data-stu-id="6f97e-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="6f97e-145">Toutefois, une solution de contournement simple consiste à définir les valeurs avec un `as any` suffixe, par exemple, `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="6f97e-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="6f97e-146">Importation de données dans des tableaux</span><span class="sxs-lookup"><span data-stu-id="6f97e-146">Importing data into tables</span></span>

<span data-ttu-id="6f97e-147">Lorsque vous tentez d’importer une quantité considérable de données directement dans un objet[tableau](https://docs.microsoft.com/javascript/api/excel/excel.table) (par exemple, à l’aide de `TableRowCollection.add()`), vous risquez de rencontrer une dégradation des performances.</span><span class="sxs-lookup"><span data-stu-id="6f97e-147">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="6f97e-148">Si vous essayez d’ajouter un nouveau tableau, vous devez remplir les données d’abord en définissant `range.values`, puis appelez `worksheet.tables.add()` pour créer un tableau sur la plage.</span><span class="sxs-lookup"><span data-stu-id="6f97e-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="6f97e-149">Si vous essayez d’écrire des données dans un tableau existant, écrivez les données dans un objet plage via `table.getDataBodyRange()`, et le tableau s’agrandit automatiquement.</span><span class="sxs-lookup"><span data-stu-id="6f97e-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="6f97e-150">Voici un exemple de cette approche :</span><span class="sxs-lookup"><span data-stu-id="6f97e-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="6f97e-151">Vous pouvez convertir un objet de Tableau en objet de Plage à l’aide de la méthode[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="6f97e-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="6f97e-152">Annuler le suivi des plages inutiles</span><span class="sxs-lookup"><span data-stu-id="6f97e-152">Untrack unneeded ranges</span></span>

<span data-ttu-id="6f97e-153">La couche JavaScript crée des objets proxy pour votre complément pour interagir avec le classeur Excel et les sous-jacentes.</span><span class="sxs-lookup"><span data-stu-id="6f97e-153">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="6f97e-154">Ces objets sont conservés en mémoire jusqu'à `context.sync()` soit appelé.</span><span class="sxs-lookup"><span data-stu-id="6f97e-154">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="6f97e-155">Les opérations par lots volumineux peuvent générer un grand nombre d’objets proxy qui sont uniquement utiles une fois pour le complément et peuvent être publiés à partir de la mémoire avant l’exécution du lot.</span><span class="sxs-lookup"><span data-stu-id="6f97e-155">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="6f97e-156">La méthode [Range.untrack()](/javascript/api/excel/excel.range#untrack--) libère un objet plage Excel à partir de la mémoire.</span><span class="sxs-lookup"><span data-stu-id="6f97e-156">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="6f97e-157">Appeler cette méthode une fois que votre complément a terminé avec la plage doit créer une amélioration notable des performances lors de l’utilisation d’un grand nombre d’objets de plage.</span><span class="sxs-lookup"><span data-stu-id="6f97e-157">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span> 

> [!NOTE]
> <span data-ttu-id="6f97e-158">`Range.untrack()` est un raccourci pour [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span><span class="sxs-lookup"><span data-stu-id="6f97e-158">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="6f97e-159">N’importe quel objet proxy peut être non suivi en le supprimant de la liste d’objets suivis dans le contexte.</span><span class="sxs-lookup"><span data-stu-id="6f97e-159">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="6f97e-160">En règle générale, les objets Plage sont les seuls objets Excel utilisés dans une quantité suffisante pour justifier le non suivi.</span><span class="sxs-lookup"><span data-stu-id="6f97e-160">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="6f97e-161">L’exemple de code suivant remplit une plage sélectionnée avec des données, une cellule à la fois.</span><span class="sxs-lookup"><span data-stu-id="6f97e-161">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="6f97e-162">Une fois que la valeur est ajoutée à la cellule, la plage représentant cette cellule est non suivie.</span><span class="sxs-lookup"><span data-stu-id="6f97e-162">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="6f97e-163">Exécuter tout d’abord ce code avec une plage sélectionnée de 10 000 à 20 000 cellules, avec la `cell.untrack()` ligne et puis sans.</span><span class="sxs-lookup"><span data-stu-id="6f97e-163">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="6f97e-164">Vous devez remarquer que le code est exécuté plus rapidement avec la `cell.untrack()` ligne que sans elle.</span><span class="sxs-lookup"><span data-stu-id="6f97e-164">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="6f97e-165">Vous pouvez également remarquer un temps de réponse plus rapide par la suite, étant donné que l’étape de nettoyage prend moins de temps.</span><span class="sxs-lookup"><span data-stu-id="6f97e-165">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

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

## <a name="enable-and-disable-events"></a><span data-ttu-id="6f97e-166">Activation et désactivation d’événements</span><span class="sxs-lookup"><span data-stu-id="6f97e-166">Enable and disable events</span></span>

<span data-ttu-id="6f97e-167">La performance d’un complément peut être améliorée en désactivant les événements.</span><span class="sxs-lookup"><span data-stu-id="6f97e-167">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="6f97e-168">Un exemple de code montrant comment activer et désactiver les événements dans l’article[manipuler les événements](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="6f97e-168">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="6f97e-169">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="6f97e-169">See also</span></span>

- [<span data-ttu-id="6f97e-170">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6f97e-170">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6f97e-171">Concepts avancés de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="6f97e-171">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="6f97e-172">Spécification d’ouverture d’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="6f97e-172">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="6f97e-173">Objet de fonctions de feuille de calcul (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="6f97e-173">Worksheet Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)
