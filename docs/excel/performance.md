---
title: Optimisation des performances de l'API JavaScript d'Excel
description: Optimiser les performances ? l'aide de l'API JavaScript d'Excel
ms.date: 03/28/2018
ms.openlocfilehash: dabbb69f8dee0df782a265edcfdfb1c89894e915
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="45953-103">Optimisation des performances ? l'aide de l'API JavaScript d'Excel</span><span class="sxs-lookup"><span data-stu-id="45953-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="45953-104">Il y a plusieurs mani?res d'effectuer des t?ches courantes avec l'API JavaScript d'Excel.</span><span class="sxs-lookup"><span data-stu-id="45953-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="45953-105">Vous trouverez des diff?rences de performances significatives entre les diverses approches.</span><span class="sxs-lookup"><span data-stu-id="45953-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="45953-106">Cet article fournit de l'aide et des exemples de code pour vous montrer comment effectuer efficacement des t?ches courantes en utilisant l'API JavaScript d'Excel.</span><span class="sxs-lookup"><span data-stu-id="45953-106">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="45953-107">R?duisez le nombre d'appels ? sync()</span><span class="sxs-lookup"><span data-stu-id="45953-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="45953-108">Dans l'API JavaScript d'Excel, ```sync()``` est la seule op?ration asynchrone, et elle peut ?tre lente dans certaines circonstances, en particulier pour Excel Online.</span><span class="sxs-lookup"><span data-stu-id="45953-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="45953-109">Pour optimiser les performances, r?duisez le nombre d'appels ? ```sync()``` en mettant en file d'attente autant de changements que possible avant de l'appeler.</span><span class="sxs-lookup"><span data-stu-id="45953-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="45953-110">Voir [Concepts de base - sync()](excel-add-ins-core-concepts.md#sync) pour des exemples de code qui suivent cette pratique.</span><span class="sxs-lookup"><span data-stu-id="45953-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="45953-111">R?duisez le nombre d'objets proxy cr??s</span><span class="sxs-lookup"><span data-stu-id="45953-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="45953-112">?vitez de cr?er r?p?titivement le m?me objet proxy.</span><span class="sxs-lookup"><span data-stu-id="45953-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="45953-113">A la place, si vous avez besoin du m?me objet proxy pour plus d'une op?ration, cr?ez-le une fois et affectez-le ? une variable, puis utilisez cette variable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="45953-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="45953-114">Ne chargez que les propri?t?s n?cessaires</span><span class="sxs-lookup"><span data-stu-id="45953-114">Load necessary properties only</span></span>

<span data-ttu-id="45953-115">Dans l'API JavaScript d'Excel, vous devez charger explicitement les propri?t?s d'un objet proxy.</span><span class="sxs-lookup"><span data-stu-id="45953-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="45953-116">Bien que vous puissiez charger toutes les propri?t?s en une fois avec un appel vide ? ```load()```, cette approche peut avoir un surco?t significatif en termes de performances.</span><span class="sxs-lookup"><span data-stu-id="45953-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="45953-117">A la place, nous vous sugg?rons de ne charger que les propri?t?s n?cessaires, en particulier pour ceux des objets qui ont un nombre important de propri?t?s.</span><span class="sxs-lookup"><span data-stu-id="45953-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="45953-118">Par exemple, si vous ne souhaitez relire que la propri?t? **address** d?un objet plage, indiquez seulement cette propri?t? lorsque vous appelez la m?thode **load()** :</span><span class="sxs-lookup"><span data-stu-id="45953-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="45953-119">Vous pouvez appeler la m?thode **load()** de l?une quelconque des fa?ons suivantes :</span><span class="sxs-lookup"><span data-stu-id="45953-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="45953-120">_Syntaxe :_</span><span class="sxs-lookup"><span data-stu-id="45953-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="45953-121">_O? :_</span><span class="sxs-lookup"><span data-stu-id="45953-121">_Where:_</span></span>
 
* <span data-ttu-id="45953-122">`properties` est la liste des propri?t?s ? charger, sp?cifi?e comme des cha?nes d?limit?es par des virgules ou comme un tableau de noms.</span><span class="sxs-lookup"><span data-stu-id="45953-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="45953-123">Pour plus d?informations, voir les m?thodes **load()** d?finies pour les objets dans la [R?f?rence de l?API JavaScript d'Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="45953-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="45953-p106">`loadOption` sp?cifie un objet qui d?crit les options selection, expansion, top et skip. Voir les [options](https://dev.office.com/reference/add-ins/excel/loadoption) de chargement d?objet pour les d?tails.</span><span class="sxs-lookup"><span data-stu-id="45953-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://dev.office.com/reference/add-ins/excel/loadoption) for details.</span></span>

<span data-ttu-id="45953-126">SVP, soyez conscient que certaines des "propri?t?s" dans un objet peuvent avoir le m?me nom qu'un autre objet.</span><span class="sxs-lookup"><span data-stu-id="45953-126">Please be aware that some of the ?properties? under an object may have the same name as another object.</span></span> <span data-ttu-id="45953-127">Par exemple, `format` est une propri?t? dans l'objet plage, mais `format` lui-m?me est un objet aussi.</span><span class="sxs-lookup"><span data-stu-id="45953-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="45953-128">Donc, si vous faites un appel tel que `range.load("format")`, c'est ?quivalent ? `range.format.load()`, qui est un appel vide ? load() qui peut engendrer des probl?mes de performances comme r?sum? pr?c?demment.</span><span class="sxs-lookup"><span data-stu-id="45953-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="45953-129">Pour ?viter cela, votre code ne devrait charger que les "n?uds feuilles" dans une arborescence d'objets.</span><span class="sxs-lookup"><span data-stu-id="45953-129">To avoid this, your code should only load the ?leaf nodes? in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="45953-130">Suspendre le calcul temporairement</span><span class="sxs-lookup"><span data-stu-id="45953-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="45953-131">Si vous essayez d'effectuer une op?ration sur un grand nombre de cellules (par exemple, en d?finissant la valeur d'un objet plage tr?s volumineux) et que cela ne vous d?range pas de suspendre temporairement le calcul dans Excel jusqu'? ce que votre op?ration se termine, nous vous recommandons de suspendre le calcul jusqu'? ce que le prochain ```context.sync()``` soit appel?.</span><span class="sxs-lookup"><span data-stu-id="45953-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="45953-132">Voir la documentation de r?f?rence de l'[Objet Application](https://dev.office.com/reference/add-ins/excel/application) pour des informations sur la fa?on d'utiliser l'```suspendApiCalculationUntilNextSync()``` API pour suspendre et r?activer les calculs d'une mani?re tr?s pratique.</span><span class="sxs-lookup"><span data-stu-id="45953-132">See [Application Object](https://dev.office.com/reference/add-ins/excel/application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="45953-133">Le code suivant montre comment suspendre le calcul temporairement :</span><span class="sxs-lookup"><span data-stu-id="45953-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="45953-134">Mettre ? jour toutes les cellules d?une plage</span><span class="sxs-lookup"><span data-stu-id="45953-134">Update all cells in a range</span></span> 

<span data-ttu-id="45953-135">Lorsque vous devez mettre ? jour toutes les cellules dans une plage avec la m?me valeur ou propri?t?, il peut ?tre lent de le faire via un tableau bidimensionnel qui indique r?p?titivement la m?me valeur, car cette approche n?cessite qu'Excel parcoure toutes les cellules de la plage pour les d?finir individuellement.</span><span class="sxs-lookup"><span data-stu-id="45953-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="45953-136">Excel a un moyen plus efficace pour mettre ? jour toutes les cellules dans une plage avec la m?me valeur ou propri?t?.</span><span class="sxs-lookup"><span data-stu-id="45953-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="45953-137">Si vous devez appliquer la m?me valeur, le m?me format num?rique ou la m?me formule ? une plage de cellules, il est plus efficace de sp?cifier une seule valeur au lieu d'un tableau de valeurs.</span><span class="sxs-lookup"><span data-stu-id="45953-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="45953-138">Proc?der ainsi am?liorera significativement les performances.</span><span class="sxs-lookup"><span data-stu-id="45953-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="45953-139">Pour un exemple de code qui montre cette approche en action, voir [Concepts de base - Mettre ? jour toutes les cellules d'une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="45953-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="45953-140">Un sc?nario courant dans lequel vous pouvez appliquer cette approche est la d?finition de formats num?riques diff?rents pour des colonnes diff?rentes dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="45953-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="45953-141">Dans ce cas, vous pouvez simplement parcourir les colonnes et d?finir le format num?rique pour chaque colonne avec une seule valeur.</span><span class="sxs-lookup"><span data-stu-id="45953-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="45953-142">Manipuler chaque colonne comme une plage, comme indiqu? dans l'exemple de code [Mettre ? jour toutes les cellules dans une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="45953-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="45953-143">Si vous utilisez TypeScript, vous remarquerez une erreur de compilation indiquant qu'une valeur unique ne peut pas ?tre affect?e ? un tableau 2D.</span><span class="sxs-lookup"><span data-stu-id="45953-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="45953-144">C'est in?vitable du fait que les valeurs *sont* un tableau 2D lors de la r?cup?ration des propri?t?s, et que TypeScript n'autorise pas des types diff?rents pour un setter et un getter.</span><span class="sxs-lookup"><span data-stu-id="45953-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="45953-145">Cependant, un contournement simple consiste ? d?finir les valeurs avec un suffixe `as any`, par exemple, `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="45953-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="45953-146">Importation de donn?es dans des tables</span><span class="sxs-lookup"><span data-stu-id="45953-146">Importing data into tables</span></span>

<span data-ttu-id="45953-147">Lorsque vous essayez d'importer un tr?s grand volume de donn?es directement dans un objet[Table](https://dev.office.com/reference/add-ins/excel/table) (par exemple, en utilisant `TableRowCollection.add()`), vous risquez de subir une performance lente.</span><span class="sxs-lookup"><span data-stu-id="45953-147">When trying to import a huge amount of data directly into a [Table](https://dev.office.com/reference/add-ins/excel/table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="45953-148">Si vous essayez d'ajouter une nouvelle table, vous devriez d'abord remplir les donn?es en d?finissant `range.values`, puis appeler alors `worksheet.tables.add()` pour cr?er une table sur la plage.</span><span class="sxs-lookup"><span data-stu-id="45953-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="45953-149">Si vous essayez d'?crire des donn?es dans une table existante, ?crivez les donn?es dans un objet plage via `table.getDataBodyRange()`, et la table s'agrandira automatiquement.</span><span class="sxs-lookup"><span data-stu-id="45953-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="45953-150">Voici un exemple de cette approche :</span><span class="sxs-lookup"><span data-stu-id="45953-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="45953-151">Vous pouvez ais?ment convertir un objet Table en objet Range en utilisant la m?thode [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange).</span><span class="sxs-lookup"><span data-stu-id="45953-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="45953-152">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="45953-152">See also</span></span>

- [<span data-ttu-id="45953-153">Concepts de base de l?API JavaScript d'Excel</span><span class="sxs-lookup"><span data-stu-id="45953-153">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="45953-154">Concepts avanc?s de l?API JavaScript d'Excel</span><span class="sxs-lookup"><span data-stu-id="45953-154">Excel JavaScript API advanced concepts</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="45953-155">Sp?cification ouverte de l?API JavaScript d'Excel</span><span class="sxs-lookup"><span data-stu-id="45953-155">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="45953-156">Objet de fonctions de feuille de calcul (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="45953-156">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/functions)
