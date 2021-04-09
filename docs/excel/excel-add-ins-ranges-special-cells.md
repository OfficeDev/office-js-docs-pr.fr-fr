---
title: Rechercher des cellules spéciales dans une plage à l’aide de l’API JavaScript pour Excel
description: Découvrez comment utiliser l’API JavaScript excel pour rechercher des cellules spéciales, telles que des cellules avec des formules, des erreurs ou des nombres.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6504873bcd8ab50bd4c03fe4f54b71d0bd920c5b
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652824"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="a7a0d-103">Rechercher des cellules spéciales dans une plage à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="a7a0d-103">Find special cells within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="a7a0d-104">Cet article fournit des exemples de code qui recherchent des cellules spéciales dans une plage à l’aide de l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-104">This article provides code samples that find special cells within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="a7a0d-105">Pour obtenir la liste complète des propriétés et des méthodes que l’objet prend en charge, voir `Range` la classe [Excel.Range.](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="a7a0d-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="find-ranges-with-special-cells"></a><span data-ttu-id="a7a0d-106">Rechercher des plages avec des cellules spéciales</span><span class="sxs-lookup"><span data-stu-id="a7a0d-106">Find ranges with special cells</span></span>

<span data-ttu-id="a7a0d-107">Les méthodes [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) et [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) recherchent des plages en fonction des caractéristiques de leurs cellules et des types de valeurs de leurs cellules.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-107">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="a7a0d-108">Ces deux méthodes renvoient à des`RangeAreas`objets.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-108">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="a7a0d-109">Voici les signatures des méthodes à partir des types de fichiers de données TypeScript:</span><span class="sxs-lookup"><span data-stu-id="a7a0d-109">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="a7a0d-110">L’exemple de code suivant utilise `getSpecialCells` la méthode pour rechercher toutes les cellules avec des formules.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-110">The following code sample uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="a7a0d-111">Tenez compte du code suivant:</span><span class="sxs-lookup"><span data-stu-id="a7a0d-111">About this code, note:</span></span>

- <span data-ttu-id="a7a0d-112">Cela limite la partie de la feuille qui nécessite d’être recherchée en appelant d’abord`Worksheet.getUsedRange`et en appelant`getSpecialCells`uniquement pour cette plage.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-112">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="a7a0d-113">La`getSpecialCells`méthode renvoie un`RangeAreas`objet, toutes les cellules alors dotées de formules seront colorées en rose même si elles ne sont pas adjacentes.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-113">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a7a0d-114">Si aucune cellule avec la caractéristique ciblée n’existe dans la plage `getSpecialCells` lève une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-114">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="a7a0d-115">Cela dévie le flux de contrôle vers un(e)`catch`bloc/méthode, s’il en existe.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-115">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="a7a0d-116">S’il n’y a `catch` pas de bloc, l’erreur arrête la méthode.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-116">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="a7a0d-117">Si vous attendez que des cellules avec la caractéristique ciblée existent toujours, vous souhaiterez probablement que votre code  lève une erreur si ces cellules ne sont pas là.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-117">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="a7a0d-118">Mais dans les scénarios où les cellules ne correspondent pas; votre code doit vérifier cette possibilité et le gérer gracieusement sans émettre d’erreur.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-118">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="a7a0d-119">Vous pouvez obtenir ce comportement avec la `getSpecialCellsOrNullObject`méthode et sa propriété renvoyée`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-119">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="a7a0d-120">L’exemple de code suivant utilise ce modèle.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-120">The following code sample uses this pattern.</span></span> <span data-ttu-id="a7a0d-121">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="a7a0d-121">About this code, note:</span></span>

- <span data-ttu-id="a7a0d-122">La méthode renvoie toujours un objet proxy, donc elle `getSpecialCellsOrNullObject` n’est jamais dans le sens `null` JavaScript ordinaire.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-122">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it's never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="a7a0d-123">Mais si les cellules non correspondantes sont introuvables, la propriété`isNullObject` de l’objet est établi à`true`.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-123">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="a7a0d-124">Il appelle`context.sync`*avant* de tester la propriété`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-124">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="a7a0d-125">Il s’agit d’une condition avec toutes les méthodes et propriétés`*OrNullObject`, car vous devez toujours télécharger et synchroniser une propriété afin de le lire. </span><span class="sxs-lookup"><span data-stu-id="a7a0d-125">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="a7a0d-126">Toutefois, il n’est pas nécessaire de *charger explicitement* la `isNullObject` propriété.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-126">However, it's not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="a7a0d-127">Il est automatiquement chargé par le même `context.sync` s’il `load` n’est pas appelé sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-127">It's automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="a7a0d-128">Pour plus d’informations, [ \* voir méthodes et propriétés OrNullObject.](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)</span><span class="sxs-lookup"><span data-stu-id="a7a0d-128">For more information, see [\*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).</span></span>
- <span data-ttu-id="a7a0d-129">Vous pouvez tester ce code en sélectionnant d’abord une plage qui n’a pas de cellules de formule et en l’exécutant.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-129">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="a7a0d-130">Puis sélectionnez une plage qui dispose au moins d’une cellule dotée d’une formule et en l’exécutant à nouveau.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-130">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

<span data-ttu-id="a7a0d-131">Par souci de simplicité, tous les autres exemples de code de cet article utilisent la `getSpecialCells` méthode au lieu de  `getSpecialCellsOrNullObject` .</span><span class="sxs-lookup"><span data-stu-id="a7a0d-131">For simplicity, all other code samples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

## <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="a7a0d-132">Réduisez les cellules cibles avec les types de valeur de cellule</span><span class="sxs-lookup"><span data-stu-id="a7a0d-132">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="a7a0d-133">Les méthodes`Range.getSpecialCells()` et `Range.getSpecialCellsOrNullObject()`acceptent un deuxième paramètre facultatif utilisé pour affiner davantage les cellules ciblées.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-133">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="a7a0d-134">Ce deuxième paramètre est un`Excel.SpecialCellValueType` que vous utilisez afin de spécifier que vous souhaitez uniquement les cellules qui contiennent certains types de valeurs.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-134">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="a7a0d-135">Le paramètre `Excel.SpecialCellValueType` peut uniquement être utilisé si le paramètre `Excel.SpecialCellType` est défini sur `Excel.SpecialCellType.formulas`ou `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-135">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="a7a0d-136">Test d’un type de valeur de cellule unique</span><span class="sxs-lookup"><span data-stu-id="a7a0d-136">Test for a single cell value type</span></span>

<span data-ttu-id="a7a0d-137">Le `Excel.SpecialCellValueType` enum dispose de ces quatre types de base (outre les autres valeurs combinées décrites plus loin dans cette section):</span><span class="sxs-lookup"><span data-stu-id="a7a0d-137">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="a7a0d-138">`Excel.SpecialCellValueType.logical` (ce qui signifie booléen)</span><span class="sxs-lookup"><span data-stu-id="a7a0d-138">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="a7a0d-139">L’exemple de code suivant recherche des cellules spéciales qui sont des constantes numériques et colore ces cellules en rose.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-139">The following code sample finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="a7a0d-140">Tenez compte du code suivant :</span><span class="sxs-lookup"><span data-stu-id="a7a0d-140">About this code, note:</span></span>

- <span data-ttu-id="a7a0d-141">Il met uniquement en évidence les cellules qui ont une valeur de nombre littérale.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-141">It only highlights cells that have a literal number value.</span></span> <span data-ttu-id="a7a0d-142">Il ne surligne pas les cellules qui ont une formule (même si le résultat est un nombre) ou des cellules booléles, de texte ou d’état d’erreur.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-142">It won't highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="a7a0d-143">Pour tester le code, assurez-vous que la feuille de calcul dispose de certaines cellules avec des valeurs de nombre littérales, certaines avec d’autres sortes de valeurs littérales, et certaines avec des formules.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-143">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="a7a0d-144">Test d’un type de valeur de cellule multiple</span><span class="sxs-lookup"><span data-stu-id="a7a0d-144">Test for multiple cell value types</span></span>

<span data-ttu-id="a7a0d-145">Parfois, vous avez besoin d’exécuter plus d’un type de valeur de cellule, tel que toutes les cellules à valeur de texte et à valeur booléen (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="a7a0d-145">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="a7a0d-146">Le `Excel.SpecialCellValueType` enum comporte des valeurs avec les types combinés.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-146">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="a7a0d-147">Par exemple,`Excel.SpecialCellValueType.logicalText`cible toutes les cellules à valeur texte et booléen.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-147">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="a7a0d-148">`Excel.SpecialCellValueType.all` est la valeur par défaut, ce qui ne limite pas les types de valeur de cellule renvoyés.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-148">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="a7a0d-149">L’exemple de code suivant colore toutes les cellules avec des formules qui produisent une valeur de nombre ou booléen.</span><span class="sxs-lookup"><span data-stu-id="a7a0d-149">The following code sample colors all cells with formulas that produce number or boolean value.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="see-also"></a><span data-ttu-id="a7a0d-150">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a7a0d-150">See also</span></span>

- [<span data-ttu-id="a7a0d-151">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="a7a0d-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="a7a0d-152">Utiliser des cellules à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="a7a0d-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="a7a0d-153">Rechercher une chaîne à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="a7a0d-153">Find a string using the Excel JavaScript API</span></span>](excel-add-ins-ranges-string-match.md)
- [<span data-ttu-id="a7a0d-154">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="a7a0d-154">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
