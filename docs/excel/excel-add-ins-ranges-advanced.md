---
title: Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)
description: ''
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: c8fbe1dcc75080c932b4c3e2946fe62747d35c6b
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395595"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="9ef0f-102">Utiliser les plages à l’aide de l’API JavaScript Excel (avancé)</span><span class="sxs-lookup"><span data-stu-id="9ef0f-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="9ef0f-103">Cet article génère des informations dans[ Utiliser des plages à l’aide de l’API JavaScript Excel (fondamental)](excel-add-ins-ranges.md) en fournissant les exemples de code qui affichent la manière d’exécuter plus de tâches avancées avec des plages à l’aide de l’API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="9ef0f-104">Pour obtenir une liste complète des propriétés et des méthodes prises en charge par l’objet **Range**, reportez-vous à la rubrique [Objet Range (API JavaScript pour Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="9ef0f-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="9ef0f-105">Utiliser des dates à l’aide de plug-in Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="9ef0f-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="9ef0f-106">La[bibliothèque Moment JavaScript](https://momentjs.com/)fournit une manière pratique d’utiliser les dates et les horodateurs.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="9ef0f-107">Le[plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate)convertit le format des moments dans un préférable pour Excel.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="9ef0f-108">Il s’agit du même format que la[fonction NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)renvoie.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="9ef0f-109">Le code suivant affiche la manière d’établir la plage à**B4**vers un horodateur du moment:</span><span class="sxs-lookup"><span data-stu-id="9ef0f-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="9ef0f-110">Il s’agit d’une technique similaire pour récupérer la date de la cellule et la convertir en un moment ou autre format, comme démontré dans le code suivant:</span><span class="sxs-lookup"><span data-stu-id="9ef0f-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="9ef0f-111">Votre complément devra mettre en forme les plages pour afficher les dates dans une forme plus lisible.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="9ef0f-112">L’exemple de`"[$-409]m/d/yy h:mm AM/PM;@"`affiche une heure comme «12/3/18 3:57 PM».</span><span class="sxs-lookup"><span data-stu-id="9ef0f-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="9ef0f-113">Pour plus d’informations concernant les formats de date et d’heure , veuillez consulter les «Instructions relatifs aux formats de date et heure» dans l’article[ Instructions revoir afin de personnaliser le format numérique](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="9ef0f-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously"></a><span data-ttu-id="9ef0f-114">Travailler simultanément avec plusieurs plages</span><span class="sxs-lookup"><span data-stu-id="9ef0f-114">Work with multiple ranges simultaneously</span></span>

<span data-ttu-id="9ef0f-115">L’`RangeAreas`objet laisse votre complément exécuter des opérations sur plusieurs plages en même temps.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-115">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="9ef0f-116">Ces plages peuvent être adjacentes, mais cela n’est pas obligatoire.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-116">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="9ef0f-117">`RangeAreas`sont abordés plus loin dans l’article[Travailler simultanément avec plusieurs plages dans des compléments Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="9ef0f-117">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range"></a><span data-ttu-id="9ef0f-118">Rechercher des cellules spéciales dans une plage</span><span class="sxs-lookup"><span data-stu-id="9ef0f-118">Find special cells within a range</span></span>

<span data-ttu-id="9ef0f-119">Les méthodes`Range.getSpecialCells()` et `Range.getSpecialCellsOrNullObject()`recherchent des plages basées sur les caractéristiques de leurs cellules et les types de valeurs de leurs cellules.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-119">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="9ef0f-120">Ces deux méthodes renvoient à des`RangeAreas`objets.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-120">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="9ef0f-121">Voici les signatures des méthodes à partir des types de fichiers de données TypeScript:</span><span class="sxs-lookup"><span data-stu-id="9ef0f-121">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="9ef0f-122">L’exemple suivant utilise la méthode`getSpecialCells`pour rechercher toutes les cellules contenant les formules.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-122">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="9ef0f-123">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="9ef0f-123">About this code, note:</span></span>

- <span data-ttu-id="9ef0f-124">Cela limite la partie de la feuille qui nécessite d’être recherchée en appelant d’abord`Worksheet.getUsedRange`et en appelant`getSpecialCells`uniquement pour cette plage.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-124">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="9ef0f-125">La`getSpecialCells`méthode renvoie un`RangeAreas`objet, toutes les cellules alors dotées de formules seront colorées en rose même si elles ne sont pas adjacentes.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-125">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="9ef0f-126">Si aucune cellule avec la caractéristique ciblée n’existe dans la plage `getSpecialCells` lève une erreur**ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-126">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="9ef0f-127">Cela dévie le flux de contrôle vers un(e)`catch`bloc/méthode, s’il en existe.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-127">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="9ef0f-128">S’il n’existe pas`catch`, l’erreur arrête la fonction.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-128">If there isn't a `catch` block, the error halts the function.</span></span>

<span data-ttu-id="9ef0f-129">Si vous attendez que des cellules avec la caractéristique ciblée existent toujours, vous souhaiterez probablement que votre code  lève une erreur si ces cellules ne sont pas là.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-129">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="9ef0f-130">Mais dans les scénarios où les cellules ne correspondent pas; votre code doit vérifier cette possibilité et le gérer gracieusement sans émettre d’erreur.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-130">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="9ef0f-131">Vous pouvez obtenir ce comportement avec la `getSpecialCellsOrNullObject`méthode et sa propriété renvoyée`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-131">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="9ef0f-132">Cet exemple utilise les valeurs suivantes.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-132">The following example uses this pattern.</span></span> <span data-ttu-id="9ef0f-133">Tenez compte des informations suivantes :</span><span class="sxs-lookup"><span data-stu-id="9ef0f-133">About this code, note:</span></span>

- <span data-ttu-id="9ef0f-134">La méthode`getSpecialCellsOrNullObject`renvoie toujours un objet proxy, donc il ne s’agit jamais du sens`null`JavaScript ordinaire.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-134">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="9ef0f-135">Mais si les cellules non correspondantes sont introuvables, la propriété`isNullObject` de l’objet est établi à`true`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-135">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="9ef0f-136">Il appelle`context.sync`*avant*de tester la propriété`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-136">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="9ef0f-137">Il s’agit d’une condition avec toutes les méthodes et propriétés`*OrNullObject`, car vous devez toujours télécharger et synchroniser une propriété afin de le lire. </span><span class="sxs-lookup"><span data-stu-id="9ef0f-137">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="9ef0f-138">Cependant, il n’est pas nécessaire de télécharger*de manière explicite*la propriété`isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-138">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="9ef0f-139">Il est automatiquement téléchargé par le`context.sync`même si`load`n’est pas appelé sur l’objet.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-139">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="9ef0f-140">Pour plus d'informations, consultez le[\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="9ef0f-140">For more information, see [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span></span>
- <span data-ttu-id="9ef0f-141">Vous pouvez tester ce code en sélectionnant d’abord une plage qui n’a pas de cellules de formule et en l’exécutant.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-141">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="9ef0f-142">Puis sélectionnez une plage qui dispose au moins d’une cellule dotée d’une formule et en l’exécutant à nouveau.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-142">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="9ef0f-143">Par simplicité, tous les autres exemples dans cet article, utilisez la méthode`getSpecialCells`au lieu de`getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-143">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="9ef0f-144">Réduisez les cellules cibles avec les types de valeur de cellule</span><span class="sxs-lookup"><span data-stu-id="9ef0f-144">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="9ef0f-145">Les méthodes`Range.getSpecialCells()` et `Range.getSpecialCellsOrNullObject()`acceptent un deuxième paramètre facultatif utilisé pour affiner davantage les cellules ciblées.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-145">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="9ef0f-146">Ce deuxième paramètre est un`Excel.SpecialCellValueType` que vous utilisez afin de spécifier que vous souhaitez uniquement les cellules qui contiennent certains types de valeurs.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-146">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="9ef0f-147">Le paramètre `Excel.SpecialCellValueType` peut uniquement être utilisé si le paramètre `Excel.SpecialCellType` est défini sur `Excel.SpecialCellType.formulas`ou `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-147">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="9ef0f-148">Test d’un type de valeur de cellule unique</span><span class="sxs-lookup"><span data-stu-id="9ef0f-148">Test for a single cell value type</span></span>

<span data-ttu-id="9ef0f-149">Le `Excel.SpecialCellValueType` enum dispose de ces quatre types de base (outre les autres valeurs combinées décrites plus loin dans cette section):</span><span class="sxs-lookup"><span data-stu-id="9ef0f-149">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="9ef0f-150">`Excel.SpecialCellValueType.logical` (ce qui signifie booléen)</span><span class="sxs-lookup"><span data-stu-id="9ef0f-150">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="9ef0f-151">L’exemple suivant recherche les cellules spéciaux qui sont des constantes numériques et colore les cellules en rose.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-151">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="9ef0f-152">Tenez compte du code suivant:</span><span class="sxs-lookup"><span data-stu-id="9ef0f-152">About this code, note:</span></span>

- <span data-ttu-id="9ef0f-153">Il surlignera uniquement les cellules qui disposent une valeur de nombre littérale.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-153">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="9ef0f-154">Il surlignera les cellules qui disposent une formule (même si le résultat est un nombre) ou un booléen, un texte ou des cellules d’instruction d’erreur.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-154">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="9ef0f-155">Pour tester le code, assurez-vous que la feuille de calcul dispose de certaines cellules avec des valeurs de nombre littérales, certaines avec d’autres sortes de valeurs littérales, et certaines avec des formules.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-155">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="9ef0f-156">Test d’un type de valeur de cellule multiple</span><span class="sxs-lookup"><span data-stu-id="9ef0f-156">Test for multiple cell value types</span></span>

<span data-ttu-id="9ef0f-157">Parfois, vous avez besoin d’exécuter plus d’un type de valeur de cellule, tel que toutes les cellules à valeur de texte et à valeur booléen (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="9ef0f-157">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="9ef0f-158">Le `Excel.SpecialCellValueType` enum comporte des valeurs avec les types combinés.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-158">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="9ef0f-159">Par exemple,`Excel.SpecialCellValueType.logicalText`cible toutes les cellules à valeur texte et booléen.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-159">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="9ef0f-160">`Excel.SpecialCellValueType.all` est la valeur par défaut, ce qui ne limite pas les types de valeur de cellule renvoyés.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-160">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="9ef0f-161">L’exemple suivant surligne toutes les cellules dotées de formules qui produisent les valeurs de nombre ou booléennes.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-161">The following example colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="copy-and-paste"></a><span data-ttu-id="9ef0f-162">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="9ef0f-162">Copy and paste</span></span>

<span data-ttu-id="9ef0f-163">La fonction`copyFrom`de la plage reproduit le comportement de copier-coller de l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-163">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="9ef0f-164">L’objet plage sur lequel`copyFrom`est appelé est la destination.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-164">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="9ef0f-165">La source à copier est transmise en tant que plage ou qu’adresse de chaîne représentant une plage.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-165">The source to be copied is passed as a range or a string address representing a range.</span></span>
<span data-ttu-id="9ef0f-166">L’exemple de code suivant copie les données de la plage **A1:E1** dans la plage commençant en **G1** (ce qui aboutit à un collage dans la plage **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="9ef0f-166">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="9ef0f-167">`Range.copyFrom`dispose de trois paramètres facultatifs.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-167">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="9ef0f-168">`copyType` spécifie les données copiées de la source vers la destination.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-168">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="9ef0f-169">`Excel.RangeCopyType.formulas` transfère les formules dans les cellules sources en préservant le positionnement relatif des plages de ces formules.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-169">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="9ef0f-170">Les entrées autres que des formules sont copiées telles quelles.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-170">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="9ef0f-171">`Excel.RangeCopyType.values` copie les valeurs des données et, s’il s’agit d’une formule, le résultat de celle-ci.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-171">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="9ef0f-172">`Excel.RangeCopyType.formats` copie la mise en forme de la plage, y compris la police, la couleur et d’autres paramètres de mise en forme, mais aucune valeur.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-172">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="9ef0f-173">`Excel.RangeCopyType.all` (option par défaut) copie les données et la mise en forme, en conservant les formules éventuelles des cellules.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-173">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="9ef0f-174">`skipBlanks` définit si les cellules vides sont copiées dans la destination.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-174">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="9ef0f-175">Quand la valeur est true, `copyFrom` ignore les cellules vides de la plage source.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-175">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="9ef0f-176">Les cellules ignorées ne remplacent pas les données existantes dans les cellules correspondantes de la plage de destination.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-176">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="9ef0f-177">La valeur par défaut est false.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-177">The default is false.</span></span>

<span data-ttu-id="9ef0f-178">`transpose` détermine si les données sont ou non transposées, ce qui signifie que ses lignes et colonnes sont permutées dans l’emplacement source.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-178">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="9ef0f-179">Une plage transposée est renversée le long de la diagonale principale, de sorte que les lignes **1**, **2** et **3** deviennent les colonnes **A**, **B** et **C**.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-179">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="9ef0f-180">L’exemple de code et les images suivants illustrent ce comportement dans un scénario simple.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-180">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="9ef0f-181">*Avant l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="9ef0f-181">*Before the preceding function has been run.*</span></span>

![Données dans Excel avant exécution de la méthode de copie de la plage.](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="9ef0f-183">*Après l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="9ef0f-183">*After the preceding function has been run.*</span></span>

![Données dans Excel après exécution de la méthode de copie de la plage.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a><span data-ttu-id="9ef0f-185">Supprimer les doublons</span><span class="sxs-lookup"><span data-stu-id="9ef0f-185">Remove duplicates</span></span>

<span data-ttu-id="9ef0f-186">La fonction`removeDuplicates`de l’objet de la plage retire les rangées avec les entrées en doublon dans les colonnes spécifiées.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-186">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="9ef0f-187">La fonction circule à travers chaque rangée de la plage de l’index à la valeur la plus basse à l’index à la valeur la plus haute de la plage ( du haut vers le bas).</span><span class="sxs-lookup"><span data-stu-id="9ef0f-187">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="9ef0f-188">Une rangée est supprimée si une valeur dans sa/ses colonne(s) spécifiée(s) apparue(s) plus tôt dans la plage.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-188">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="9ef0f-189">Les rangées de la plage en-dessous de la rangée supprimée sont déplacées.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-189">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="9ef0f-190">`removeDuplicates` n’affecte pas la position des cellules en dehors de la rangée.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-190">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="9ef0f-191">`removeDuplicates`prend un `number[]` représentant les indices de la colonne qui sont vérifiés pour les doublons.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-191">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="9ef0f-192">Ce tableau est à base zéro et lié à la rangée, et non à la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-192">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="9ef0f-193">La fonction prend également un paramètre booléen qui spécifie si la première rangée est un-tête.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-193">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="9ef0f-194">Lorsque**true**, la rangée du dessus est ignorée lorsque les doublons sont pris en considération.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-194">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="9ef0f-195">La fonction`removeDuplicates`renvoie un objet`RemoveDuplicatesResult` qui spécifie le nombre de rangée retirées et le nombre de rangées uniques restantes.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-195">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="9ef0f-196">Lors de l’usage d’une fonction`removeDuplicates`de la plage, gardez ce qui suit à l’esprit:</span><span class="sxs-lookup"><span data-stu-id="9ef0f-196">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="9ef0f-197">`removeDuplicates`considère les valeurs de cellule, et non les résultats de la fonction.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-197">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="9ef0f-198">Si deux fonctions différentes évaluent le même résultat, les valeurs de la cellule ne sont pas considérées comme doublons.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-198">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="9ef0f-199">Les cellules vides ne sont pas ignorées par`removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-199">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="9ef0f-200">La valeur d’une cellule vide est traitée comme toute autre valeur.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-200">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="9ef0f-201">Cela signifie que les rangées vides contenues au sein de la plage seront incluses dans le `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-201">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="9ef0f-202">L’exemple suivant affiche la suppression des entrées avec des valeurs de doublons dans la première colonne.</span><span class="sxs-lookup"><span data-stu-id="9ef0f-202">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="9ef0f-203">*Avant l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="9ef0f-203">*Before the preceding function has been run.*</span></span>

![Données dans Excel avant exécution de la méthode de copie de suppression de la plage.](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="9ef0f-205">*Après l’exécution de la fonction précédente.*</span><span class="sxs-lookup"><span data-stu-id="9ef0f-205">*After the preceding function has been run.*</span></span>

![Données dans Excel après exécution de la méthode de copie de suppression de la plage.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="9ef0f-207">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="9ef0f-207">See also</span></span>

- [<span data-ttu-id="9ef0f-208">Utiliser les plages à l’aide de l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="9ef0f-208">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="9ef0f-209">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="9ef0f-209">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9ef0f-210">Travailler simultanément avec plusieurs plages dans des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="9ef0f-210">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
