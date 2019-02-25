---
title: Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: 1597a21940dbe0fbecb4f5976f13088692a10b6f
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/22/2019
ms.locfileid: "30199563"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="67e26-102">Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="67e26-102">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="67e26-103">Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de feuilles de calcul utilisant l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="67e26-103">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span> <span data-ttu-id="67e26-104">Pour une liste complète des propriétés et des méthodes prises en charge par les objets **Worksheet** et **WorksheetCollection**, reportez-vous aux rubriques [Objet Worksheet (API JavaScript pour Excel)](/javascript/api/excel/excel.worksheet) et [Objet WorksheetCollection (API JavaScript pour Excel)](/javascript/api/excel/excel.worksheetcollection).</span><span class="sxs-lookup"><span data-stu-id="67e26-104">For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="67e26-105">les informations contenues dans cet article s’appliquent uniquement aux feuilles de calcul standard. Elles ne concernent pas les feuilles « chart » ou « macro ».</span><span class="sxs-lookup"><span data-stu-id="67e26-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="67e26-106">Obtenir des feuilles de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-106">Get worksheets</span></span>

<span data-ttu-id="67e26-107">L’exemple de code suivant obtient la collection de feuilles de calcul, charge la propriété **name** de chaque feuille de calcul et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length > 1) {
                console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
            } else {
                console.log(`There is one worksheet in the workbook:`);
            }
            for (var i in sheets.items) {
                console.log(sheets.items[i].name);
            }
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="67e26-108">La propriété **id** d’une feuille de calcul identifie de manière unique la feuille de calcul dans un classeur donné et sa valeur ne change pas, même lorsque la feuille de calcul est renommée ou déplacée.</span><span class="sxs-lookup"><span data-stu-id="67e26-108">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved.</span></span> <span data-ttu-id="67e26-109">Lorsqu’une feuille de calcul est supprimée d’un classeur dans Excel pour Mac, la propriété **id** de la feuille de calcul supprimée peut être réaffectée à une nouvelle feuille de calcul créée par la suite.</span><span class="sxs-lookup"><span data-stu-id="67e26-109">When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="67e26-110">Obtenir la feuille de calcul active</span><span class="sxs-lookup"><span data-stu-id="67e26-110">Get the active worksheet</span></span>

<span data-ttu-id="67e26-111">L’exemple de code suivant obtient la feuille de calcul active, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-the-active-worksheet"></a><span data-ttu-id="67e26-112">Définir la feuille de calcul active</span><span class="sxs-lookup"><span data-stu-id="67e26-112">Set the active worksheet</span></span>

<span data-ttu-id="67e26-113">L’exemple de code suivant définit la feuille de calcul active sur la feuille de calcul nommée **Sample**, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-113">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="67e26-114">S’il n’existe aucune feuille de calcul portant ce nom, la méthode **activate()** lève une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="67e26-114">If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="67e26-115">Référencer des feuilles de calcul en fonction de leur position relative</span><span class="sxs-lookup"><span data-stu-id="67e26-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="67e26-116">Ces exemples montrent comment référencer une feuille de calcul en fonction de sa position relative.</span><span class="sxs-lookup"><span data-stu-id="67e26-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="67e26-117">Obtenir la première feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-117">Get the first worksheet</span></span>

<span data-ttu-id="67e26-118">L’exemple de code suivant obtient la première feuille de calcul du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the first worksheet is "${firstSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-last-worksheet"></a><span data-ttu-id="67e26-119">Obtenir la dernière feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-119">Get the last worksheet</span></span>

<span data-ttu-id="67e26-120">L’exemple de code suivant obtient la dernière feuille de calcul du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the last worksheet is "${lastSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-next-worksheet"></a><span data-ttu-id="67e26-121">Obtenir la feuille de calcul suivante</span><span class="sxs-lookup"><span data-stu-id="67e26-121">Get the next worksheet</span></span>

<span data-ttu-id="67e26-122">L’exemple de code suivant obtient la feuille de calcul qui suit la feuille de calcul active du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-122">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="67e26-123">S’il n’existe aucune feuille de calcul après la feuille de calcul active, la méthode **getNext()** lève une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="67e26-123">If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

```js
 Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="67e26-124">Obtenir la feuille de calcul précédente</span><span class="sxs-lookup"><span data-stu-id="67e26-124">Get the previous worksheet</span></span>

<span data-ttu-id="67e26-125">L’exemple de code suivant obtient la feuille de calcul qui précède la feuille de calcul active du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-125">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="67e26-126">S’il n’existe aucune feuille de calcul avant la feuille de calcul active, la méthode **getPrevious()** lève une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="67e26-126">If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="add-a-worksheet"></a><span data-ttu-id="67e26-127">Ajouter une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-127">Add a worksheet</span></span>

<span data-ttu-id="67e26-p106">L’exemple de code suivant ajoute une nouvelle feuille de calcul nommée **Sample** au classeur, charge ses propriétés **name** et **position**, et écrit un message dans la console. Le nouveau tableur est ajouté après toutes les feuilles de calcul existantes.</span><span class="sxs-lookup"><span data-stu-id="67e26-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;

    var sheet = sheets.add("Sample");
    sheet.load("name, position");

    return context.sync()
        .then(function () {
            console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
        });
}).catch(errorHandlerFunction);
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="67e26-130">Supprimer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-130">Delete a worksheet</span></span>

<span data-ttu-id="67e26-131">L’exemple de code suivant supprime la dernière feuille de calcul dans le classeur (sous réserve qu’il ne s’agisse pas de la seule feuille dans le classeur) et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length === 1) {
                console.log("Unable to delete the only worksheet in the workbook");
            } else {
                var lastSheet = sheets.items[sheets.items.length - 1];

                console.log(`Deleting worksheet named "${lastSheet.name}"`);
                lastSheet.delete();

                return context.sync();
            };
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="67e26-132">Une feuille de calcul avec une visibilité «[Très masquée](/javascript/api/excel/excel.sheetvisibility)» ne peut pas être supprimée avec la méthode `delete`.</span><span class="sxs-lookup"><span data-stu-id="67e26-132">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="67e26-133">Si vous souhaitez quand-même supprimer la feuille de calcul, vous devez commencer par modifier la visibilité.</span><span class="sxs-lookup"><span data-stu-id="67e26-133">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="67e26-134">Renommer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-134">Rename a worksheet</span></span>

<span data-ttu-id="67e26-135">L’exemple de code suivant renomme la feuille de calcul active comme suit : **New Name**.</span><span class="sxs-lookup"><span data-stu-id="67e26-135">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="67e26-136">Déplacer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-136">Move a worksheet</span></span>

<span data-ttu-id="67e26-137">L’exemple de code suivant fait passer une feuille de calcul de la dernière position à la première position dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="67e26-137">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items");

    return context.sync()
        .then(function () {
            var lastSheet = sheets.items[sheets.items.length - 1];
            lastSheet.position = 0;

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

## <a name="set-worksheet-visibility"></a><span data-ttu-id="67e26-138">Définir la visibilité d’une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-138">Set worksheet visibility</span></span>

<span data-ttu-id="67e26-139">Ces exemples montrent comment définir la visibilité d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="67e26-139">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="67e26-140">Masquer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-140">Hide a worksheet</span></span>

<span data-ttu-id="67e26-141">L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Sample** à masquer, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-141">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is hidden`);
        });
}).catch(errorHandlerFunction);
```

### <a name="unhide-a-worksheet"></a><span data-ttu-id="67e26-142">Afficher une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-142">Unhide a worksheet</span></span>

<span data-ttu-id="67e26-143">L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Sample** à afficher, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-143">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is visible`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="67e26-144">Obtenir une cellule simple dans une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="67e26-144">Get a single cell within a worksheet</span></span>

<span data-ttu-id="67e26-145">L’exemple de code suivant obtient la cellule située ligne 2, colonne 5 de la feuille de calcul nommée **Sample**, charge ses propriétés **address** et **values**, et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="67e26-145">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="67e26-146">Les valeurs transmises par la méthode `getCell(row: number, column:number)` sont le numéro de ligne avec indice zéro et le numéro de colonne pour la cellule en cours d’extraction.</span><span class="sxs-lookup"><span data-stu-id="67e26-146">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var cell = sheet.getCell(1, 4);
    cell.load("address, values");

    return context.sync()
        .then(function() {
            console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
        })
}).catch(errorHandlerFunction);
```

## <a name="find-all-cells-with-matching-text-preview"></a><span data-ttu-id="67e26-147">Trouver toutes les cellules avec du texte correspondant (préversion)</span><span class="sxs-lookup"><span data-stu-id="67e26-147">Find all cells with matching text (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="67e26-148">La fonction`findAll` de l’objet Worksheet est actuellement disponible uniquement en préversion publique.</span><span class="sxs-lookup"><span data-stu-id="67e26-148">The Worksheet object's `findAll` function is currently available only in public preview (beta).</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

<span data-ttu-id="67e26-149">L’objet `Worksheet` dispose d’une méthode`find` pour rechercher une chaîne spécifiée dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="67e26-149">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="67e26-150">Il renvoie un objet`RangeAreas`, qui est une collection d’objets `Range` qui peuvent être modifiés tous en même temps.</span><span class="sxs-lookup"><span data-stu-id="67e26-150">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="67e26-151">L’exemple de code suivant recherche toutes les cellules contenant des valeurs égales à la chaîne **Complète** et les colore en vert.</span><span class="sxs-lookup"><span data-stu-id="67e26-151">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="67e26-152">Notez que `findAll` génère une erreur `ItemNotFound` si la chaîne spécifiée n’existe pas dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="67e26-152">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="67e26-153">Si vous pensez que la chaîne spécifiée peut ne pas exister dans la feuille de calcul, utilisez la méthode[findAllOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) à la place, pour que votre code gère ce scénario plus facilement.</span><span class="sxs-lookup"><span data-stu-id="67e26-153">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    return context.sync()
        .then(function() {
            foundRanges.format.fill.color = "green"
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="67e26-154">Cette section décrit comment rechercher des cellules et plages à l’aide des `Worksheet` fonctions de l’objet.</span><span class="sxs-lookup"><span data-stu-id="67e26-154">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="67e26-155">Plus d’informations sur l’extraction de plage sont disponibles dans les articles spécifiques.</span><span class="sxs-lookup"><span data-stu-id="67e26-155">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="67e26-156">Pour obtenir des exemples qui montrent comment obtenir une plage dans une feuille de calcul à l’aide de l’objet `Range`, reportez-vous à la rubrique [Utiliser des plages à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="67e26-156">For examples that show how to get a range within a worksheet using the `Range` object, see [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>
> - <span data-ttu-id="67e26-157">Pour obtenir des exemples qui montrent comment obtenir une plage dans un objet `Table`, reportez-vous à la rubrique [Utiliser des tableaux à l’aide de l’API JavaScript pour Excel](excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="67e26-157">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="67e26-158">Pour consulter des exemples qui montrent comment rechercher une grande plage pour plusieurs sous-plages basées sur les caractéristiques de cellule, voir [Travailler avec plusieurs plages simultanément dans des compléments Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="67e26-158">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="data-protection"></a><span data-ttu-id="67e26-159">Protection des données</span><span class="sxs-lookup"><span data-stu-id="67e26-159">Data protection</span></span>

<span data-ttu-id="67e26-160">Votre complément permet de contrôler la possibilité qu’a un utilisateur de modifier des données dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="67e26-160">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="67e26-161">La propriété `protection` de la feuille de calcul est un objet [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) avec une méthode `protect()`.</span><span class="sxs-lookup"><span data-stu-id="67e26-161">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="67e26-162">L’exemple suivant illustre un scénario de base permettant d’activer/de désactiver la protection complète de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="67e26-162">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");

    return context.sync().then(function() {
        if (!activeSheet.protection.protected) {
            activeSheet.protection.protect();
        }
    })
}).catch(errorHandlerFunction);
```

<span data-ttu-id="67e26-163">La méthode `protect` présente deux paramètres facultatifs :</span><span class="sxs-lookup"><span data-stu-id="67e26-163">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="67e26-164">`options` : objet [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) définissant des restrictions de modification spécifiques.</span><span class="sxs-lookup"><span data-stu-id="67e26-164">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="67e26-165">`password` : chaîne représentant le mot de passe nécessaire pour qu’un utilisateur puisse ignorer la protection et modifier la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="67e26-165">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="67e26-166">L’article [Protéger une feuille de calcul](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) comporte davantage d’informations sur la protection des feuilles de calcul et leur modification via l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="67e26-166">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="67e26-167">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="67e26-167">See also</span></span>

- [<span data-ttu-id="67e26-168">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="67e26-168">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
