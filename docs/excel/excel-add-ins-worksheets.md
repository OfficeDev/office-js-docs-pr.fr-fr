---
title: Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 11/27/2018
ms.openlocfilehash: ef74dc622f3e857314874763a54df67bcff1d8ff
ms.sourcegitcommit: 026437bd3819f4e9cd4153ebe60c98ab04e18f4e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/30/2018
ms.locfileid: "26992225"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="e0adf-102">Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e0adf-102">Work with Worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="e0adf-103">Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de feuilles de calcul utilisant l’API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="e0adf-103">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span> <span data-ttu-id="e0adf-104">Pour une liste complète des propriétés et des méthodes prises en charge par les objets **Worksheet** et **WorksheetCollection**, reportez-vous aux rubriques [Objet Worksheet (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) et [Objet WorksheetCollection (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection).</span><span class="sxs-lookup"><span data-stu-id="e0adf-104">For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="e0adf-105">les informations contenues dans cet article s’appliquent uniquement aux feuilles de calcul standard. Elles ne concernent pas les feuilles « chart » ou « macro ».</span><span class="sxs-lookup"><span data-stu-id="e0adf-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="e0adf-106">Obtenir des feuilles de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-106">Get worksheets</span></span>

<span data-ttu-id="e0adf-107">L’exemple de code suivant obtient la collection de feuilles de calcul, charge la propriété **name** de chaque feuille de calcul et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

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
> <span data-ttu-id="e0adf-108">La propriété **id** d’une feuille de calcul identifie de manière unique la feuille de calcul dans un classeur donné et sa valeur ne change pas, même lorsque la feuille de calcul est renommée ou déplacée.</span><span class="sxs-lookup"><span data-stu-id="e0adf-108">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved.</span></span> <span data-ttu-id="e0adf-109">Lorsqu’une feuille de calcul est supprimée d’un classeur dans Excel pour Mac, la propriété **id** de la feuille de calcul supprimée peut être réaffectée à une nouvelle feuille de calcul créée par la suite.</span><span class="sxs-lookup"><span data-stu-id="e0adf-109">When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="e0adf-110">Obtenir la feuille de calcul active</span><span class="sxs-lookup"><span data-stu-id="e0adf-110">Get the active worksheet</span></span>

<span data-ttu-id="e0adf-111">L’exemple de code suivant obtient la feuille de calcul active, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="e0adf-112">Définir la feuille de calcul active</span><span class="sxs-lookup"><span data-stu-id="e0adf-112">Set the active worksheet</span></span>

<span data-ttu-id="e0adf-113">L’exemple de code suivant définit la feuille de calcul active sur la feuille de calcul nommée **Sample**, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-113">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="e0adf-114">S’il n’existe aucune feuille de calcul portant ce nom, la méthode **activate()** lève une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="e0adf-114">If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="e0adf-115">Référencer des feuilles de calcul en fonction de leur position relative</span><span class="sxs-lookup"><span data-stu-id="e0adf-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="e0adf-116">Ces exemples montrent comment référencer une feuille de calcul en fonction de sa position relative.</span><span class="sxs-lookup"><span data-stu-id="e0adf-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="e0adf-117">Obtenir la première feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-117">Get the first worksheet</span></span>

<span data-ttu-id="e0adf-118">L’exemple de code suivant obtient la première feuille de calcul du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="e0adf-119">Obtenir la dernière feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-119">Get the last worksheet</span></span>

<span data-ttu-id="e0adf-120">L’exemple de code suivant obtient la dernière feuille de calcul du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="e0adf-121">Obtenir la feuille de calcul suivante</span><span class="sxs-lookup"><span data-stu-id="e0adf-121">Get the next worksheet</span></span>

<span data-ttu-id="e0adf-122">L’exemple de code suivant obtient la feuille de calcul qui suit la feuille de calcul active du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-122">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="e0adf-123">S’il n’existe aucune feuille de calcul après la feuille de calcul active, la méthode **getNext()** lève une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="e0adf-123">If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="e0adf-124">Obtenir la feuille de calcul précédente</span><span class="sxs-lookup"><span data-stu-id="e0adf-124">Get the previous worksheet</span></span>

<span data-ttu-id="e0adf-125">L’exemple de code suivant obtient la feuille de calcul qui précède la feuille de calcul active du classeur, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-125">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="e0adf-126">S’il n’existe aucune feuille de calcul avant la feuille de calcul active, la méthode **getPrevious()** lève une erreur **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="e0adf-126">If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="e0adf-127">Ajouter une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-127">Add a worksheet</span></span>

<span data-ttu-id="e0adf-p106">L’exemple de code suivant ajoute une nouvelle feuille de calcul nommée **Sample** au classeur, charge ses propriétés **name** et **position**, et écrit un message dans la console. Le nouveau tableur est ajouté après toutes les feuilles de calcul existantes.</span><span class="sxs-lookup"><span data-stu-id="e0adf-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

## <a name="delete-a-worksheet"></a><span data-ttu-id="e0adf-130">Supprimer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-130">Delete a worksheet</span></span>

<span data-ttu-id="e0adf-131">L’exemple de code suivant supprime la dernière feuille de calcul dans le classeur (sous réserve qu’il ne s’agisse pas de la seule feuille dans le classeur) et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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

## <a name="rename-a-worksheet"></a><span data-ttu-id="e0adf-132">Renommer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-132">Rename a worksheet</span></span>

<span data-ttu-id="e0adf-133">L’exemple de code suivant renomme la feuille de calcul active comme suit : **New Name**.</span><span class="sxs-lookup"><span data-stu-id="e0adf-133">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="e0adf-134">Déplacer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-134">Move a worksheet</span></span>

<span data-ttu-id="e0adf-135">L’exemple de code suivant fait passer une feuille de calcul de la dernière position à la première position dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="e0adf-135">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="e0adf-136">Définir la visibilité d’une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-136">Set worksheet visibility</span></span>

<span data-ttu-id="e0adf-137">Ces exemples montrent comment définir la visibilité d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e0adf-137">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="e0adf-138">Masquer une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-138">Hide a worksheet</span></span>

<span data-ttu-id="e0adf-139">L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Sample** à masquer, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-139">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="e0adf-140">Afficher une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-140">Unhide a worksheet</span></span>

<span data-ttu-id="e0adf-141">L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Sample** à afficher, charge sa propriété **name** et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-141">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="get-a-cell-within-a-worksheet"></a><span data-ttu-id="e0adf-142">Obtenir une cellule dans une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-142">Get a cell within a worksheet</span></span>

<span data-ttu-id="e0adf-143">L’exemple de code suivant obtient la cellule située ligne 2, colonne 5 de la feuille de calcul nommée **Sample**, charge ses propriétés **address** et **values**, et écrit un message dans la console.</span><span class="sxs-lookup"><span data-stu-id="e0adf-143">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="e0adf-144">Les valeurs transmises par la méthode **getCell(row: number, column:number)** sont le numéro de ligne avec indice zéro et le numéro de colonne pour la cellule en cours d’extraction.</span><span class="sxs-lookup"><span data-stu-id="e0adf-144">The values that are passed into the **getCell(row: number, column:number)** method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="get-a-range-within-a-worksheet"></a><span data-ttu-id="e0adf-145">Obtenir une plage dans une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="e0adf-145">Get a range within a worksheet</span></span>

<span data-ttu-id="e0adf-146">Pour obtenir des exemples qui montrent comment obtenir une plage dans une feuille de calcul, reportez-vous à la rubrique [Utiliser des plages à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="e0adf-146">For examples that show how to get a range within a worksheet, see [Work with Ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>

## <a name="data-protection"></a><span data-ttu-id="e0adf-147">Protection des données</span><span class="sxs-lookup"><span data-stu-id="e0adf-147">Data protection</span></span>

<span data-ttu-id="e0adf-148">Votre complément permet de contrôler la possibilité qu’a un utilisateur de modifier des données dans une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e0adf-148">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="e0adf-149">La propriété `protection` de la feuille de calcul est un objet [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) avec une méthode `protect()`.</span><span class="sxs-lookup"><span data-stu-id="e0adf-149">The worksheet's `protection` property is a [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="e0adf-150">L’exemple suivant illustre un scénario de base permettant d’activer/de désactiver la protection complète de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="e0adf-150">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

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

<span data-ttu-id="e0adf-151">La méthode `protect` présente deux paramètres facultatifs :</span><span class="sxs-lookup"><span data-stu-id="e0adf-151">The AddDependentLookup`protect` method has two parameters:</span></span>

 - <span data-ttu-id="e0adf-152">`options` : objet [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) définissant des restrictions de modification spécifiques.</span><span class="sxs-lookup"><span data-stu-id="e0adf-152">`options`: A [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
 - <span data-ttu-id="e0adf-153">`password` : chaîne représentant le mot de passe nécessaire pour qu’un utilisateur puisse ignorer la protection et modifier la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="e0adf-153">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="e0adf-154">L’article [Protéger une feuille de calcul](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) comporte davantage d’informations sur la protection des feuilles de calcul et leur modification via l’interface utilisateur Excel.</span><span class="sxs-lookup"><span data-stu-id="e0adf-154">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="e0adf-155">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="e0adf-155">See also</span></span>

- [<span data-ttu-id="e0adf-156">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="e0adf-156">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

