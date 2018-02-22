---
title: Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 12/04/2017
---


# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel

Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de feuilles de calcul utilisant l’API JavaScript pour Excel. Pour une liste complète des propriétés et des méthodes prises en charge par les objets **Worksheet** et **WorksheetCollection**, reportez-vous aux rubriques [Objet Worksheet (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/worksheet) et [Objet WorksheetCollection (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/worksheetcollection).

> [!NOTE]
> les informations contenues dans cet article s’appliquent uniquement aux feuilles de calcul standard. Elles ne concernent pas les feuilles « chart » ou « macro ».

## <a name="get-worksheets"></a>Obtenir des feuilles de calcul

L’exemple de code suivant obtient la collection de feuilles de calcul, charge la propriété **name** de chaque feuille de calcul et écrit un message dans la console.

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
> La propriété **id** d’une feuille de calcul identifie de manière unique la feuille de calcul dans un classeur donné et sa valeur ne change pas, même lorsque la feuille de calcul est renommée ou déplacée. Lorsqu’une feuille de calcul est supprimée d’un classeur dans Excel pour Mac, la propriété **id** de la feuille de calcul supprimée peut être réaffectée à une nouvelle feuille de calcul créée par la suite.

## <a name="get-the-active-worksheet"></a>Obtenir la feuille de calcul active

L’exemple de code suivant obtient la feuille de calcul active, charge sa propriété **name** et écrit un message dans la console.

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

## <a name="set-the-active-worksheet"></a>Définir la feuille de calcul active

L’exemple de code suivant définit la feuille de calcul active sur la feuille de calcul nommée **Sample**, charge sa propriété **name** et écrit un message dans la console. S’il n’existe aucune feuille de calcul portant ce nom, la méthode **activate()** lève une erreur **ItemNotFound**.

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

## <a name="reference-worksheets-by-relative-position"></a>Référencer des feuilles de calcul en fonction de leur position relative

Ces exemples montrent comment référencer une feuille de calcul en fonction de sa position relative.

### <a name="get-the-first-worksheet"></a>Obtenir la première feuille de calcul

L’exemple de code suivant obtient la première feuille de calcul du classeur, charge sa propriété **name** et écrit un message dans la console.

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

### <a name="get-the-last-worksheet"></a>Obtenir la dernière feuille de calcul

L’exemple de code suivant obtient la dernière feuille de calcul du classeur, charge sa propriété **name** et écrit un message dans la console.

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

### <a name="get-the-next-worksheet"></a>Obtenir la feuille de calcul suivante

L’exemple de code suivant obtient la feuille de calcul qui suit la feuille de calcul active du classeur, charge sa propriété **name** et écrit un message dans la console. S’il n’existe aucune feuille de calcul après la feuille de calcul active, la méthode **getNext()** lève une erreur **ItemNotFound**.

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

### <a name="get-the-previous-worksheet"></a>Obtenir la feuille de calcul précédente

L’exemple de code suivant obtient la feuille de calcul qui précède la feuille de calcul active du classeur, charge sa propriété **name** et écrit un message dans la console. S’il n’existe aucune feuille de calcul avant la feuille de calcul active, la méthode **getPrevious()** lève une erreur **ItemNotFound**.

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

## <a name="add-a-worksheet"></a>Ajouter une feuille de calcul

L’exemple de code suivant ajoute une nouvelle feuille de calcul nommée **Sample** au classeur, charge ses propriétés **name** et **position**, et écrit un message dans la console. Le nouveau tableur est ajouté après toutes les feuilles de calcul existantes.

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

## <a name="delete-a-worksheet"></a>Supprimer une feuille de calcul

L’exemple de code suivant supprime la dernière feuille de calcul dans le classeur (sous réserve qu’il ne s’agisse pas de la seule feuille dans le classeur) et écrit un message dans la console.

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

## <a name="rename-a-worksheet"></a>Renommer une feuille de calcul

L’exemple de code suivant renomme la feuille de calcul active comme suit : **New Name**.

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a>Déplacer une feuille de calcul

L’exemple de code suivant fait passer une feuille de calcul de la dernière position à la première position dans le classeur.

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

## <a name="set-worksheet-visibility"></a>Définir la visibilité d’une feuille de calcul

Ces exemples montrent comment définir la visibilité d’une feuille de calcul.

### <a name="hide-a-worksheet"></a>Masquer une feuille de calcul

L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Sample** à masquer, charge sa propriété **name** et écrit un message dans la console.

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

### <a name="unhide-a-worksheet"></a>Afficher une feuille de calcul

L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Sample** à afficher, charge sa propriété **name** et écrit un message dans la console.

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

## <a name="get-a-cell-within-a-worksheet"></a>Obtenir une cellule dans une feuille de calcul

L’exemple de code suivant obtient la cellule située ligne 2, colonne 5 de la feuille de calcul nommée **Sample**, charge ses propriétés **address** et **values**, et écrit un message dans la console. Les valeurs transmises par la méthode **getCell(row: number, column:number)** sont le numéro de ligne avec indice zéro et le numéro de colonne pour la cellule en cours d’extraction.

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

## <a name="get-a-range-within-a-worksheet"></a>Obtenir une plage dans une feuille de calcul

Pour obtenir des exemples qui montrent comment obtenir une plage dans une feuille de calcul, reportez-vous à la rubrique [Utiliser des plages à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges.md).

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet Worksheet (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/worksheet)
- [Objet WorksheetCollection (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/worksheetcollection)
