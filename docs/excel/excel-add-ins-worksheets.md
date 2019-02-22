---
title: Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 02/15/2018
localization_priority: Priority
ms.openlocfilehash: 6d34807b1511573c507d43dad678811c5c1592ec
ms.sourcegitcommit: 03773fef3d2a380028ba0804739d2241d4b320e5
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2019
ms.locfileid: "30091245"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel

Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de feuilles de calcul utilisant l’API JavaScript pour Excel. Pour une liste complète des propriétés et des méthodes prises en charge par les objets **Worksheet** et **WorksheetCollection**, reportez-vous aux rubriques [Objet Worksheet (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) et [Objet WorksheetCollection (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection).

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

> [!NOTE]
> Une feuille de calcul avec une visibilité «[Très masquée](/javascript/api/excel/excel.sheetvisibility)» ne peut pas être supprimée avec la méthode `delete`. Si vous souhaitez quand-même supprimer la feuille de calcul, vous devez commencer par modifier la visibilité.

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

## <a name="get-a-single-cell-within-a-worksheet"></a>Obtenir une cellule simple dans une feuille de calcul

L’exemple de code suivant obtient la cellule située ligne 2, colonne 5 de la feuille de calcul nommée **Sample**, charge ses propriétés **address** et **values**, et écrit un message dans la console. Les valeurs transmises par la méthode `getCell(row: number, column:number)` sont le numéro de ligne avec indice zéro et le numéro de colonne pour la cellule en cours d’extraction.

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

## <a name="find-all-cells-with-matching-text-preview"></a>Trouver toutes les cellules avec texte correspondant (Aperçu)

> [!NOTE]
> La fonction`findAll` de l’objet de la Feuille de calcul est actuellement disponible uniquement en préversion publique (bêta). Pour utiliser cette fonctionnalité, vous devez utiliser la bibliothèque bêta du CDN Office.js : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Si vous utilisez TypeScript ou si votre éditeur de code utilise des fichiers de définition de type TypeScript pour IntelliSense, utilisez https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

L’objet `Worksheet` dispose d’une méthode`find` pour rechercher une chaîne spécifiée dans la feuille de calcul. Il renvoie un objet`RangeAreas`, qui est une collection d’objets `Range` qui peuvent être modifiés tous en même temps. L’exemple de code suivant recherche toutes les cellules contenant des valeurs égales à la chaîne **Complète** et les colore en vert. Notez que `findAll` génère une erreur `ItemNotFound` si la chaîne spécifiée n’existe pas dans la feuille de calcul. Si vous pensez que la chaîne spécifiée peut ne pas exister dans la feuille de calcul, utilisez la méthode[findAllOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) à la place, pour que votre code gère ce scénario plus facilement.

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
> Cette section décrit comment rechercher des cellules et plages à l’aide des `Worksheet` fonctions de l’objet. Plus d’informations sur l’extraction de plage sont disponibles dans les articles spécifiques.
> - Pour obtenir des exemples qui montrent comment obtenir une plage dans une feuille de calcul à l’aide de l’objet `Range`, reportez-vous à la rubrique [Utiliser des plages à l’aide de l’API JavaScript pour Excel](excel-add-ins-ranges.md).
> - Pour obtenir des exemples qui montrent comment obtenir une plage dans un objet `Table`, reportez-vous à la rubrique [Utiliser des tableaux à l’aide de l’API JavaScript pour Excel](excel-add-ins-tables.md).
> - Pour consulter des exemples qui montrent comment rechercher une grande plage pour plusieurs sous-plages basées sur les caractéristiques de cellule, voir [Travailler avec plusieurs plages simultanément dans des compléments Excel](excel-add-ins-multiple-ranges.md).

## <a name="data-protection"></a>Protection des données

Votre complément permet de contrôler la possibilité qu’a un utilisateur de modifier des données dans une feuille de calcul. La propriété `protection` de la feuille de calcul est un objet [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) avec une méthode `protect()`. L’exemple suivant illustre un scénario de base permettant d’activer/de désactiver la protection complète de la feuille de calcul active.

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

La méthode `protect` présente deux paramètres facultatifs :

- `options` : objet [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) définissant des restrictions de modification spécifiques.
- `password` : chaîne représentant le mot de passe nécessaire pour qu’un utilisateur puisse ignorer la protection et modifier la feuille de calcul.

L’article [Protéger une feuille de calcul](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) comporte davantage d’informations sur la protection des feuilles de calcul et leur modification via l’interface utilisateur Excel.

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
