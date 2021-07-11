---
title: Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel
description: Exemples de code qui montrent comment effectuer des tâches courantes avec des feuilles de calcul à l’aide Excel API JavaScript.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: a8a7da6ce01f8c0cc82c8ab9c764b032027f585c
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349412"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Utiliser des feuilles de calcul à l’aide de l’API JavaScript pour Excel

Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes à l’aide de feuilles de calcul utilisant l’API JavaScript pour Excel. Pour une liste complète des propriétés et des méthodes prises en charge par les objets `Worksheet` et `WorksheetCollection`, reportez-vous aux rubriques [Objet Worksheet (API JavaScript pour Excel)](/javascript/api/excel/excel.worksheet) et [Objet WorksheetCollection (API JavaScript pour Excel)](/javascript/api/excel/excel.worksheetcollection).

> [!NOTE]
> les informations contenues dans cet article s’appliquent uniquement aux feuilles de calcul standard. Elles ne concernent pas les feuilles « chart » ou « macro ».

## <a name="get-worksheets"></a>Obtenir des feuilles de calcul

L’exemple de code suivant obtient la collection de feuilles de calcul, charge la propriété `name` de chaque feuille de calcul et écrit un message dans la console.

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
            sheets.items.forEach(function (sheet) {
              console.log(sheet.name);
            });
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> La propriété `id` d’une feuille de calcul identifie de manière unique la feuille de calcul dans un classeur donné et sa valeur ne change pas, même lorsque la feuille de calcul est renommée ou déplacée. Lorsqu’une feuille de calcul est supprimée d’un classeur dans Excel sur Mac, la propriété `id` de la feuille de calcul supprimée peut être réaffectée à une nouvelle feuille de calcul créée par la suite.

## <a name="get-the-active-worksheet"></a>Obtenir la feuille de calcul active

L’exemple de code suivant obtient la feuille de calcul active, charge sa propriété `name` et écrit un message dans la console.

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

L’exemple de code suivant définit la feuille de calcul active sur la feuille de calcul nommée **Sample**, charge sa propriété `name` et écrit un message dans la console. S’il n’existe aucune feuille de calcul portant ce nom, la méthode `activate()` lève une erreur `ItemNotFound`.

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

L’exemple de code suivant obtient la première feuille de calcul du classeur, charge sa propriété `name` et écrit un message dans la console.

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

L’exemple de code suivant obtient la dernière feuille de calcul du classeur, charge sa propriété `name` et écrit un message dans la console.

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

L’exemple de code suivant obtient la feuille de calcul qui suit la feuille de calcul active du classeur, charge sa propriété `name` et écrit un message dans la console. S’il n’existe aucune feuille de calcul après la feuille de calcul active, la méthode `getNext()` lève une erreur `ItemNotFound`.

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

L’exemple de code suivant obtient la feuille de calcul qui précède la feuille de calcul active du classeur, charge sa propriété `name` et écrit un message dans la console. S’il n’existe aucune feuille de calcul avant la feuille de calcul active, la méthode `getPrevious()` lève une erreur `ItemNotFound`.

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

L’exemple de code suivant ajoute une nouvelle feuille de calcul nommée **Sample** au classeur, charge ses propriétés `name` et `position`, et écrit un message dans la console. Le nouveau tableur est ajouté après toutes les feuilles de calcul existantes.

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

### <a name="copy-an-existing-worksheet"></a>Copier une feuille de calcul existante

`Worksheet.copy` ajoute une nouvelle feuille de calcul qui est une copie d’une feuille de calcul existante. Le nom de la nouvelle feuille de calcul aura un nombre ajouté à la fin, de façon cohérente avec la copie d’une feuille de calcul dans l’interface utilisateur d’Excel (par exemple, **MySheet (2)**). `Worksheet.copy` peut prendre deux paramètres, qui sont tous deux facultatifs :

- `positionType` – Un enum [WorksheetPositionType ](/javascript/api/excel/excel.worksheetpositiontype) spécifiant l’emplacement dans le classeur où la nouvelle feuille de calcul doit être ajoutée.
- `relativeTo` – Si le `positionType` est `Before` ou `After`, vous devez spécifier une feuille de calcul par rapport à laquelle ajouter la nouvelle feuille (ce paramètre répond à la question « Avant ou après quoi ? »).

L’exemple de code suivant copie la feuille de calcul active et insère la nouvelle feuille directement après la feuille de calcul active.

```js
Excel.run(function (context) {
    var myWorkbook = context.workbook;
    var sampleSheet = myWorkbook.worksheets.getActiveWorksheet();
    var copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);
    return context.sync();
});
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

L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Exemple** à masquer, charge sa propriété `name` et écrit un message dans la console.

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

L’exemple de code suivant définit la visibilité de la feuille de calcul nommée **Exemple** à afficher, charge sa propriété `name` et écrit un message dans la console.

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

L’exemple de code suivant obtient la cellule située ligne 2, colonne 5 de la feuille de calcul nommée **Sample**, charge ses propriétés `address` et `values`, et écrit un message dans la console. Les valeurs transmises par la méthode `getCell(row: number, column:number)` sont le numéro de ligne avec indice zéro et le numéro de colonne pour la cellule en cours d’extraction.

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

## <a name="detect-data-changes"></a>Détecter les modifications de données

Votre complément peut avoir besoin de réagir aux utilisateurs modifiant les données dans une feuille de calcul. Pour détecter ces modifications, vous pouvez [inscrire un gestionnaire d’événements](excel-add-ins-events.md#register-an-event-handler) à l’événement `onChanged` d’une feuille de calcul. Le gestionnaires d’événements de l’événement `onChanged` reçoit un objet [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) lorsque l’événement se déclenche.

L’objet `WorksheetChangedEventArgs` fournit des informations sur les modifications et la source. Puisque `onChanged` se déclenche lorsque le format ou la valeur des données sont modifiés, il peut être utile que votre complément vérifie si les valeurs ont réellement été modifiées. La propriété de `details` regroupe ces informations en tant qu’un [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail). L’exemple de code suivant illustre la procédure d’affichage des valeurs et des types d’une cellule qui a été modifiée, avant et après modification.

```js
// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="detect-formula-changes"></a>Détecter les modifications de formule

Votre add-in peut suivre les modifications apportées aux formules dans une feuille de calcul. Cela est utile lorsqu’une feuille de calcul est connectée à une base de données externe. Lorsque la formule change dans la feuille de calcul, l’événement dans ce scénario déclenche les mises à jour correspondantes dans la base de données externe.

Pour détecter les modifications apportées aux formules, inscrivez un [handler](excel-add-ins-events.md#register-an-event-handler) d’événements pour [l’événement onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged) d’une feuille de calcul. Les handlers d’événements `onFormulaChanged` de l’événement reçoivent un objet [WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs) lorsque l’événement se déclenche.

> [!IMPORTANT]
> `onFormulaChanged`L’événement détecte lorsqu’une formule elle-même change, et non la valeur de données résultant du calcul de la formule.

L’exemple de code suivant montre comment inscrire le handler d’événements, utiliser l’objet pour récupérer le tableau formulaDetails de la formule modifiée, puis imprimer les détails sur la formule modifiée avec les `onFormulaChanged` `WorksheetFormulaChangedEventArgs` propriétés [FormulaChangedEventDetail.](/javascript/api/excel/excel.formulachangedeventdetail) [](/javascript/api/excel/excel.worksheetformulachangedeventargs#formulaDetails)

> [!NOTE]
> Cet exemple de code fonctionne uniquement lorsqu’une seule formule est modifiée.

```js
Excel.run(function (context) {
    // Retrieve the worksheet named "Sample".
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Register the formula changed event handler for this worksheet.
    sheet.onFormulaChanged.add(formulaChangeHandler);

    return context.sync();
});

function formulaChangeHandler(event) {
    Excel.run(function (context) {
        // Retrieve details about the formula change event.
        // Note: This method assumes only a single formula is changed at a time. 
        var cellAddress = event.formulaDetails[0].cellAddress;
        var previousFormula = event.formulaDetails[0].previousFormula;
        var source = event.source;
    
        // Print out the change event details.
        console.log(
          `The formula in cell ${cellAddress} changed. 
          The previous formula was: ${previousFormula}. 
          The source of the change was: ${source}.`
        );         
    });
}
```

## <a name="handle-sorting-events"></a>Gérer les événements de tri

Les  événements `onColumnSorted` et `onRowSorted` indiquent quand les données d’une feuille de calcul sont triées. Ces événements sont connectés à des objets individuels `Worksheet` et aux classeurs `WorkbookCollection`. Il se déclenche si le tri est effectué par programme ou manuellement via l’interface utilisateur d’Excel.

> [!NOTE]
> `onColumnSorted` est déclenché lorsque les colonnes sont triées suite à une opération de tri de gauche à droite. `onRowSorted` est déclenché lorsque les lignes sont triées suite à une opération de tri de haut en bas. Le tri d’un tableau à l’aide du menu déroulant sur un en-tête de colonne génère un événement `onRowSorted`. L’événement correspond au déplacement, et non à ce qui est considéré comme les critères de tri.

Les événements `onColumnSorted` et `onRowSorted` fournissent leurs rappels avec [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) ou [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs), respectivement. Ces éléments fournissent des détails supplémentaires sur l’événement. En particulier, les `EventArgs` ont une propriété `address` qui représente les lignes ou les colonnes déplacées suite à l’opération de tri. Une cellule avec du contenu trié est incluse, même si la valeur de cette cellule ne faisait pas partie des critères de tri.

Les images suivantes montrent les plages retournées par la propriété `address` pour les événements de tri. Voici d’abord les exemples de données avant le tri :

![Les données de tableau Excel avant d’être triées.](../images/excel-sort-event-before.png)

Si un tri de haut en bas est effectué sur «**Q1**» (les valeurs dans «**B**») les lignes mises en surbrill plan suivantes sont renvoyées par `WorksheetRowSortedEventArgs.address` .

![Données d’un tableau dans Excel après un tri de haut en bas. Les lignes qui ont été déplacées sont mises en surbrillance.](../images/excel-sort-event-after-row.png)

Si un tri de gauche à droite est effectué sur «**Quinces**» (les valeurs de «**4**») sur les données d’origine, les colonnes mises en surbrillance suivantes sont renvoyées par `WorksheetColumnsSortedEventArgs.address` .

![Données d’un tableau dans Excel après un tri de gauche à droite. Les colonnes qui ont été déplacées sont mises en surbrillance.](../images/excel-sort-event-after-column.png)

L’exemple de code suivant montre comment inscrire un gestionnaire d’événements pour l’événement `Worksheet.onRowSorted`. Le rappel du gestionnaire efface la couleur de remplissage de la plage, puis remplit les cellules des lignes déplacées.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // This will fire whenever a row has been moved as the result of a sort action.
    sheet.onRowSorted.add(function (event) {
        return Excel.run(function (context) {
            console.log("Row sorted: " + event.address);
            var sheet = context.workbook.worksheets.getActiveWorksheet();

            // Clear formatting for section, then highlight the sorted area.
            sheet.getRange("A1:E5").format.fill.clear();
            if (event.address !== "") {
                sheet.getRanges(event.address).format.fill.color = "yellow";
            }

            return context.sync();
        });
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="find-all-cells-with-matching-text"></a>Trouver toutes les cellules avec du texte correspondant

L’objet `Worksheet` dispose d’une méthode`find` pour rechercher une chaîne spécifiée dans la feuille de calcul. Il renvoie un objet`RangeAreas`, qui est une collection d’objets `Range` qui peuvent être modifiés tous en même temps. L’exemple de code suivant recherche toutes les cellules contenant des valeurs égales à la chaîne **Complète** et les colore en vert. Notez que `findAll` génère une erreur `ItemNotFound` si la chaîne spécifiée n’existe pas dans la feuille de calcul. Si vous pensez que la chaîne spécifiée peut ne pas exister dans la feuille de calcul, utilisez la méthode[findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) à la place, pour que votre code gère ce scénario plus facilement.

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
> - Pour obtenir des exemples qui montrent comment obtenir une plage dans une feuille de calcul à l’aide de l’objet, voir Obtenir une plage à l’aide de `Range` [l Excel API JavaScript .](excel-add-ins-ranges-get.md)
> - Pour obtenir des exemples qui montrent comment obtenir une plage dans un objet `Table`, reportez-vous à la rubrique [Utiliser des tableaux à l’aide de l’API JavaScript pour Excel](excel-add-ins-tables.md).
> - Pour consulter des exemples qui montrent comment rechercher une grande plage pour plusieurs sous-plages basées sur les caractéristiques de cellule, voir [Travailler avec plusieurs plages simultanément dans des compléments Excel](excel-add-ins-multiple-ranges.md).

## <a name="filter-data"></a>Filtrer les données

Un [filtre automatique](/javascript/api/excel/excel.autofilter) applique des filtres de données sur une plage de cellules dans la feuille de calcul. Il est créé avec `Worksheet.autoFilter.apply` , qui a les paramètres suivants.

- `range`: La plage à laquelle le filtre est appliqué, spécifiée sous la forme d’un`Range` objet ou d’une chaîne.
- `columnIndex`: L’index de colonne de base zéro par rapport à laquelle les critères de filtre sont évaluées.
- `criteria`: Un objet[FilterCriteria](/javascript/api/excel/excel.filtercriteria)afin de déterminer les lignes doivent être filtrées en fonction de la cellule de la colonne.

Le premier exemple de code montre comment ajouter un filtre à la plage utilisée de la feuille de calcul. Ce filtre masque les entrées qui ne sont pas dans les premiers 25%, basé sur les valeurs de colonne **3**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    return context.sync();
}).catch(errorHandlerFunction);
```

L’exemple de code suivant montre comment actualiser le filtre automatique à l’aide de la méthode`reapply`. Cette opération doit être effectuée lorsque les données dans la plage changent.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    return context.sync();
}).catch(errorHandlerFunction);
```

L’exemple de code de filtre automatique final montre comment supprimer le filtre automatique de la feuille de calcul avec la méthode`remove`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    return context.sync();
}).catch(errorHandlerFunction);
```

Un `AutoFilter` peut également être appliqué aux tableaux individuels. Pour plus d’informations, consultez [Utiliser des tableaux avec l’API JavaScript Excel](excel-add-ins-tables.md#autofilter).

## <a name="data-protection"></a>Protection des données

Votre complément permet de contrôler la possibilité qu’a un utilisateur de modifier des données dans une feuille de calcul. La propriété `protection` de la feuille de calcul est un objet [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) avec une méthode `protect()`. L’exemple suivant illustre un scénario de base permettant d’activer/de désactiver la protection complète de la feuille de calcul active.

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

- `options` : objet [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) définissant des restrictions de modification spécifiques.
- `password` : chaîne représentant le mot de passe nécessaire pour qu’un utilisateur puisse ignorer la protection et modifier la feuille de calcul.

L’article [Protéger une feuille de calcul](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) comporte davantage d’informations sur la protection des feuilles de calcul et leur modification via l’interface utilisateur Excel.

## <a name="page-layout-and-print-settings"></a>Mise en page et paramètres d’impression

Les compléments ont accès aux paramètres de mise en page à un niveau de feuille de calcul. Ils contrôlent comment la feuille est imprimée. Un `Worksheet` objet a trois propriétés de mise en page : `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.

`Worksheet.horizontalPageBreaks` et `Worksheet.verticalPageBreaks` sont [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection). Il s’agit de collections de [PageBreaks](/javascript/api/excel/excel.pagebreak), lequel spécifient des plages dans lesquelles les sauts de page manuels sont insérés. Exemple de code suivant ajoute un saut de page horizontal au-dessus de la ligne **21**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    return context.sync();
}).catch(errorHandlerFunction);
```

`Worksheet.pageLayout` est un objet [PageLayout](/javascript/api/excel/excel.pagelayout). Cet objet contient les paramètres de mise en page et impression qui ne dépendent pas d’une implémentation spécifique à l’imprimante. Ces paramètres incluent marges, orientation, numérotation, lignes de titre et zone d’impression.

Exemple de code suivant centre la page (horizontalement et verticalement), définit une ligne de titre qui est imprimée en haut de chaque page et définit la zone imprimée sur une sous-section de la feuille de calcul.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
