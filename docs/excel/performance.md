---
title: Optimisation des performances API JavaScript Excel
description: Optimisation des performances à l’aide de l’API JavaScript d’Excel
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: f48b62b47c4000b128043fe2e01f949af7179e73
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872136"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Optimisation des performances à l’aide de l’API JavaScript d’Excel

Il existe plusieurs façons d’effectuer des tâches courantes avec l’API JavaScript Excel. Vous trouverez des différences de performances significatives entre les différentes approches. Cet article fournit des instructions et exemples de code pour vous montrer comment effectuer des tâches courantes efficacement à l’aide des API JavaScript Excel.

## <a name="minimize-the-number-of-sync-calls"></a>Limitez le nombre d’appels sync()

Dans l’API JavaScript Excel, ```sync()``` est la seule opération asynchrone et elle peut être lente dans certaines circonstances, en particulier pour Excel Online. Pour optimiser les performances, vous devez limiter le nombre de fois que vous appelez ```sync()``` et mettre en file d’attente autant de modifications que possible avant d’appeler.

Voir [Concepts principaux - sync()](excel-add-ins-core-concepts.md#sync) pour des exemples de code qui suivent cette pratique.

## <a name="minimize-the-number-of-proxy-objects-created"></a>Réduire le nombre d’objets proxy créés

Éviter de créer le même objet proxy à plusieurs reprises. Au lieu de cela, si vous avez besoin du même objet proxy pour plus d’une opération, créez-le une seule fois et affectez-le à une variable, puis utilisez cette variable dans votre code.

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

## <a name="load-necessary-properties-only"></a>Charger les propriétés nécessaires uniquement

Dans l’API JavaScript Excel, vous devez charger explicitement les propriétés d’un objet proxy. Bien que vous soyez en mesure de charger les propriétés en une fois avec un appel vide```load()```, cette approche peut causer une surcharge significative des performances. Au lieu de cela, nous vous conseillons de charger uniquement les propriétés nécessaires, en particulier pour ces objets qui ont un grand nombre de propriétés.

Par exemple, si vous souhaitez uniquement lire la propriété **adresse** d’un objet de la plage, spécifiez uniquement cette propriété lorsque vous appelez la méthode **load()**  :

```js
range.load('address');
```

Vous pouvez appeler la méthode **load()** de l’une des façons suivantes :

_Syntaxe :_

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

_Où :_

* `properties` est la liste des propriétés à charger, fournie sous forme de chaînes séparées par des virgules ou de tableau de noms. Pour plus d’informations, reportez-vous aux méthodes **load()** définies pour les objets dans la rubrique [Référence de l’API JavaScript pour Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).
* `loadOption` spécifie un objet qui décrit les options select, expand, top et skip. Pour plus d’informations, reportez-vous aux [options](/javascript/api/office/officeextension.loadoption) de chargement d’objet.

N’oubliez pas que certaines des « propriétés » sous un objet peuvent avoir le même nom qu’un autre objet. Par exemple, `format` est une propriété sous plage d’objet, mais `format` lui-même est également un objet. Par conséquent, si vous passez un appel comme `range.load("format")`, cela équivaut à `range.format.load()`, c'est-à-dire, un appel load() vide pouvant entraîner des problèmes de performances comme indiqué précédemment. Pour éviter cela, votre code devrait charger uniquement les nœuds « terminaux » dans une arborescence d’objets. 

## <a name="suspend-excel-processes-temporarily"></a>Suspendre temporairement les processus Excel

Excel a des tâches en arrière-plan qui réagissent à l’entrée des utilisateurs et de votre complément. Certains de ces processus Excel peuvent être contrôlés pour accroître les performances. Ceci est particulièrement utile lorsque votre complément utilise de grands ensembles de données.

### <a name="suspend-calculation-temporarily"></a>Suspendre temporairement les calculs

Si vous essayez d’effectuer une opération sur un grand nombre de cellules (par exemple, la définition de la valeur d’un objet plage énorme) et que vous n’avez rien contre la suspension de calcul dans Excel temporairement le temps que votre opération se termine, nous vous recommandons de suspendre le calcul jusqu’à ce que le prochain `context.sync()` soit appelé.

Reportez-vous à la documentation de référence [Objet Application](/javascript/api/excel/excel.application) pour plus d’informations sur l’utilisation de l’API `suspendApiCalculationUntilNextSync()` pour suspendre et réactiver les calculs de manière très pratique. Le code suivant montre comment suspendre temporairement le calcul :

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

### <a name="suspend-screen-updating"></a>Suspendre la mise à jour de l’écran

> [!NOTE]
> La méthode`suspendScreenUpdatingUntilNextSync`décrit dans cet article est actuellement disponible uniquement dans la version d’affichage publique. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Excel affiche les modifications effectuées par votre complément à peu près au moment où elles ont lieu dans le code. Dans le cas de grands ensembles de données itératifs, il se peut que vous ne deviez pas afficher cette progression sur l’écran en temps réel. `Application.suspendScreenUpdatingUntilNextSync()` interrompt les mises à jour visuelles vers Excel tant que le complément n’appelle pas `context.sync()`, ou tant que `Excel.run` ne se termine pas (appelant implicitement `context.sync`). N’oubliez pas qu'Excel n’affiche aucun signe d’activité jusqu'à la synchronisation suivante. Votre complément doit donner des conseils aux utilisateurs pour les préparer à ce délai ou fournir une barre d’état pour démontrer l’activité.

### <a name="enable-and-disable-events"></a>Activation et désactivation d’événements

La performance d’un complément peut être améliorée en désactivant les événements. Un exemple de code montrant comment activer et désactiver les événements dans l’article[manipuler les événements](excel-add-ins-events.md#enable-and-disable-events).

## <a name="update-all-cells-in-a-range"></a>Mettre à jour toutes les cellules d’une plage

Lorsque vous devez mettre à jour toutes les cellules dans une plage avec la même valeur ou propriété, il peut être lent de le faire via une matrice 2 dimensions indiquant à plusieurs reprises la même valeur étant donné que cette approche nécessite qu’Excel le répète sur toutes les cellules dans la plage pour définir chacune séparément. Excel propose une méthode plus efficace pour mettre à jour toutes les cellules dans une plage avec la même valeur ou propriété.

Si vous voulez appliquer la même valeur, le même format de nombre ou la même formule à une plage de cellules, il est plus efficace de spécifier une valeur unique au lieu d’une matrice de valeurs. Cette opération va améliorer sensiblement les performances. Pour voir un exemple de code indiquant cette approche en action, [principaux concepts - mettre à jour toutes les cellules dans une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Un scénario classique où vous pouvez appliquer cette approche est lors de la définition de différents formats numériques différents sur différentes colonnes dans une feuille de calcul. Dans ce cas, vous pouvez simplement itérer dans les colonnes et définir le format de nombre dans chaque colonne avec une valeur unique. Traiter chaque colonne comme une plage, comme illustré dans l’exemple de code[mettre à jour toutes les cellules dans une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

> [!NOTE]
> Si vous utilisez TypeScript, vous remarquerez une erreur de compilation indiquant qu’une seule valeur ne peut pas être définie à une matrice 2D.  Ceci est inévitable puisque les valeurs *sont* un tableau 2D qui extrait les propriétés et TypeScript n’autorise pas de types différents pour configurer et récolter.  Toutefois, une solution de contournement simple consiste à définir les valeurs avec un `as any` suffixe, par exemple, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Importation de données dans des tableaux

Lorsque vous tentez d’importer une quantité considérable de données directement dans un objet[tableau](/javascript/api/excel/excel.table) (par exemple, à l’aide de `TableRowCollection.add()`), vous risquez de rencontrer une dégradation des performances. Si vous essayez d’ajouter un nouveau tableau, vous devez remplir les données d’abord en définissant `range.values`, puis appelez `worksheet.tables.add()` pour créer un tableau sur la plage. Si vous essayez d’écrire des données dans un tableau existant, écrivez les données dans un objet plage via `table.getDataBodyRange()`, et le tableau s’agrandit automatiquement. 

Voici un exemple de cette approche :

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
> Vous pouvez convertir un objet de Tableau en objet de Plage à l’aide de la méthode[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).

## <a name="untrack-unneeded-ranges"></a>Annuler le suivi des plages inutiles

La couche JavaScript crée des objets proxy pour votre complément pour interagir avec le classeur Excel et les sous-jacentes. Ces objets sont conservés en mémoire jusqu'à `context.sync()` soit appelé. Les opérations par lots volumineux peuvent générer un grand nombre d’objets proxy qui sont uniquement utiles une fois pour le complément et peuvent être publiés à partir de la mémoire avant l’exécution du lot.

La méthode [Range.untrack()](/javascript/api/excel/excel.range#untrack--) libère un objet plage Excel à partir de la mémoire. Appeler cette méthode une fois que votre complément a terminé avec la plage doit créer une amélioration notable des performances lors de l’utilisation d’un grand nombre d’objets de plage.

> [!NOTE]
> `Range.untrack()` est un raccourci pour [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-). N’importe quel objet proxy peut être non suivi en le supprimant de la liste d’objets suivis dans le contexte. En règle générale, les objets Plage sont les seuls objets Excel utilisés dans une quantité suffisante pour justifier le non suivi.

L’exemple de code suivant remplit une plage sélectionnée avec des données, une cellule à la fois. Une fois que la valeur est ajoutée à la cellule, la plage représentant cette cellule est non suivie. Exécuter tout d’abord ce code avec une plage sélectionnée de 10 000 à 20 000 cellules, avec la `cell.untrack()` ligne et puis sans. Vous devez remarquer que le code est exécuté plus rapidement avec la `cell.untrack()` ligne que sans elle. Vous pouvez également remarquer un temps de réponse plus rapide par la suite, étant donné que l’étape de nettoyage prend moins de temps.

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

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Concepts avancés de programmation avec l’API JavaScript Excel](excel-add-ins-advanced-concepts.md)
- [Limites des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md)
- [Spécification d’ouverture d’API JavaScript pour Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Objet de fonctions de feuille de calcul (API JavaScript pour Excel)](/javascript/api/excel/excel.functions)
