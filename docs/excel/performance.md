---
title: Optimisation des performances de l'API JavaScript d'Excel
description: Optimiser les performances à l'aide de l'API JavaScript d'Excel
ms.date: 03/28/2018
ms.openlocfilehash: dabbb69f8dee0df782a265edcfdfb1c89894e915
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437408"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Optimisation des performances à l'aide de l'API JavaScript d'Excel

Il y a plusieurs manières d'effectuer des tâches courantes avec l'API JavaScript d'Excel. Vous trouverez des différences de performances significatives entre les diverses approches. Cet article fournit de l'aide et des exemples de code pour vous montrer comment effectuer efficacement des tâches courantes en utilisant l'API JavaScript d'Excel.

## <a name="minimize-the-number-of-sync-calls"></a>Réduisez le nombre d'appels à sync()

Dans l'API JavaScript d'Excel, ```sync()``` est la seule opération asynchrone, et elle peut être lente dans certaines circonstances, en particulier pour Excel Online. Pour optimiser les performances, réduisez le nombre d'appels à ```sync()``` en mettant en file d'attente autant de changements que possible avant de l'appeler.

Voir [Concepts de base - sync()](excel-add-ins-core-concepts.md#sync) pour des exemples de code qui suivent cette pratique.

## <a name="minimize-the-number-of-proxy-objects-created"></a>Réduisez le nombre d'objets proxy créés

Évitez de créer répétitivement le même objet proxy. A la place, si vous avez besoin du même objet proxy pour plus d'une opération, créez-le une fois et affectez-le à une variable, puis utilisez cette variable dans votre code.

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

## <a name="load-necessary-properties-only"></a>Ne chargez que les propriétés nécessaires

Dans l'API JavaScript d'Excel, vous devez charger explicitement les propriétés d'un objet proxy. Bien que vous puissiez charger toutes les propriétés en une fois avec un appel vide à ```load()```, cette approche peut avoir un surcoût significatif en termes de performances. A la place, nous vous suggérons de ne charger que les propriétés nécessaires, en particulier pour ceux des objets qui ont un nombre important de propriétés.

Par exemple, si vous ne souhaitez relire que la propriété **address** d’un objet plage, indiquez seulement cette propriété lorsque vous appelez la méthode **load()** :
 
```js
range.load('address');
```
 
Vous pouvez appeler la méthode **load()** de l’une quelconque des façons suivantes :
 
_Syntaxe :_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_Où :_
 
* `properties` est la liste des propriétés à charger, spécifiée comme des chaînes délimitées par des virgules ou comme un tableau de noms. Pour plus d’informations, voir les méthodes **load()** définies pour les objets dans la [Référence de l’API JavaScript d'Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).
* `loadOption` spécifie un objet qui décrit les options selection, expansion, top et skip. Voir les [options](https://dev.office.com/reference/add-ins/excel/loadoption) de chargement d’objet pour les détails.

SVP, soyez conscient que certaines des "propriétés" dans un objet peuvent avoir le même nom qu'un autre objet. Par exemple, `format` est une propriété dans l'objet plage, mais `format` lui-même est un objet aussi. Donc, si vous faites un appel tel que `range.load("format")`, c'est équivalent à `range.format.load()`, qui est un appel vide à load() qui peut engendrer des problèmes de performances comme résumé précédemment. Pour éviter cela, votre code ne devrait charger que les "nœuds feuilles" dans une arborescence d'objets. 

## <a name="suspend-calculation-temporarily"></a>Suspendre le calcul temporairement

Si vous essayez d'effectuer une opération sur un grand nombre de cellules (par exemple, en définissant la valeur d'un objet plage très volumineux) et que cela ne vous dérange pas de suspendre temporairement le calcul dans Excel jusqu'à ce que votre opération se termine, nous vous recommandons de suspendre le calcul jusqu'à ce que le prochain ```context.sync()``` soit appelé.

Voir la documentation de référence de l'[Objet Application](https://dev.office.com/reference/add-ins/excel/application) pour des informations sur la façon d'utiliser l'```suspendApiCalculationUntilNextSync()``` API pour suspendre et réactiver les calculs d'une manière très pratique. Le code suivant montre comment suspendre le calcul temporairement :

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

## <a name="update-all-cells-in-a-range"></a>Mettre à jour toutes les cellules d’une plage 

Lorsque vous devez mettre à jour toutes les cellules dans une plage avec la même valeur ou propriété, il peut être lent de le faire via un tableau bidimensionnel qui indique répétitivement la même valeur, car cette approche nécessite qu'Excel parcoure toutes les cellules de la plage pour les définir individuellement. Excel a un moyen plus efficace pour mettre à jour toutes les cellules dans une plage avec la même valeur ou propriété.

Si vous devez appliquer la même valeur, le même format numérique ou la même formule à une plage de cellules, il est plus efficace de spécifier une seule valeur au lieu d'un tableau de valeurs. Procéder ainsi améliorera significativement les performances. Pour un exemple de code qui montre cette approche en action, voir [Concepts de base - Mettre à jour toutes les cellules d'une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Un scénario courant dans lequel vous pouvez appliquer cette approche est la définition de formats numériques différents pour des colonnes différentes dans une feuille de calcul. Dans ce cas, vous pouvez simplement parcourir les colonnes et définir le format numérique pour chaque colonne avec une seule valeur. Manipuler chaque colonne comme une plage, comme indiqué dans l'exemple de code [Mettre à jour toutes les cellules dans une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

> [!NOTE]
> Si vous utilisez TypeScript, vous remarquerez une erreur de compilation indiquant qu'une valeur unique ne peut pas être affectée à un tableau 2D.  C'est inévitable du fait que les valeurs *sont* un tableau 2D lors de la récupération des propriétés, et que TypeScript n'autorise pas des types différents pour un setter et un getter.  Cependant, un contournement simple consiste à définir les valeurs avec un suffixe `as any`, par exemple, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Importation de données dans des tables

Lorsque vous essayez d'importer un très grand volume de données directement dans un objet[Table](https://dev.office.com/reference/add-ins/excel/table) (par exemple, en utilisant `TableRowCollection.add()`), vous risquez de subir une performance lente. Si vous essayez d'ajouter une nouvelle table, vous devriez d'abord remplir les données en définissant `range.values`, puis appeler alors `worksheet.tables.add()` pour créer une table sur la plage. Si vous essayez d'écrire des données dans une table existante, écrivez les données dans un objet plage via `table.getDataBodyRange()`, et la table s'agrandira automatiquement. 

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
> Vous pouvez aisément convertir un objet Table en objet Range en utilisant la méthode [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange).

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l’API JavaScript d'Excel](excel-add-ins-core-concepts.md)
- [Concepts avancés de l’API JavaScript d'Excel](excel-add-ins-advanced-concepts.md)
- [Spécification ouverte de l’API JavaScript d'Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Objet de fonctions de feuille de calcul (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/functions)
