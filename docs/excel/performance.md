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
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Optimisation des performances ? l'aide de l'API JavaScript d'Excel

Il y a plusieurs mani?res d'effectuer des t?ches courantes avec l'API JavaScript d'Excel. Vous trouverez des diff?rences de performances significatives entre les diverses approches. Cet article fournit de l'aide et des exemples de code pour vous montrer comment effectuer efficacement des t?ches courantes en utilisant l'API JavaScript d'Excel.

## <a name="minimize-the-number-of-sync-calls"></a>R?duisez le nombre d'appels ? sync()

Dans l'API JavaScript d'Excel, ```sync()``` est la seule op?ration asynchrone, et elle peut ?tre lente dans certaines circonstances, en particulier pour Excel Online. Pour optimiser les performances, r?duisez le nombre d'appels ? ```sync()``` en mettant en file d'attente autant de changements que possible avant de l'appeler.

Voir [Concepts de base - sync()](excel-add-ins-core-concepts.md#sync) pour des exemples de code qui suivent cette pratique.

## <a name="minimize-the-number-of-proxy-objects-created"></a>R?duisez le nombre d'objets proxy cr??s

?vitez de cr?er r?p?titivement le m?me objet proxy. A la place, si vous avez besoin du m?me objet proxy pour plus d'une op?ration, cr?ez-le une fois et affectez-le ? une variable, puis utilisez cette variable dans votre code.

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

## <a name="load-necessary-properties-only"></a>Ne chargez que les propri?t?s n?cessaires

Dans l'API JavaScript d'Excel, vous devez charger explicitement les propri?t?s d'un objet proxy. Bien que vous puissiez charger toutes les propri?t?s en une fois avec un appel vide ? ```load()```, cette approche peut avoir un surco?t significatif en termes de performances. A la place, nous vous sugg?rons de ne charger que les propri?t?s n?cessaires, en particulier pour ceux des objets qui ont un nombre important de propri?t?s.

Par exemple, si vous ne souhaitez relire que la propri?t? **address** d?un objet plage, indiquez seulement cette propri?t? lorsque vous appelez la m?thode **load()** :
 
```js
range.load('address');
```
 
Vous pouvez appeler la m?thode **load()** de l?une quelconque des fa?ons suivantes :
 
_Syntaxe :_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_O? :_
 
* `properties` est la liste des propri?t?s ? charger, sp?cifi?e comme des cha?nes d?limit?es par des virgules ou comme un tableau de noms. Pour plus d?informations, voir les m?thodes **load()** d?finies pour les objets dans la [R?f?rence de l?API JavaScript d'Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).
* `loadOption` sp?cifie un objet qui d?crit les options selection, expansion, top et skip. Voir les [options](https://dev.office.com/reference/add-ins/excel/loadoption) de chargement d?objet pour les d?tails.

SVP, soyez conscient que certaines des "propri?t?s" dans un objet peuvent avoir le m?me nom qu'un autre objet. Par exemple, `format` est une propri?t? dans l'objet plage, mais `format` lui-m?me est un objet aussi. Donc, si vous faites un appel tel que `range.load("format")`, c'est ?quivalent ? `range.format.load()`, qui est un appel vide ? load() qui peut engendrer des probl?mes de performances comme r?sum? pr?c?demment. Pour ?viter cela, votre code ne devrait charger que les "n?uds feuilles" dans une arborescence d'objets. 

## <a name="suspend-calculation-temporarily"></a>Suspendre le calcul temporairement

Si vous essayez d'effectuer une op?ration sur un grand nombre de cellules (par exemple, en d?finissant la valeur d'un objet plage tr?s volumineux) et que cela ne vous d?range pas de suspendre temporairement le calcul dans Excel jusqu'? ce que votre op?ration se termine, nous vous recommandons de suspendre le calcul jusqu'? ce que le prochain ```context.sync()``` soit appel?.

Voir la documentation de r?f?rence de l'[Objet Application](https://dev.office.com/reference/add-ins/excel/application) pour des informations sur la fa?on d'utiliser l'```suspendApiCalculationUntilNextSync()``` API pour suspendre et r?activer les calculs d'une mani?re tr?s pratique. Le code suivant montre comment suspendre le calcul temporairement :

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

## <a name="update-all-cells-in-a-range"></a>Mettre ? jour toutes les cellules d?une plage 

Lorsque vous devez mettre ? jour toutes les cellules dans une plage avec la m?me valeur ou propri?t?, il peut ?tre lent de le faire via un tableau bidimensionnel qui indique r?p?titivement la m?me valeur, car cette approche n?cessite qu'Excel parcoure toutes les cellules de la plage pour les d?finir individuellement. Excel a un moyen plus efficace pour mettre ? jour toutes les cellules dans une plage avec la m?me valeur ou propri?t?.

Si vous devez appliquer la m?me valeur, le m?me format num?rique ou la m?me formule ? une plage de cellules, il est plus efficace de sp?cifier une seule valeur au lieu d'un tableau de valeurs. Proc?der ainsi am?liorera significativement les performances. Pour un exemple de code qui montre cette approche en action, voir [Concepts de base - Mettre ? jour toutes les cellules d'une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Un sc?nario courant dans lequel vous pouvez appliquer cette approche est la d?finition de formats num?riques diff?rents pour des colonnes diff?rentes dans une feuille de calcul. Dans ce cas, vous pouvez simplement parcourir les colonnes et d?finir le format num?rique pour chaque colonne avec une seule valeur. Manipuler chaque colonne comme une plage, comme indiqu? dans l'exemple de code [Mettre ? jour toutes les cellules dans une plage](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

> [!NOTE]
> Si vous utilisez TypeScript, vous remarquerez une erreur de compilation indiquant qu'une valeur unique ne peut pas ?tre affect?e ? un tableau 2D.  C'est in?vitable du fait que les valeurs *sont* un tableau 2D lors de la r?cup?ration des propri?t?s, et que TypeScript n'autorise pas des types diff?rents pour un setter et un getter.  Cependant, un contournement simple consiste ? d?finir les valeurs avec un suffixe `as any`, par exemple, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Importation de donn?es dans des tables

Lorsque vous essayez d'importer un tr?s grand volume de donn?es directement dans un objet[Table](https://dev.office.com/reference/add-ins/excel/table) (par exemple, en utilisant `TableRowCollection.add()`), vous risquez de subir une performance lente. Si vous essayez d'ajouter une nouvelle table, vous devriez d'abord remplir les donn?es en d?finissant `range.values`, puis appeler alors `worksheet.tables.add()` pour cr?er une table sur la plage. Si vous essayez d'?crire des donn?es dans une table existante, ?crivez les donn?es dans un objet plage via `table.getDataBodyRange()`, et la table s'agrandira automatiquement. 

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
> Vous pouvez ais?ment convertir un objet Table en objet Range en utilisant la m?thode [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange).

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l?API JavaScript d'Excel](excel-add-ins-core-concepts.md)
- [Concepts avanc?s de l?API JavaScript d'Excel](excel-add-ins-advanced-concepts.md)
- [Sp?cification ouverte de l?API JavaScript d'Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Objet de fonctions de feuille de calcul (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/functions)
