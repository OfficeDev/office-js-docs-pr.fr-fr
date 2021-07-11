---
title: Optimisation des performances API JavaScript Excel
description: Optimisez Excel de votre application à l’aide de l’API JavaScript.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 5313bb3fe25d165e49cc0508e81d58294db48798
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349384"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Optimisation des performances à l’aide de l’API JavaScript d’Excel

Il existe plusieurs façons d’effectuer des tâches courantes avec l’API JavaScript Excel. Vous trouverez des différences de performances significatives entre les différentes approches. Cet article fournit des instructions et exemples de code pour vous montrer comment effectuer des tâches courantes efficacement à l’aide des API JavaScript Excel.

> [!IMPORTANT]
> De nombreux problèmes de performances peuvent être résolus par le biais d’une utilisation recommandée `load` et `sync` d’appels. Consultez la section « Améliorations des performances avec les API propres à l’application » des limites de ressources et de l’optimisation des performances pour les Office Pour obtenir des [conseils](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) sur l’utilisation efficace des API propres à l’application.

## <a name="suspend-excel-processes-temporarily"></a>Suspendre temporairement les processus Excel

Excel a des tâches en arrière-plan qui réagissent à l’entrée des utilisateurs et de votre complément. Certains de ces processus Excel peuvent être contrôlés pour accroître les performances. Ceci est particulièrement utile lorsque votre complément utilise de grands ensembles de données.

### <a name="suspend-calculation-temporarily"></a>Suspendre temporairement les calculs

Si vous essayez d’effectuer une opération sur un grand nombre de cellules (par exemple, la définition de la valeur d’un objet plage énorme) et que vous n’avez rien contre la suspension de calcul dans Excel temporairement le temps que votre opération se termine, nous vous recommandons de suspendre le calcul jusqu’à ce que le prochain `context.sync()` soit appelé.

Reportez-vous à la documentation de référence [Objet Application](/javascript/api/excel/excel.application) pour plus d’informations sur l’utilisation de l’API `suspendApiCalculationUntilNextSync()` pour suspendre et réactiver les calculs de manière très pratique. Le code suivant montre comment suspendre temporairement le calcul.

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

Notez que seuls les calculs de formule sont suspendus. Toutes les références modifiées sont toujours reconstruites. Par exemple, le fait de renommer une feuille de calcul met toujours à jour les références dans les formules de cette feuille de calcul.

### <a name="suspend-screen-updating"></a>Suspendre la mise à jour de l’écran

Excel affiche les modifications effectuées par votre complément à peu près au moment où elles ont lieu dans le code. Dans le cas de grands ensembles de données itératifs, il se peut que vous ne deviez pas afficher cette progression sur l’écran en temps réel. `Application.suspendScreenUpdatingUntilNextSync()` interrompt les mises à jour visuelles vers Excel tant que le complément n’appelle pas `context.sync()`, ou tant que `Excel.run` ne se termine pas (appelant implicitement `context.sync`). N’oubliez pas qu'Excel n’affiche aucun signe d’activité jusqu'à la synchronisation suivante. Votre complément doit donner des conseils aux utilisateurs pour les préparer à ce délai ou fournir une barre d’état pour démontrer l’activité.

> [!NOTE]
> N’appelez `suspendScreenUpdatingUntilNextSync` pas à plusieurs reprises (par exemple, dans une boucle). Les appels répétés entraînent le clignotement Excel fenêtre.

### <a name="enable-and-disable-events"></a>Activation et désactivation d’événements

La performance d’un complément peut être améliorée en désactivant les événements. Un exemple de code montrant comment activer et désactiver les événements dans l’article[manipuler les événements](excel-add-ins-events.md#enable-and-disable-events).

## <a name="importing-data-into-tables"></a>Importation de données dans des tableaux

Lorsque vous tentez d’importer une quantité considérable de données directement dans un objet[tableau](/javascript/api/excel/excel.table) (par exemple, à l’aide de `TableRowCollection.add()`), vous risquez de rencontrer une dégradation des performances. Si vous essayez d’ajouter un nouveau tableau, vous devez remplir les données d’abord en définissant `range.values`, puis appelez `worksheet.tables.add()` pour créer un tableau sur la plage. Si vous essayez d’écrire des données dans un tableau existant, écrivez les données dans un objet plage via `table.getDataBodyRange()`, et le tableau s’agrandit automatiquement.

Voici un exemple de cette approche :

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
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

## <a name="see-also"></a>Voir aussi

* [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
* [Limites des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md)
* [Objet de fonctions de feuille de calcul (API JavaScript pour Excel)](/javascript/api/excel/excel.functions)
