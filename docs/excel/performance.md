---
title: Optimisation des performances API JavaScript Excel
description: Optimisez Excel de votre application à l’aide de l’API JavaScript.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5dbaa566138666a049aa5a0c1d940adff056c92e
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745636"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Optimisation des performances à l’aide de l’API JavaScript d’Excel

Il existe plusieurs façons d’effectuer des tâches courantes avec l’API JavaScript Excel. Vous trouverez des différences de performances significatives entre les différentes approches. Cet article fournit des instructions et exemples de code pour vous montrer comment effectuer des tâches courantes efficacement à l’aide des API JavaScript Excel.

> [!IMPORTANT]
> De nombreux problèmes de performances peuvent être résolus par le biais d’une utilisation recommandée et `load` d’appels `sync` . Consultez la section « Améliorations des performances avec les API propres à l’application » des limites de ressources et de l’optimisation des performances pour les Office Pour obtenir des [conseils](../concepts/resource-limits-and-performance-optimization.md#performance-improvements-with-the-application-specific-apis) sur l’utilisation efficace des API propres à l’application.

## <a name="suspend-excel-processes-temporarily"></a>Suspendre temporairement les processus Excel

Excel a des tâches en arrière-plan qui réagissent à l’entrée des utilisateurs et de votre complément. Certains de ces processus Excel peuvent être contrôlés pour accroître les performances. Ceci est particulièrement utile lorsque votre complément utilise de grands ensembles de données.

### <a name="suspend-calculation-temporarily"></a>Suspendre temporairement les calculs

Si vous essayez d’effectuer une opération sur un grand nombre de cellules (par exemple, la définition de la valeur d’un objet plage énorme) et que vous n’avez rien contre la suspension de calcul dans Excel temporairement le temps que votre opération se termine, nous vous recommandons de suspendre le calcul jusqu’à ce que le prochain `context.sync()` soit appelé.

Reportez-vous à la documentation de référence [Objet Application](/javascript/api/excel/excel.application) pour plus d’informations sur l’utilisation de l’API `suspendApiCalculationUntilNextSync()` pour suspendre et réactiver les calculs de manière très pratique. Le code suivant montre comment suspendre temporairement le calcul.

```js
await Excel.run(async (context) => {
    let app = context.workbook.application;
    let sheet = context.workbook.worksheets.getItem("sheet1");
    let rangeToSet: Excel.Range;
    let rangeToGet: Excel.Range;
    app.load("calculationMode");
    await context.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);

    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await context.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await context.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await context.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
});
```

Notez que seuls les calculs de formule sont suspendus. Toutes les références modifiées sont toujours reconstruites. Par exemple, le fait de renommer une feuille de calcul met toujours à jour les références dans les formules de cette feuille de calcul.

### <a name="suspend-screen-updating"></a>Suspendre la mise à jour de l’écran

Excel affiche les modifications effectuées par votre complément à peu près au moment où elles ont lieu dans le code. Dans le cas de grands ensembles de données itératifs, il se peut que vous ne deviez pas afficher cette progression sur l’écran en temps réel. `Application.suspendScreenUpdatingUntilNextSync()` interrompt les mises à jour visuelles vers Excel tant que le complément n’appelle pas `context.sync()`, ou tant que `Excel.run` ne se termine pas (appelant implicitement `context.sync`). N’oubliez pas qu'Excel n’affiche aucun signe d’activité jusqu'à la synchronisation suivante. Votre complément doit donner des conseils aux utilisateurs pour les préparer à ce délai ou fournir une barre d’état pour démontrer l’activité.

> [!NOTE]
> N’appelez pas `suspendScreenUpdatingUntilNextSync` à plusieurs reprises (par exemple, dans une boucle). Les appels répétés entraînent le clignotement Excel fenêtre.

### <a name="enable-and-disable-events"></a>Activation et désactivation d’événements

La performance d’un complément peut être améliorée en désactivant les événements. Un exemple de code montrant comment activer et désactiver les événements dans l’article[manipuler les événements](excel-add-ins-events.md#enable-and-disable-events).

## <a name="importing-data-into-tables"></a>Importation de données dans des tableaux

Lorsque vous tentez d’importer une quantité considérable de données directement dans un objet[tableau](/javascript/api/excel/excel.table) (par exemple, à l’aide de `TableRowCollection.add()`), vous risquez de rencontrer une dégradation des performances. Si vous essayez d’ajouter un nouveau tableau, vous devez remplir les données d’abord en définissant `range.values`, puis appelez `worksheet.tables.add()` pour créer un tableau sur la plage. Si vous essayez d’écrire des données dans un tableau existant, écrivez les données dans un objet plage via `table.getDataBodyRange()`, et le tableau s’agrandit automatiquement.

Voici un exemple de cette approche :

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first.
    let range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    let table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await context.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await context.sync();
});
```

> [!NOTE]
> Vous pouvez convertir un objet de Tableau en objet de Plage à l’aide de la méthode[Table.convertToRange()](/javascript/api/excel/excel.table#excel-excel-table-converttorange-member(1)).

## <a name="payload-size-limit-best-practices"></a>Meilleures pratiques en matière de limite de taille de charge utile

L Excel’API JavaScript présente des limites de taille pour les appels d’API. Excel sur le Web a une limite de taille de charge utile pour les demandes et les réponses de 5 Mo, et une API `RichAPI.Error` renvoie une erreur si cette limite est dépassée. Sur toutes les plateformes, une plage est limitée à cinq millions de cellules pour obtenir des opérations. Les plages importantes dépassent généralement ces deux limitations.

La taille de la charge utile d’une demande est une combinaison des trois composants suivants.

* Nombre d’appels d’API
* Nombre d’objets, tels que des `Range` objets
* Longueur de la valeur à définir ou à obtenir

Si une API renvoie l’erreur `RequestPayloadSizeLimitExceeded` , utilisez les stratégies de meilleures pratiques documentées dans cet article pour optimiser votre script et éviter l’erreur.

### <a name="strategy-1-move-unchanged-values-out-of-loops"></a>Stratégie 1 : déplacer des valeurs inchangées hors des boucles

Limitez le nombre de processus qui se produisent au sein de boucles pour améliorer les performances. Dans l’exemple de code suivant, `context.workbook.worksheets.getActiveWorksheet()` peut être déplacé hors de la `for` boucle, car il ne change pas dans cette boucle.

```js
// DO NOT USE THIS CODE SAMPLE. This sample shows a poor performance strategy. 
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    
    // This sample retrieves the worksheet every time the loop runs, which is bad for performance.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

L’exemple de code suivant présente une logique similaire à l’exemple de code précédent, mais avec une stratégie de performances améliorée. La valeur `context.workbook.worksheets.getActiveWorksheet()` est récupérée avant `for` la boucle, car cette valeur n’a pas besoin d’être `for` récupérée à chaque fois que la boucle s’exécute. Seules les valeurs qui changent dans le contexte d’une boucle doivent être récupérées dans cette boucle.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    let ranges = [];
    // Retrieve the worksheet outside the loop.
    let worksheet = context.workbook.worksheets.getActiveWorksheet(); 

    // Only process the necessary values inside the loop.
    for (let i = 0; i < 7500; i++) {
      let rangeByIndex = worksheet.getRangeByIndexes(i, 1, 1, 1);
    }    
    await context.sync();
  });
}
```

### <a name="strategy-2-create-fewer-range-objects"></a>Stratégie 2 : créer moins d’objets de plage

Créez moins d’objets de plage pour améliorer les performances et réduire la taille de la charge utile. Deux approches pour créer moins d’objets de plage sont décrites dans les sections d’article et les exemples de code suivants.

#### <a name="split-each-range-array-into-multiple-arrays"></a>Fractionner chaque tableau de plages en plusieurs tableaux

Pour créer moins d’objets de plage, vous pouvez fractionner chaque tableau de plages en plusieurs tableaux, puis traiter chaque nouveau tableau avec une boucle et un nouvel `context.sync()` appel.

> [!IMPORTANT]
> Utilisez cette stratégie uniquement si vous avez d’abord déterminé que vous dépassez la limite de taille de demande de charge utile. L’utilisation de boucles multiples peut réduire la taille de chaque demande de charge utile pour éviter de dépasser la limite de 5 Mo, mais l’utilisation de plusieurs boucles `context.sync()` et appels a également un impact négatif sur les performances.

L’exemple de code suivant tente de traiter un grand tableau de plages en une seule boucle, puis un seul `context.sync()` appel. Si vous traitez trop de valeurs de plage dans un `context.sync()` appel, la taille de la demande de charge utile dépasse la limite de 5 Mo.

```js
// This code sample does not show a recommended strategy.
// Calling 10,000 rows would likely exceed the 5MB payload size limit in a real-world situation.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // This sample attempts to process too many ranges at once. 
    for (let row = 1; row < 10000; row++) {
      let range = sheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

L’exemple de code suivant présente une logique similaire à l’exemple de code précédent, mais avec une stratégie qui évite de dépasser la limite de taille de demande de charge utile de 5 Mo. Dans l’exemple de code suivant, les plages sont traitées en deux boucles distinctes, et chaque boucle est suivie d’un `context.sync()` appel.

```js
// This code sample shows a strategy for reducing payload request size.
// However, using multiple loops and `context.sync()` calls negatively impacts performance.
// Only use this strategy if you've determined that you're exceeding the payload request limit.
async function run() {
  await Excel.run(async (context) => {
    let worksheet = context.workbook.worksheets.getActiveWorksheet();

    // Split the ranges into two loops, rows 1-5000 and then 5001-10000.
    for (let row = 1; row < 5000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    // Sync after each loop. 
    await context.sync(); 
    
    for (let row = 5001; row < 10000; row++) {
      let range = worksheet.getRangeByIndexes(row, 1, 1, 1);
      range.values = [["1"]];
    }
    await context.sync(); 
  });
}
```

#### <a name="set-range-values-in-an-array"></a>Définir des valeurs de plage dans un tableau

Une autre façon de créer moins d’objets de plage consiste à créer un tableau, à utiliser une boucle pour définir toutes les données de ce tableau, puis à transmettre les valeurs du tableau à une plage. Cela bénéficie à la fois des performances et de la taille de la charge utile. Au lieu d’appeler `range.values` chaque plage d’une boucle, `range.values` est appelé une fois en dehors de la boucle.

L’exemple de code suivant montre comment créer un tableau, `for` définir les valeurs de ce tableau dans une boucle, puis passer les valeurs du tableau à une plage en dehors de la boucle.

```js
// This code sample shows a good performance strategy.
async function run() {
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();    
    // Create an array.
    const array = new Array(10000);

    // Set the values of the array inside the loop.
    for (let i = 0; i < 10000; i++) {
      array[i] = [1];
    }

    // Pass the array values to a range outside the loop. 
    let range = worksheet.getRange("A1:A10000");
    range.values = array;
    await context.sync();
  });
}
```

## <a name="see-also"></a>Voir aussi

* [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
* [Gestion des erreurs avec l Excel API JavaScript](excel-add-ins-error-handling.md)
* [Limites des ressources et optimisation des performances pour les compléments Office](../concepts/resource-limits-and-performance-optimization.md)
* [Objet de fonctions de feuille de calcul (API JavaScript pour Excel)](/javascript/api/excel/excel.functions)
