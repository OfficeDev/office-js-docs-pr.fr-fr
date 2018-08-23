---
title: Utilisation de tableaux croisés dynamiques à l’aide de l’API JavaScript pour Excel
description: Utilisez l'API JavaScript pour Excel afin de créer des tableaux croisés dynamiques et d’interagir avec leurs composants.
ms.date: 08/17/2018
ms.openlocfilehash: aa6da2e82ab9b0c255208a86012d51db77982934
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/20/2018
ms.locfileid: "22493953"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Utilisation de tableaux croisés dynamiques à l’aide de l’API JavaScript pour Excel

Les tableaux croisés dynamiques rationalisent les jeux de données plus volumineux. Ils permettent la manipulation rapide des données groupées. L'API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d'interagir avec leurs composants. 

Si vous ne connaissez pas les fonctionnalités des tableaux croisés dynamiques, envisagez de les découvrir en tant qu’utilisateur final. Consultez [Créer un tableau croisé dynamique pour analyser les données d'une feuille de calcul](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) afin d'obtenir une présentation de ces outils. 

Cet article fournit des exemples de code pour des scénarios courants. [Excel OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel) fournit une documentation de référence complète pour cette fonction de préversion. 

Pour améliorer votre compréhension de l'API Tableau croisé dynamique, consultez [**Tableau croisé dynamique**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottable.md) et [**Collection Tableau croisé dynamique**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottablecollection.md).

> [!NOTE]
> Ces exemples utilisent des API uniquement disponibles dans la préversion publique (bêta) actuellement. Ces exemples nécessitent l'exécution de préversions. Utilisez la bibliothèque bêta de [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ou participez au [programme Office Insider](https://products.office.com/office-insider). Les fonctionnalités du tableau croisé dynamique sont actuellement disponibles dans la version 16.0.10801.20004.

## <a name="hierarchies"></a>Hiérarchies

Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie : ligne, colonne, données et filtre. Les données suivantes décrivant des ventes de fruits provenant de différentes fermes seront utilisées dans cet article.

![Collection de ventes de fruits de différents types provenant de différentes fermes.](../images/excel-pivots-raw-data.png)

Ces données ont cinq hiérarchies : **Fermes**, **Type**, **Classification**, **Caisses vendues à la ferme** et **Caisses vendues en gros**. Chaque hiérarchie ne peut exister que dans l’une des quatre catégories. Si le **Type** est ajouté aux hiérarchies de colonnes puis aux hiérarchies de lignes, il ne reste que dans ces dernières.

Les hiérarchies de lignes et de colonnes définissent la façon dont les données sont regroupées. Par exemple, une hiérarchie de lignes de **Fermes** regroupe tous les jeux de données provenant de la même ferme. Le choix entre la hiérarchie de lignes et de colonnes définit l'orientation du tableau croisé dynamique.

Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de lignes et de colonnes. Un tableau croisé dynamique avec une hiérarchie de lignes de **Fermes** et une hiérarchie de données de **Caisses vendues en gros** affiche la somme totale (par défaut) de tous les différents fruits pour chaque ferme.

Les hiérarchies de filtres incluent ou excluent les données provenant du pivot en fonction des valeurs dans ce type filtré. Une hiérarchie de filtres de **Classification** avec le type **Biologique** sélectionné n'affiche que les données pour les fruits biologiques.

Voici à nouveau les données des fermes, à côté d’un tableau croisé dynamique. Le tableau croisé dynamique utilise **Ferme** et **Type** en tant que hiérarchies de lignes, **Caisses vendues à la ferme** et **Caisses vendues en gros** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut de la somme) et **Classification** en tant que hiérarchie de filtres (avec **Biologique** sélectionné). 

![Sélection de données de ventes de fruits à côté d'un tableau croisé dynamique avec les hiérarches de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

Ce tableau croisé dynamique peut être généré via l’API JavaScript ou l’interface utilisateur d’Excel. Les deux options permettent une manipulation plus poussée via les compléments.

## <a name="create-a-pivottable"></a>Créer un tableau croisé dynamique

Les tableaux croisés dynamiques nécessitent un nom, une source et une destination. La source peut être une adresse de plage ou un nom de table (passés comme un type `Range`, `string` ou `Table`). La destination est une adresse de plage (donnée sous forme de `Range` ou `string`). Les exemples suivants présentent différentes techniques de création de tableau croisé dynamique.

### <a name="create-a-pivottable-with-range-addresses"></a>Créer un tableau croisé dynamique avec des adresses de plages

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Créer un tableau croisé dynamique avec des objets Plage

```typescript
await Excel.run(async (context) => {    
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add("Farm Sales", rangeToAnalyze, rangeToPlacePivot);
    
    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>Créer un tableau croisé dynamique au niveau classeur

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Utiliser un tableau croisé dynamique existant

Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection Tableau croisé dynamique du classeur ou des feuilles de calcul individuelles. 

Le code suivant récupère le premier tableau croisé dynamique du classeur. Il donne ensuite un nom à la table pour faciliter les références ultérieures.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Ajouter des lignes et des colonnes à un tableau croisé dynamique

Les lignes et les colonnes regroupent les données autour des valeurs de ces champs.

L'ajout de la colonne **Ferme** regroupe toutes les ventes autour de chaque ferme. L'ajout des lignes **Type** et **Classification** décompose davantage les données en fonction du fruit vendu et de sa classification biologique ou non.

![Tableau croisé dynamique avec une colonne Ferme et des lignes Type et Classification.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

Un tableau croisé dynamique peut également ne contenir que des lignes ou que des colonnes.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Ajouter des hiérarchies de données au tableau croisé dynamique

Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes. L'ajout des hiérarchies de données de **Caisses vendues à la ferme** et **Caisses vendues en gros** donne les sommes de ces chiffres pour chaque ligne et chaque colonne. 

Dans l’exemple, **Ferme** et **Type** sont des lignes, tandis que les ventes de caisses sont les données. 

![Tableau croisé dynamique affichant les ventes totales des différents fruits en fonction de la ferme d'où ils proviennent.](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the heirarchies that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a>Modifier la fonction d’agrégation

Les hiérarchies de données voient leurs valeurs agrégées. Pour les jeux de données de nombres, il s’agit d’une somme par défaut. Le `summarizeBy` propriété définit ce comportement en fonction d'un `AggregrationFunction` type. 

Les types de fonctions d’agrégation actuellement prises en charge sont `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` et `Automatic` (par défaut).

Les exemples de code suivants modifient l'agrégation en moyennes des données.

```typescript
    await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        pivotTable.dataHierarchies.load("no-properties-needed");
        await context.sync();

        // changing the aggregation from the default sum to an average of all the values in the hierarchy
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        await context.sync();
    });
```

## <a name="pivottable-layouts"></a>Dispositions des tableaux croisés dynamiques

La disposition d'un tableau croisé dynamique définit le positionnement des hiérarchies et de leurs données. Accéder à la disposition permet de déterminer les plages de stockage des données. 

Le diagramme suivant présente la correspondance des appels de fonction de disposition avec les plages du tableau croisé dynamique.

![Diagramme présentant les sections d’un tableau croisé dynamique renvoyées par les fonctions de récupération de plage de la disposition.](../images/excel-pivots-layout-breakdown.png)

Le code suivant indique comment récupérer la dernière ligne des données de tableau croisé dynamique via la disposition. Ces valeurs sont ensuite additionnées pour obtenir un total général.


```typescript
    await Excel.run(async (context) => {
        const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
        
        // get the totals for each data hierarchy from the layout
        const range = pivotTable.layout.getDataBodyRange();
        const grandTotalRange = range.getLastRow();
        grandTotalRange.load("address");
        await context.sync();
        
        // sum the totals from the PivotTable data hierarchies and place them in a new range
        const masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
        masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
        await context.sync();
    });
```

Les tableaux croisés dynamiques ont trois styles de disposition : Compact, Plan et Tabulaire. Nous avons vu le style compact dans les exemples précédents. 

Les exemples suivants utilisent respectivement le style plan et tabulaire. L’exemple de code montre comment passer d'une disposition à une autre.

### <a name="outline-layout"></a>Disposition Plan

![Tableau croisé dynamique utilisant la disposition plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>Disposition Tabulaire

![Tableau croisé dynamique utilisant la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    await context.sync();
    
    // cycling through layout styles
    if (pivotTable.layout.layoutType === "Compact") {
        pivotTable.layout.layoutType = "Outline";
    } else if (pivotTable.layout.layoutType === "Outline") {
        pivotTable.layout.layoutType = "Tabular";
    } else {
        pivotTable.layout.layoutType = "Compact";
    }
    
    await context.sync();
});
```

## <a name="change-hierarchy-names"></a>Modifier les noms des hiérarchies

Les champs des hiérarchies peuvent être modifiés. Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.

```typescript
await Excel.run(async (context) => {
    const dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();
    
    // changing the displayed names of these entries
    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
    await context.sync();
});
```

## <a name="delete-a-pivottable"></a>Supprimer un tableau croisé dynamique

Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

> [!NOTE]
> Votre avis sur la conception de nos préversions est le bienvenu. Si vous avez des commentaires, des suggestions ou des problèmes avec la nouvelle API Tableau croisé dynamique, veuillez laisser vos commentaires sur [UserVoice](https://officespdev.uservoice.com/forums/224641-feature-requests-and-feedback?category_id=163563) ou dans le [répertoire OpenSpec GitHub](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec).
