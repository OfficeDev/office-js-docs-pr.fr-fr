---
title: Utilisation des tableaux croisés dynamiques avec l'API JavaScript pour Excel
description: Utilisez l'API JavaScript pour Excel pour créer des tableaux croisés dynamiques et interagir avec leurs composants.
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: b53d734e676417a6438f1008bac720a38a244d1f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870323"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Utilisation des tableaux croisés dynamiques avec l'API JavaScript pour Excel

Les tableaux croisés dynamiques rationalisent les grands ensembles de données. Ils permettent la manipulation rapide des données groupées. L'API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d'interagir avec leurs composants.

Si vous n'êtes pas familiarisé avec la fonctionnalité de tableaux croisés dynamiques, envisagez de les explorer comme un utilisateur final. RePortez-vous à la rubrique [créer un tableau croisé dynamique pour analyser les données de feuille de calcul](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) pour obtenir une introduction à ces outils. 

Cet article fournit des exemples de code pour les scénarios courants. Pour mieux comprendre l'API PivotTable, consultez la rubrique [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).

> [!IMPORTANT]
> Les tableaux croisés dynamiques créés avec OLAP ne sont actuellement pas pris en charge. Il n'existe pas non plus de prise en charge de PowerPivot.

## <a name="hierarchies"></a>Hierarchies

Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie: ligne, colonne, données et filtre. Les données suivantes décrivant les ventes de fruit de différentes batteries de serveurs seront utilisées tout au long de cet article.

![Collection de ventes de fruit de différents types de batteries de serveurs différentes.](../images/excel-pivots-raw-data.png)

Ces données ont cinq hiérarchies **: batteries de serveurs**, **type**, **classification**, **caisses vendues à la batterie de serveurs**et **caisses vendues en gros**. Chaque hiérarchie peut uniquement exister dans l'une des quatre catégories. Si le **type** est ajouté aux hiérarchies de colonnes, puis ajouté aux hiérarchies de lignes, il n'est conservé que dans ce dernier.

Les hiérarchies de ligne et de colonne définissent le mode de regroupement des données. Par exemple, une hiérarchie de lignes de **batteries de serveurs** regroupe tous les jeux de données de la même batterie de serveurs. Le choix entre la hiérarchie de ligne et de colonne définit l'orientation du tableau croisé dynamique.

Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de ligne et de colonne. Un tableau croisé dynamique avec une hiérarchie de lignes de **batteries de serveurs** et une hiérarchie de données de grossistes **vendus en gros** indique le total de tous les fruits de chaque batterie de serveurs.

Les hiérarchies de filtre incluent ou excluent les données du tableau croisé dynamique en fonction des valeurs contenues dans ce type filtré. Une hiérarchie de filtrage de **classification** avec le type **Organic** Selected affiche uniquement les données pour les fruits organiques.

Voici les données de la batterie de serveurs à nouveau, ainsi qu'un tableau croisé dynamique. Le tableau croisé dynamique utilise la **batterie de serveurs** et le **type** comme hiérarchies de lignes, les **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** en tant que hiérarchies de données (avec la fonction d'agrégation par défaut Sum) et une **classification** en tant que filtre hiérarchie (avec l'option **Organic** sélectionnée). 

![Sélection de données sur les ventes de fruit en regard d'un tableau croisé dynamique avec des hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

Ce tableau croisé dynamique peut être généré via l'API JavaScript ou via l'interface utilisateur d'Excel. Ces deux options permettent une manipulation supplémentaire via les compléments.

## <a name="create-a-pivottable"></a>Créer un tableau croisé dynamique

Les tableaux croisés dynamiques nécessitent un nom, une source et une destination. La source peut être une adresse de plage ou un nom de table ( `Range`transmis `string`en tant `Table` que type, ou type). La destination est une adresse de plage (sous la forme `Range` a `string`ou). Les exemples suivants illustrent différentes techniques de création de tableau croisé dynamique.

### <a name="create-a-pivottable-with-range-addresses"></a>Créer un tableau croisé dynamique avec des adresses de plage

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Création d'un tableau croisé dynamique avec des objets Range

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21
    const rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    const rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
        "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    await context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>Création d'un tableau croisé dynamique au niveau du classeur

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Utiliser un tableau croisé dynamique existant

Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection de tableau croisé dynamique du classeur ou de feuilles de calcul individuelles. 

Le code suivant obtient le premier tableau croisé dynamique du classeur. Il donne ensuite un nom à la table pour une référence facile.

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Ajouter des lignes et des colonnes à un tableau croisé dynamique

Lignes et colonnes tableau croisé dynamique des données autour de ces valeurs.

L'ajout de la colonne **batterie de serveurs** pivote toutes les ventes autour de chaque batterie de serveurs. L'ajout des lignes de type et de **classification** répartit davantage les données en fonction des fruits vendus et s'il s'agit d'un **type** Organic ou non.

![Un tableau croisé dynamique avec une colonne de batterie de serveurs et des lignes de type et de classification.](../images/excel-pivots-table-rows-and-columns.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    await context.sync();
});
```

Vous pouvez également utiliser un tableau croisé dynamique avec uniquement des lignes ou des colonnes.

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

Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes. L'ajout des hiérarchies de données des **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** fournit des sommes de ces chiffres pour chaque ligne et colonne. 

Dans l'exemple, la **batterie de serveurs** et le **type** sont des lignes, avec le caisse ventes comme données. 

![Tableau croisé dynamique illustrant les ventes totales de fruits différents en fonction de la batterie de serveurs à partir de laquelle ils provenaient.](../images/excel-pivots-data-hierarchy.png)

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case)
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    await context.sync();
});
```

## <a name="change-aggregation-function"></a>Modifier la fonction d'agrégation

Les hiérarchies de données ont leurs valeurs agrégées. Pour les jeux de données de nombres, il s'agit d'une somme par défaut. La `summarizeBy` propriété définit ce comportement en fonction d'un type [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .

Les types de fonction d'agrégation actuellement `Sum`pris `Count`en `Average`charge `Max`sont `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`,, `Automatic` ,, et (valeur par défaut).

Les exemples de code suivants modifient l'agrégation pour qu'elle soit la moyenne des données.

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

## <a name="change-calculations-with-a-showasrule"></a>Modifier les calculs avec une ShowAsRule

Les tableaux croisés dynamiques agrègent par défaut les données de leurs hiérarchies de ligne et de colonne indépendamment. Un [ShowAsRule](/javascript/api/excel/excel.showasrule) modifie la hiérarchie des données en valeurs de sortie en fonction d'autres éléments du tableau croisé dynamique.

L' `ShowAsRule` objet possède trois propriétés:

-   `calculation`: Type de calcul relatif à appliquer à la hiérarchie de données (la valeur par `none`défaut est).
-   `baseField`: Champ au sein de la hiérarchie contenant les données de base avant l'application du calcul. Le [champ PivotField](/javascript/api/excel/excel.pivotfield) a généralement le même nom que sa hiérarchie parente.
-   `baseItem`: La valeur de [PivotItem](/javascript/api/excel/excel.pivotitem) individuelle comparée aux valeurs des champs de base basés sur le type de calcul. Tous les calculs ne nécessitent pas ce champ.

L'exemple suivant montre comment définir le calcul sur la **somme des caisses vendues dans** la hiérarchie des données de la batterie de serveurs pour qu'elle soit un pourcentage du total de la colonne. Nous souhaitons toujours que la granularité s'étende au niveau du type de fruit, c'est pourquoi nous allons utiliser la hiérarchie des lignes de **type** et le champ sous-jacent. L'exemple dispose également d'une **batterie de serveurs** comme première hiérarchie de lignes, de sorte que le nombre total d'entrées de batterie de serveurs affiche également le pourcentage de production de chaque batterie.

![Tableau croisé dynamique indiquant le pourcentage de ventes de fruits par rapport au total général pour les batteries individuelles et les types de fruits individuels au sein de chaque batterie de serveurs.](../images/excel-pivots-showas-percentage.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the crates of each fruit type sold at the farm as a percentage of the column's total
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
    farmDataHierarchy.showAs = farmShowAs; 
    farmDataHierarchy.name = "Percentage of Total Farm Sales";

    await context.sync();
});
```

L'exemple précédent définit le calcul de la colonne, par rapport à une hiérarchie de lignes individuelle. Lorsque le calcul est lié à un élément individuel, utilisez `baseItem` la propriété.

L'exemple suivant montre le `differenceFrom` calcul. Il affiche la différence entre les entrées de hiérarchie des données sur les ventes de la batterie de serveurs par rapport à celles des «batteries de serveurs».
La `baseField` **batterie de serveurs**is, de sorte que nous voyons les différences entre les autres batteries de serveurs, ainsi que les répartitions pour chaque type de fruit similaire (**type** est également une hiérarchie de lignes dans cet exemple).

![Tableau croisé dynamique montrant les différences entre les ventes de fruit et les autres. Cela montre à la fois la différence entre les ventes de fruits totales des batteries de serveurs et les ventes de types de fruits. Si «une batterie de serveurs» n'a pas vendu un type particulier de fruit, «#N/A» s'affiche.](../images/excel-pivots-showas-differencefrom.png)

``` TypeScript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    const farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    await context.sync();

    // show the difference between crate sales of the "A Farms" and the other farms
    // this difference is both aggregated and shown for individual fruit types (where applicable)
    let farmShowAs = farmDataHierarchy.showAs;
    farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
    farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
    farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
    farmDataHierarchy.showAs = farmShowAs;
    farmDataHierarchy.name = "Difference from A Farms";
    await context.sync();
});
```

## <a name="pivottable-layouts"></a>Dispositions du tableau croisé dynamique

Un [PivotLayout](/javascript/api/excel/excel.pivotlayout) définit l'emplacement des hiérarchies et de leurs données. Vous accédez à la disposition pour déterminer les plages dans lesquelles les données sont stockées.

Le diagramme suivant montre les appels de fonction de disposition qui correspondent aux plages du tableau croisé dynamique.

![Diagramme montrant les sections d'un tableau croisé dynamique renvoyées par les fonctions Get Range de la disposition.](../images/excel-pivots-layout-breakdown.png)

Le code suivant montre comment obtenir la dernière ligne des données du tableau croisé dynamique en parcourant la disposition. Ces valeurs sont ensuite additionnées pour un total général.

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

Les tableaux croisés dynamiques ont trois styles de disposition: compact, plan et tabulaire. Nous avons vu le style compact dans les exemples précédents. 

Les exemples suivants utilisent respectivement les styles de plan et de tableau. L'exemple de code montre comment effectuer un basculement entre les différentes dispositions.

### <a name="outline-layout"></a>Mise en page du plan

![Tableau croisé dynamique à l'aide de la mise en forme du plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>Disposition tabulaire

![Un tableau croisé dynamique à l'aide de la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>Modifier les noms de hiérarchie

Les champs de hiérarchie sont modifiables. Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.

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

Les tableaux croisés dynamiques sont supprimés à l'aide de leur nom.

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Référence de l'API JavaScript pour Excel](/javascript/api/excel)
