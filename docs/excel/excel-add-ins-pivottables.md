---
title: Utilisation des tableaux croisés dynamiques avec l’API JavaScript pour Excel
description: Utilisez l’API JavaScript pour Excel pour créer des tableaux croisés dynamiques et interagir avec leurs composants.
ms.date: 01/22/2020
localization_priority: Normal
ms.openlocfilehash: 39dca0ca3f964133af64066641d7bb07222c7834
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554028"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Utilisation des tableaux croisés dynamiques avec l’API JavaScript pour Excel

Les tableaux croisés dynamiques rationalisent les grands ensembles de données. Ils permettent la manipulation rapide des données groupées. L’API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d’interagir avec leurs composants. Cet article explique comment les tableaux croisés dynamiques sont représentés par l’API JavaScript Office et fournit des exemples de code pour les scénarios clés.

Si vous n’êtes pas familiarisé avec la fonctionnalité de tableaux croisés dynamiques, envisagez de les explorer comme un utilisateur final.
Reportez-vous à la rubrique [créer un tableau croisé dynamique pour analyser les données de feuille de calcul](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) pour obtenir une introduction à ces outils.

> [!IMPORTANT]
> Les tableaux croisés dynamiques créés avec OLAP ne sont actuellement pas pris en charge. Il n’existe pas non plus de prise en charge de PowerPivot.

## <a name="object-model"></a>Modèle d’objet

Le [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) est l’objet central pour les tableaux croisés dynamiques de l’API JavaScript pour Office.

- `Workbook.pivotTables`et `Worksheet.pivotTables` sont [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) qui contiennent respectivement les [tableaux croisés dynamiques](/javascript/api/excel/excel.pivottable) dans le classeur et la feuille de calcul.
- Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) qui comporte plusieurs [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).
- Un [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contient un [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) qui comporte exactement un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield). Si la conception s’étend pour inclure des tableaux croisés dynamiques OLAP, cela peut changer.
- Un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield) contient un [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) avec plusieurs [PivotItems](/javascript/api/excel/excel.pivotitem).
- Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotLayout](/javascript/api/excel/excel.pivotlayout) qui définit où les [champs PivotFields](/javascript/api/excel/excel.pivotfield) et [PivotItems](/javascript/api/excel/excel.pivotitem) sont affichés dans la feuille de calcul.

Examinons comment ces relations s’appliquent à certains exemples de données. Les données suivantes décrivent les ventes de fruit de différentes batteries de serveurs. Il s’agit de l’exemple de cet article.

![Collection de ventes de fruit de différents types de batteries de serveurs différentes.](../images/excel-pivots-raw-data.png)

Les données de ventes de la batterie de fruits seront utilisées pour créer un tableau croisé dynamique. Chaque colonne, telle que **types**, est `PivotHierarchy`. La hiérarchie de **types** contient le champ **types** . Le champ **types** contient les éléments **Apple**, **Kiwi**, **citron**, **citron**et **orange**.

### <a name="hierarchies"></a>Hierarchies

Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie : [ligne](/javascript/api/excel/excel.rowcolumnpivothierarchy), [colonne](/javascript/api/excel/excel.rowcolumnpivothierarchy), [données](/javascript/api/excel/excel.datapivothierarchy)et [filtre](/javascript/api/excel/excel.filterpivothierarchy).

Les données de la batterie de serveurs affichées précédemment ont cinq hiérarchies : **batteries**de serveurs, **type**, **classification**, **caisses vendues à la batterie de serveurs**et **caisses vendues en gros**. Chaque hiérarchie peut uniquement exister dans l’une des quatre catégories. Si le **type** est ajouté aux hiérarchies de colonne, il ne peut pas également se trouver dans les hiérarchies de ligne, de données ou de filtre. Si **type** est par la suite ajouté aux hiérarchies de lignes, il est supprimé des hiérarchies de colonne. Ce comportement est le même, que l’attribution de hiérarchie soit réalisée via l’interface utilisateur Excel ou les API JavaScript pour Excel.

Les hiérarchies de ligne et de colonne définissent le mode de regroupement des données. Par exemple, une hiérarchie de lignes de **batteries de serveurs** regroupe tous les jeux de données de la même batterie de serveurs. Le choix entre la hiérarchie de ligne et de colonne définit l’orientation du tableau croisé dynamique.

Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de ligne et de colonne. Un tableau croisé dynamique avec une hiérarchie de lignes de **batteries de serveurs** et une hiérarchie de données de **grossistes vendus en gros** indique le total de tous les fruits de chaque batterie de serveurs.

Les hiérarchies de filtre incluent ou excluent les données du tableau croisé dynamique en fonction des valeurs contenues dans ce type filtré. Une hiérarchie de filtrage de **classification** avec le type **Organic** Selected affiche uniquement les données pour les fruits organiques.

Voici les données de la batterie de serveurs à nouveau, ainsi qu’un tableau croisé dynamique. Le tableau croisé dynamique utilise la **batterie de serveurs** et le **type** comme hiérarchies de lignes, les **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut Sum) et une **classification** en tant que hiérarchie de filtres (avec l’option **Organic** sélectionnée).

![Sélection de données sur les ventes de fruit en regard d’un tableau croisé dynamique avec des hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

Ce tableau croisé dynamique peut être généré via l’API JavaScript ou via l’interface utilisateur d’Excel. Ces deux options permettent une manipulation supplémentaire via les compléments.

## <a name="create-a-pivottable"></a>Créer un tableau croisé dynamique

Les tableaux croisés dynamiques nécessitent un nom, une source et une destination. La source peut être une adresse de plage ou un nom de table ( `Range`transmis `string`en tant `Table` que type, ou type). La destination est une adresse de plage (sous la forme `Range` a `string`ou).
Les exemples suivants illustrent différentes techniques de création de tableau croisé dynamique.

### <a name="create-a-pivottable-with-range-addresses"></a>Créer un tableau croisé dynamique avec des adresses de plage

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a>Création d’un tableau croisé dynamique avec des objets Range

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data comes from the worksheet "DataWorksheet" across the range A1:E21.
    var rangeToAnalyze = context.workbook.worksheets.getItem("DataWorksheet").getRange("A1:E21");
    var rangeToPlacePivot = context.workbook.worksheets.getItem("PivotWorksheet").getRange("A2");
    context.workbook.worksheets.getItem("PivotWorksheet").pivotTables.add(
      "Farm Sales", rangeToAnalyze, rangeToPlacePivot);

    return context.sync();
});
```

### <a name="create-a-pivottable-at-the-workbook-level"></a>Création d’un tableau croisé dynamique au niveau du classeur

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a>Utiliser un tableau croisé dynamique existant

Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection de tableau croisé dynamique du classeur ou de feuilles de calcul individuelles. Le code suivant obtient un tableau croisé dynamique nommé **mon tableau croisé dynamique** à partir du classeur.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Ajouter des lignes et des colonnes à un tableau croisé dynamique

Lignes et colonnes tableau croisé dynamique des données autour de ces valeurs.

L’ajout de la colonne **batterie de serveurs** pivote toutes les ventes autour de chaque batterie de serveurs. L’ajout des lignes de type et de **classification** répartit davantage les données en fonction des fruits vendus et s’il s’agit d’un **type** Organic ou non.

![Un tableau croisé dynamique avec une colonne de batterie de serveurs et des lignes de type et de classification.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

Vous pouvez également utiliser un tableau croisé dynamique avec uniquement des lignes ou des colonnes.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a>Ajouter des hiérarchies de données au tableau croisé dynamique

Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes. L’ajout des hiérarchies de données des **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** fournit des sommes de ces chiffres pour chaque ligne et colonne.

Dans l’exemple, la **batterie de serveurs** et le **type** sont des lignes, avec le caisse ventes comme données.

![Tableau croisé dynamique illustrant les ventes totales de fruits différents en fonction de la batterie de serveurs à partir de laquelle ils provenaient.](../images/excel-pivots-data-hierarchy.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // "Farm" and "Type" are the hierarchies on which the aggregation is based.
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));

    // "Crates Sold at Farm" and "Crates Sold Wholesale" are the hierarchies
    // that will have their data aggregated (summed in this case).
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    return context.sync();
});
```

## <a name="slicers"></a>Slicers

Les [segments](/javascript/api/excel/excel.slicer) permettent aux données d’être filtrées à partir d’un tableau croisé dynamique ou d’un tableau Excel. Un segment utilise des valeurs d’une colonne ou d’un champ PivotField spécifié pour filtrer les lignes correspondantes. Ces valeurs sont stockées en [](/javascript/api/excel/excel.sliceritem) tant qu’objets SlicerItem `Slicer`dans le. Votre complément peut ajuster ces filtres, comme les utilisateurs peuvent les[utiliser (par le biais de l’interface utilisateur Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)). Le segment se trouve au-dessus de la feuille de calcul de la couche de dessin, comme illustré dans la capture d’écran suivante.

![Données de filtrage de segment sur un tableau croisé dynamique.](../images/excel-slicer.png)

> [!NOTE]
> Les techniques décrites dans cette section se concentrent sur l’utilisation des segments connectés aux tableaux croisés dynamiques. Les mêmes techniques s’appliquent également à l’utilisation de segments connectés à des tables.

### <a name="create-a-slicer"></a>Créer un segment

Vous pouvez créer un segment dans un classeur ou une feuille de calcul `Workbook.slicers.add` à l' `Worksheet.slicers.add` aide de la méthode ou de la méthode. Cette opération ajoute un Slicer au [SlicerCollection](/javascript/api/excel/excel.slicercollection) de l’objet spécifié `Workbook` ou `Worksheet` . La `SlicerCollection.add` méthode comporte trois paramètres :

- `slicerSource`: La source de données sur laquelle le nouveau segment est basé. Il peut s’agir `PivotTable`d' `Table`un,, ou d’une chaîne représentant le nom `PivotTable` ou `Table`l’ID d’un ou d’un.
- `sourceField`: Champ dans la source de données à utiliser pour filtrer. Il peut s’agir `PivotField`d' `TableColumn`un,, ou d’une chaîne représentant le nom `PivotField` ou `TableColumn`l’ID d’un ou d’un.
- `slicerDestination`: La feuille de calcul dans laquelle le nouveau segment sera créé. Il peut s’agir `Worksheet` d’un objet ou du nom ou de `Worksheet`l’ID d’un. Ce paramètre n’est pas nécessaire `SlicerCollection` lorsque le est `Worksheet.slicers`accessible via. Dans ce cas, la feuille de calcul de la collection est utilisée comme destination.

L’exemple de code suivant ajoute un nouveau segment à la feuille de calcul de **tableau croisé dynamique** . La source du Slicer est le tableau croisé dynamique de la **batterie de serveurs** et les filtres utilisant les données de **type** . Le segment est également nommé **segment de fruit** pour référence ultérieure.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Pivot");
    var slicer = sheet.slicers.add(
        "Farm Sales" /* The slicer data source. For PivotTables, this can be the PivotTable object reference or name. */,
        "Type" /* The field in the data to filter by. For PivotTables, this can be a PivotField object reference or ID. */
    );
    slicer.name = "Fruit Slicer";
    return context.sync();
});
```

### <a name="filter-items-with-a-slicer"></a>Filtrer des éléments avec un segment

Le segment filtre le tableau croisé dynamique avec les éléments `sourceField`de la. La `Slicer.selectItems` méthode définit les éléments qui restent dans le Slicer. Ces éléments sont transmis à la méthode en tant `string[]`que, représentant les clés des éléments. Toutes les lignes contenant ces éléments restent dans l’agrégation du tableau croisé dynamique. Appels suivants permettant `selectItems` de définir la liste aux clés spécifiées dans ces appels.

> [!NOTE]
> Si `Slicer.selectItems` reçoit un élément qui ne se trouve pas dans la source de données `InvalidArgument` , une erreur est générée. Le contenu peut être vérifié via la `Slicer.slicerItems` propriété, qui est une [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).

L’exemple de code suivant montre trois éléments sélectionnés pour le Slicer : **citron**, **citron**et **orange**.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

Pour supprimer tous les filtres du segment, utilisez la `Slicer.clearFilters` méthode, comme illustré dans l’exemple suivant.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

### <a name="style-and-format-a-slicer"></a>Style et formatage d’un segment

Vous pouvez ajuster les paramètres d’affichage d’un segment par le biais `Slicer` de propriétés. L’exemple de code suivant définit le style sur **SlicerStyleLight6**, définit le texte en haut du Slicer sur **types de fruit**, place le segment à la position **(395, 15)** sur la couche de dessin et définit la taille du Slicer sur **135x150** pixels.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.caption = "Fruit Types";
    slicer.left = 395;
    slicer.top = 15;
    slicer.height = 135;
    slicer.width = 150;
    slicer.style = "SlicerStyleLight6";
    return context.sync();
});
```

### <a name="delete-a-slicer"></a>Supprimer un segment

Pour supprimer un segment, appelez la `Slicer.delete` méthode. L’exemple de code suivant supprime le premier segment de la feuille de calcul active.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>Modifier la fonction d’agrégation

Les hiérarchies de données ont leurs valeurs agrégées. Pour les jeux de données de nombres, il s’agit d’une somme par défaut. La `summarizeBy` propriété définit ce comportement en fonction d’un type [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .

Les types de fonction d’agrégation actuellement `Sum`pris `Count`en `Average`charge `Max`sont `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`,, `Automatic` ,, et (valeur par défaut).

Les exemples de code suivants modifient l’agrégation pour qu’elle soit la moyenne des données.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.dataHierarchies.load("no-properties-needed");
    return context.sync().then(function() {

        // Change the aggregation from the default sum to an average of all the values in the hierarchy.
        pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.average;
        pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.average;
        return context.sync();
    });
});
```

## <a name="change-calculations-with-a-showasrule"></a>Modifier les calculs avec une ShowAsRule

Les tableaux croisés dynamiques agrègent par défaut les données de leurs hiérarchies de ligne et de colonne indépendamment. Un [ShowAsRule](/javascript/api/excel/excel.showasrule) modifie la hiérarchie des données en valeurs de sortie en fonction d’autres éléments du tableau croisé dynamique.

L' `ShowAsRule` objet possède trois propriétés :

- `calculation`: Type de calcul relatif à appliquer à la hiérarchie de données (la valeur par `none`défaut est).
- `baseField`: [Champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield) dans la hiérarchie contenant les données de base avant l’application du calcul. Étant donné que les tableaux croisés dynamiques Excel ont un mappage un-à-un de la hiérarchie sur champ, vous utiliserez le même nom pour accéder à la hiérarchie et au champ.
- `baseItem`: La valeur de [PivotItem](/javascript/api/excel/excel.pivotitem) individuelle comparée aux valeurs des champs de base basés sur le type de calcul. Tous les calculs ne nécessitent pas ce champ.

L’exemple suivant montre comment définir le calcul sur la **somme des caisses vendues dans** la hiérarchie des données de la batterie de serveurs pour qu’elle soit un pourcentage du total de la colonne.
Nous souhaitons toujours que la granularité s’étende au niveau du type de fruit, c’est pourquoi nous allons utiliser la hiérarchie des lignes de **type** et le champ sous-jacent.
L’exemple dispose également d’une **batterie de serveurs** comme première hiérarchie de lignes, de sorte que le nombre total d’entrées de batterie de serveurs affiche également le pourcentage de production de chaque batterie.

![Tableau croisé dynamique indiquant le pourcentage de ventes de fruits par rapport au total général pour les batteries individuelles et les types de fruits individuels au sein de chaque batterie de serveurs.](../images/excel-pivots-showas-percentage.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {

        // Show the crates of each fruit type sold at the farm as a percentage of the column's total.
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.percentOfColumnTotal;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Type").fields.getItem("Type");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Percentage of Total Farm Sales";
    });
});
```

L’exemple précédent définit le calcul sur la colonne, par rapport au champ d’une hiérarchie de lignes individuelle. Lorsque le calcul est lié à un élément individuel, utilisez `baseItem` la propriété.

L’exemple suivant montre le `differenceFrom` calcul. Il affiche la différence entre les entrées de hiérarchie de données ventes de la batterie de serveurs par rapport à celles d' **une**batterie de serveurs.
La `baseField` **batterie de serveurs**is, de sorte que nous voyons les différences entre les autres batteries de serveurs, ainsi que les répartitions pour chaque type de fruit similaire (**type** est également une hiérarchie de lignes dans cet exemple).

![Tableau croisé dynamique montrant les différences entre les ventes de fruit et les autres. Cela montre à la fois la différence entre les ventes de fruits totales des batteries de serveurs et les ventes de types de fruits. Si « une batterie de serveurs » n’a pas vendu un type particulier de fruit, « #N/A » s’affiche.](../images/excel-pivots-showas-differencefrom.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var farmDataHierarchy = pivotTable.dataHierarchies.getItem("Sum of Crates Sold at Farm");

    farmDataHierarchy.load("showAs");
    return context.sync().then(function () {
        // Show the difference between crate sales of the "A Farms" and the other farms.
        // This difference is both aggregated and shown for individual fruit types (where applicable).
        var farmShowAs = farmDataHierarchy.showAs;
        farmShowAs.calculation = Excel.ShowAsCalculation.differenceFrom;
        farmShowAs.baseField = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm");
        farmShowAs.baseItem = pivotTable.rowHierarchies.getItem("Farm").fields.getItem("Farm").items.getItem("A Farms");
        farmDataHierarchy.showAs = farmShowAs;
        farmDataHierarchy.name = "Difference from A Farms";
    });
});
```

## <a name="pivottable-layouts"></a>Dispositions du tableau croisé dynamique

Un [PivotLayout](/javascript/api/excel/excel.pivotlayout) définit l’emplacement des hiérarchies et de leurs données. Vous accédez à la disposition pour déterminer les plages dans lesquelles les données sont stockées.

Le diagramme suivant montre les appels de fonction de disposition qui correspondent aux plages du tableau croisé dynamique.

![Diagramme montrant les sections d’un tableau croisé dynamique renvoyées par les fonctions Get Range de la disposition.](../images/excel-pivots-layout-breakdown.png)

Le code suivant montre comment obtenir la dernière ligne des données du tableau croisé dynamique en parcourant la disposition. Ces valeurs sont ensuite additionnées pour un total général.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("B27:C27");
        masterTotalRange.formulas = [["All Crates", "=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

Les tableaux croisés dynamiques ont trois styles de disposition : compact, plan et tabulaire. Nous avons vu le style compact dans les exemples précédents.

Les exemples suivants utilisent respectivement les styles de plan et de tableau. L’exemple de code montre comment effectuer un basculement entre les différentes dispositions.

### <a name="outline-layout"></a>Mise en page du plan

![Tableau croisé dynamique à l’aide de la mise en forme du plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a>Disposition tabulaire

![Un tableau croisé dynamique à l’aide de la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a>Modifier les noms de hiérarchie

Les champs de hiérarchie sont modifiables. Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.

```js
Excel.run(function (context) {
    var dataHierarchies = context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.getItem("Farm Sales").dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    return context.sync().then(function () {
        // changing the displayed names of these entries
        dataHierarchies.items[0].name = "Farm Sales";
        dataHierarchies.items[1].name = "Wholesale";
    });
});
```

## <a name="delete-a-pivottable"></a>Supprimer un tableau croisé dynamique

Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Référence de l’API JavaScript pour Excel](/javascript/api/excel)
