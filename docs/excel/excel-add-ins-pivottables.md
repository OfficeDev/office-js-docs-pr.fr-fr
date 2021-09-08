---
title: Utiliser des tableaux croisés dynamiques à l’aide Excel API JavaScript
description: Utilisez l Excel API JavaScript pour créer des tableaux croisés dynamiques et interagir avec leurs composants.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: d9ccaf72be4fa23b73f1f91d38d240ea02569eca
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936262"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a>Utiliser des tableaux croisés dynamiques à l’aide Excel API JavaScript

Les tableaux croisés dynamiques simplifient les jeux de données plus volumineux. Elles permettent la manipulation rapide des données groupées. L Excel API JavaScript permet à votre application de créer des tableaux croisés dynamiques et d’interagir avec leurs composants. Cet article décrit comment les tableaux croisés dynamiques sont représentés par Office API JavaScript et fournit des exemples de code pour les scénarios clés.

Si vous ne connaissez pas la fonctionnalité des tableaux croisés dynamiques, envisagez de les explorer en tant qu’utilisateur final.
Voir [Créer un tableau croisé dynamique pour analyser les](https://support.microsoft.com/office/ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EBBD=PivotTables) données de feuille de calcul afin d’obtenir une bonne base sur ces outils.

> [!IMPORTANT]
> Les tableaux croisés dynamiques créés avec OLAP ne sont actuellement pas pris en charge. Il n’existe pas non plus de prise en charge de Power Pivot.

## <a name="object-model"></a>Modèle d’objet

Le [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) est l’objet central des tableaux croisés dynamiques dans l Office API JavaScript.

- `Workbook.pivotTables` et `Worksheet.pivotTables` sont [des PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) qui contiennent respectivement les tableaux [croisés dynamiques](/javascript/api/excel/excel.pivottable) dans le workbook et la feuille de calcul.
- Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) qui possède plusieurs [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).
- Ces [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) peuvent être ajoutées à des collections de hiérarchies spécifiques pour définir la façon dont le tableau croisé dynamique analyse les données (comme expliqué dans la [section suivante).](#hierarchies)
- Une [PivotHierarchy contient](/javascript/api/excel/excel.pivothierarchy) un [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) qui possède exactement un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield). Si la conception est étendue pour inclure des tableaux croisés dynamiques OLAP, cela peut changer.
- Un [champ de](/javascript/api/excel/excel.pivotfield) tableau croisé dynamique peut avoir un ou plusieurs filtres de tableau croisé dynamique [appliqués,](/javascript/api/excel/excel.pivotfilters) tant que la [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) du champ est affectée à une catégorie de hiérarchie.
- Un [champ de](/javascript/api/excel/excel.pivotfield) tableau croisé dynamique contient un [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) qui a plusieurs [pivotItems](/javascript/api/excel/excel.pivotitem).
- Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [pivotLayout](/javascript/api/excel/excel.pivotlayout) qui définit l’endroit où les [pivotFields](/javascript/api/excel/excel.pivotfield) et [pivotItems](/javascript/api/excel/excel.pivotitem) sont affichés dans la feuille de calcul. La disposition contrôle également certains paramètres d’affichage du tableau croisé dynamique.

Examinons comment ces relations s’appliquent à certains exemples de données. Les données suivantes décrivent les ventes de fruit de différentes batteries de serveurs. Ce sera l’exemple tout au long de cet article.

![Collection de ventes de fruit de différents types de batteries de serveurs.](../images/excel-pivots-raw-data.png)

Les données de ventes de cette batterie de serveurs de fruit seront utilisées pour la production d’un tableau croisé dynamique. Chaque colonne, telle que **Types,** est une `PivotHierarchy` . La **hiérarchie Types** contient le champ **Types.** Le **champ Types** contient les éléments **Apple**, **Domaine,** **Citron,** **Vert** et **Orange**.

### <a name="hierarchies"></a>Hierarchies

Les tableaux croisés dynamiques sont organisés en quatre catégories hiérarchiques : [ligne,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [colonne,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [données](/javascript/api/excel/excel.datapivothierarchy)et [filtre.](/javascript/api/excel/excel.filterpivothierarchy)

Les données de batterie de serveurs indiquées précédemment disposent de cinq hiérarchies : Batteries **de** serveurs, **Type**, **Classification**, **Caisses vendues** à la batterie de serveurs et **Caisses vendues en commun**. Chaque hiérarchie ne peut exister que dans l’une des quatre catégories. Si **type** est ajouté aux hiérarchies de colonnes, il ne peut pas non plus se trouver dans les hiérarchies de lignes, de données ou de filtres. Si **Type** est ensuite ajouté aux hiérarchies de lignes, il est supprimé des hiérarchies de colonnes. Ce comportement est le même si l’affectation de hiérarchie est effectuée via l’interface Excel’interface utilisateur ou Excel api JavaScript.

Les hiérarchies de lignes et de colonnes définissent le regroupement des données. Par exemple, une hiérarchie de lignes **de** batteries de serveurs groupe tous les ensembles de données de la même batterie de serveurs. Le choix entre la hiérarchie de lignes et de colonnes définit l’orientation du tableau croisé dynamique.

Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de lignes et de colonnes. Un tableau croisé dynamique avec  une hiérarchie de lignes de batteries de serveurs et une hiérarchie de données de l’ordre des **caisses vendues indique** le total total (par défaut) de tous les différents produits pour chaque batterie de serveurs.

Les hiérarchies de filtres incluent ou excluent des données du tableau croisé dynamique en fonction des valeurs de ce type filtré. Une hiérarchie de filtres de **classification** avec le type **organique** sélectionné affiche uniquement les données pour les fruit organiques.

Voici à nouveau les données de la batterie de serveurs, ainsi qu’un tableau croisé dynamique. Le tableau croisé dynamique utilise  Farm **and** **Type** comme hiérarchies de lignes, La vente des **caisses** sur la batterie de serveurs et la vente **de caisses** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut de somme) et **classification** en tant que hiérarchie de filtre (avec l’alimentation organique sélectionnée).

![Sélection de données de ventes de fruit à côté d’un tableau croisé dynamique avec des hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

Ce tableau croisé dynamique peut être généré via l’API JavaScript ou par le biais Excel’interface utilisateur. Les deux options permettent d’autres manipulations par le biais de leurs add-ins.

## <a name="create-a-pivottable"></a>Créer un tableau croisé dynamique

Les tableaux croisés dynamiques ont besoin d’un nom, d’une source et d’une destination. La source peut être une adresse de plage ou un nom de table (transmis en tant `Range` `string` que , ou `Table` type). La destination est une adresse de plage (donnée en tant que a `Range` ou `string` ).
Les exemples suivants montrent différentes techniques de création de tableau croisé dynamique.

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

### <a name="create-a-pivottable-with-range-objects"></a>Créer un tableau croisé dynamique avec des objets Range

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

### <a name="create-a-pivottable-at-the-workbook-level"></a>Créer un tableau croisé dynamique au niveau du workbook

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

Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection de tableaux croisés dynamiques du manuel ou des feuilles de calcul individuelles. Le code suivant obtient un tableau croisé dynamique nommé **My Pivot** à partir du workbook.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a>Ajouter des lignes et des colonnes à un tableau croisé dynamique

Les lignes et les colonnes pivotent les données autour des valeurs de ces champs.

L’ajout **de la** colonne Batterie de serveurs pivote toutes les ventes autour de chaque batterie de serveurs. L’ajout **des lignes Type** et **Classification** décompose davantage les données en fonction des fruit vendus et selon qu’il s’agit d’un produit organique ou non.

![Tableau croisé dynamique avec une colonne de batterie de serveurs et des lignes Type et Classification.](../images/excel-pivots-table-rows-and-columns.png)

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));

    return context.sync();
});
```

Vous pouvez également avoir un tableau croisé dynamique avec uniquement des lignes ou des colonnes.

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

Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes. L’ajout des hiérarchies de données **des caisses vendues** au niveau de la batterie de serveurs et des **caisses vendues permet** d’obtenir les sommes de ces chiffres pour chaque ligne et colonne.

Dans l’exemple, **la batterie de** serveurs et le **type** sont des lignes, avec les ventes de caisse en tant que données.

![Tableau croisé dynamique montrant les ventes totales de différents fruit en fonction de la batterie de serveurs d’où ils sont issus.](../images/excel-pivots-data-hierarchy.png)

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a>Dispositions de tableau croisé dynamique et obtention de données pivotées

Un [pivotLayout](/javascript/api/excel/excel.pivotlayout) définit le placement des hiérarchies et leurs données. Vous accédez à la disposition pour déterminer les plages où les données sont stockées.

Le diagramme suivant montre les appels de fonction de disposition qui correspondent aux plages du tableau croisé dynamique.

![Diagramme montrant quelles sections d’un tableau croisé dynamique sont renvoyées par les fonctions get range de la disposition.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a>Obtenir des données à partir du tableau croisé dynamique

La disposition définit la façon dont le tableau croisé dynamique est affiché dans la feuille de calcul. Cela signifie que `PivotLayout` l’objet contrôle les plages utilisées pour les éléments de tableau croisé dynamique. Utilisez les plages fournies par la disposition pour obtenir des données collectées et agrégées par le tableau croisé dynamique. En particulier, utilisez `PivotLayout.getDataBodyRange` cette information pour accéder aux données produites par le tableau croisé dynamique.

Le code suivant montre comment obtenir la dernière ligne des données du tableau croisé dynamique en passant par la disposition (le **total total des** **montants** vendus à la batterie de serveurs et la somme des **caisses vendues dans** l’exemple précédent). Ces valeurs sont ensuite additionées pour un total final, qui est affiché dans la cellule **E30** (en dehors du tableau croisé dynamique).

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // Get the totals for each data hierarchy from the layout.
    var range = pivotTable.layout.getDataBodyRange();
    var grandTotalRange = range.getLastRow();
    grandTotalRange.load("address");
    return context.sync().then(function () {
        // Sum the totals from the PivotTable data hierarchies and place them in a new range, outside of the PivotTable.
        var masterTotalRange = context.workbook.worksheets.getActiveWorksheet().getRange("E30");
        masterTotalRange.formulas = [["=SUM(" + grandTotalRange.address + ")"]];
    });
});
```

### <a name="layout-types"></a>Types de disposition

Les tableaux croisés dynamiques ont trois styles de disposition : Compact, Plan et Tabulaire. Nous avons vu le style compact dans les exemples précédents.

Les exemples suivants utilisent respectivement les styles plan et tabulaire. L’exemple de code montre comment faire un cycle entre les différentes dispositions.

#### <a name="outline-layout"></a>Disposition du plan

![Tableau croisé dynamique utilisant la disposition du plan.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a>Disposition tabulaire

![Tableau croisé dynamique utilisant la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a>Exemple de code de commutateur de type PivotLayout

```js
Excel.run(function (context) {
    // Change the PivotLayout.type to a new type.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    return context.sync().then(function () {
        // Cycle between the three layout types.
        if (pivotTable.layout.layoutType === "Compact") {
            pivotTable.layout.layoutType = "Outline";
        } else if (pivotTable.layout.layoutType === "Outline") {
            pivotTable.layout.layoutType = "Tabular";
        } else {
            pivotTable.layout.layoutType = "Compact";
        }
    
        return context.sync();
    });
});
```

### <a name="other-pivotlayout-functions"></a>Autres fonctions PivotLayout

Par défaut, les tableaux croisés dynamiques ajustent les tailles de lignes et de colonnes selon les besoins. Cette chose est effectuée lorsque le tableau croisé dynamique est actualisé. `PivotLayout.autoFormat` spécifie ce comportement. Les modifications de taille de ligne ou de colonne apportées par votre add-in sont persistantes `autoFormat` lorsqu’elles le `false` sont. En outre, les paramètres par défaut d’un tableau croisé dynamique conservent toute mise en forme personnalisée dans le tableau croisé dynamique (par exemple, les remplissages et les modifications de police). Définir `PivotLayout.preserveFormatting` pour appliquer le format par défaut lors de `false` l’actualisation.

A `PivotLayout` contrôle également les paramètres d’en-tête et de ligne totale, la façon dont les cellules de données vides sont affichées et les options de texte [de](https://support.microsoft.com/topic/44989b2a-903c-4d9a-b742-6a75b451c669) alt. La [référence PivotLayout](/javascript/api/excel/excel.pivotlayout) fournit une liste complète de ces fonctionnalités.

L’exemple de code suivant permet aux cellules de données vides d’afficher la chaîne, met en forme la plage de corps avec un alignement horizontal cohérent et garantit que les modifications de mise en forme restent même après l’actualisation du tableau croisé `"--"` dynamique.

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("Farm Sales");
    var pivotLayout = pivotTable.layout;

    // Set a default value for an empty cell in the PivotTable. This doesn't include cells left blank by the layout.
    pivotLayout.emptyCellText = "--";

    // Set the text alignment to match the rest of the PivotTable.
    pivotLayout.getDataBodyRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;

    // Ensure empty cells are filled with a default value.
    pivotLayout.fillEmptyCells = true;

    // Ensure that the format settings persist, even after the PivotTable is refreshed and recalculated.
    pivotLayout.preserveFormatting = true;
    return context.sync();
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

## <a name="filter-a-pivottable"></a>Filtrer un tableau croisé dynamique

La méthode principale de filtrage des données de tableau croisé dynamique est avec des filtres de tableau croisé dynamique. Les slicers offrent une autre méthode de filtrage moins flexible.

[Les filtres de tableau](/javascript/api/excel/excel.pivotfilters) croisé dynamique filtrent les données en fonction des quatre [catégories hiérarchiques](#hierarchies) d’un tableau croisé dynamique (filtres, colonnes, lignes et valeurs). Il existe quatre types de filtres de tableau croisé dynamique, ce qui permet le filtrage basé sur les dates du calendrier, l’comparaison des chaînes, la comparaison des nombres et le filtrage en fonction d’une entrée personnalisée.

[Les slicers](/javascript/api/excel/excel.slicer) peuvent être appliqués à la fois aux tableaux croisés dynamiques et aux tableaux Excel tableaux. Lorsqu’ils sont appliqués à un tableau croisé dynamique, les slicers fonctionnent comme un [pivotManualFilter](#pivotmanualfilter) et autorisent le filtrage en fonction d’une entrée personnalisée. Contrairement aux filtres de tableau croisé dynamique, les slicers ont [un Excel’interface utilisateur.](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) Avec la `Slicer` classe, vous créez ce composant d’interface utilisateur, gérez le filtrage et contrôlez son apparence visuelle.

### <a name="filter-with-pivotfilters"></a>Filtrer avec des filtres de tableau croisé dynamique

[Les filtres de tableau](/javascript/api/excel/excel.pivotfilters) croisé dynamique vous permettent de filtrer les données de tableau croisé dynamique en fonction des quatre [catégories hiérarchiques (filtres,](#hierarchies) colonnes, lignes et valeurs). Dans le modèle objet de tableau croisé dynamique, sont appliqués à un champ de tableau croisé dynamique , et chacun peut `PivotFilters` avoir un ou plusieurs [](/javascript/api/excel/excel.pivotfield) `PivotField` `PivotFilters` attribués . Pour appliquer des filtres de tableau croisé dynamique à un champ de tableau croisé dynamique, la [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) correspondante du champ doit être affectée à une catégorie de hiérarchie.

#### <a name="types-of-pivotfilters"></a>Types de filtres de tableau croisé dynamique

| Type de filtre | Objectif du filtre | Référence de l’API JavaScript pour Excel |
|:--- |:--- |:--- |
| DateFilter | Filtrage basé sur les dates du calendrier. | [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) |
| LabelFilter | Filtrage de comparaison de texte. | [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) |
| ManualFilter | Filtrage des entrées personnalisé. | [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) |
| ValueFilter | Filtrage de comparaison de nombres. | [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a>Créer un filtre de tableau croisé dynamique

Pour filtrer des données de tableau croisé dynamique avec `Pivot*Filter` un (par `PivotDateFilter` exemple, un ), appliquez le filtre à un champ [de tableau croisé dynamique.](/javascript/api/excel/excel.pivotfield) Les quatre exemples de code suivants montrent comment utiliser chacun des quatre types de filtres croisés dynamiques.

##### <a name="pivotdatefilter"></a>PivotDateFilter

Le premier exemple de code applique un [pivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) au champ de tableau croisé dynamique **date-mise** à jour, masquant les données antérieures au **08-2020-08-01**.

> [!IMPORTANT]
> A `Pivot*Filter` can’t be applied to a PivotField unless that field’s PivotHierarchy is assigned to a hierarchy category. Dans l’exemple de code suivant, le tableau croisé dynamique doit être ajouté à la catégorie du tableau croisé dynamique avant de pouvoir être `dateHierarchy` `rowHierarchies` utilisé pour le filtrage.

```js
Excel.run(function (context) {
    // Get the PivotTable and the date hierarchy.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    var dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date Updated");
    
    return context.sync().then(function () {
        // PivotFilters can only be applied to PivotHierarchies that are being used for pivoting.
        // If it's not already there, add "Date Updated" to the hierarchies.
        if (dateHierarchy.isNullObject) {
          dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date Updated"));
        }

        // Apply a date filter to filter out anything logged before August.
        var filterField = dateHierarchy.fields.getItem("Date Updated");
        var dateFilter = {
          condition: Excel.DateFilterCondition.afterOrEqualTo,
          comparator: {
            date: "2020-08-01",
            specificity: Excel.FilterDatetimeSpecificity.month
          }
        };
        filterField.applyFilter({ dateFilter: dateFilter });
        
        return context.sync();
    });
});
```

> [!NOTE]
> Les trois extraits de code suivants affichent uniquement des extraits spécifiques au filtre, au lieu d’appels `Excel.run` complets.

##### <a name="pivotlabelfilter"></a>PivotLabelFilter

Le deuxième extrait de code montre comment appliquer un [pivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) au champ de tableau croisé dynamique **de type,** en utilisant la propriété pour exclure les étiquettes qui commencent par la lettre `LabelFilterCondition.beginsWith` **L**.

```js
    // Get the "Type" field.
    var filterField = pivotTable.hierarchies.getItem("Type").fields.getItem("Type");

    // Filter out any types that start with "L" ("Lemons" and "Limes" in this case).
    var filter: Excel.PivotLabelFilter = {
      condition: Excel.LabelFilterCondition.beginsWith,
      substring: "L",
      exclusive: true
    };

    // Apply the label filter to the field.
    filterField.applyFilter({ labelFilter: filter });
```

##### <a name="pivotmanualfilter"></a>PivotManualFilter

Le troisième extrait de code applique un filtre manuel avec [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) au champ **Classification,** en filtrant les données qui n’incluent pas la classification **Organique**.

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a>PivotValueFilter

Pour comparer des nombres, utilisez un filtre de valeurs avec [PivotValueFilter,](/javascript/api/excel/excel.pivotvaluefilter)comme illustré dans l’extrait de code final. Le tableau croisé dynamique compare les données du champ pivot de la batterie de serveurs aux données du champ PivotField ventes de caisses, y compris uniquement les batteries dont la somme des caisses vendues dépasse la valeur `PivotValueFilter` **500**.  

```js
    // Get the "Farm" field.
    var filterField = pivotTable.hierarchies.getItem("Farm").fields.getItem("Farm");
    
    // Filter to only include rows with more than 500 wholesale crates sold.
    var filter: Excel.PivotValueFilter = {
      condition: Excel.ValueFilterCondition.greaterThan,
      comparator: 500,
      value: "Sum of Crates Sold Wholesale"
    };
    
    // Apply the value filter to the field.
    filterField.applyFilter({ valueFilter: filter });
```

#### <a name="remove-pivotfilters"></a>Supprimer des filtres de tableau croisé dynamique

Pour supprimer tous les filtres de tableau croisé dynamique, appliquez la méthode à chaque champ de tableau croisé dynamique, comme `clearAllFilters` illustré dans l’exemple de code suivant.

```js
Excel.run(function (context) {
    // Get the PivotTable.
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.hierarchies.load("name");
    
    return context.sync().then(function () {
        // Clear the filters on each PivotField.
        pivotTable.hierarchies.items.forEach(function (hierarchy) {
          hierarchy.fields.getItem(hierarchy.name).clearAllFilters();
        });
        return context.sync();
    });
});
```

### <a name="filter-with-slicers"></a>Filtrer avec des slicers

[Les slicers](/javascript/api/excel/excel.slicer) permettent de filtrer les données à partir d’Excel tableau croisé dynamique ou d’un tableau. Un slicer utilise les valeurs d’une colonne spécifiée ou d’un champ de tableau croisé dynamique pour filtrer les lignes correspondantes. Ces valeurs sont stockées en tant [qu’objets SlicerItem](/javascript/api/excel/excel.sliceritem) dans `Slicer` le . Votre add-in peut ajuster ces filtres, tout comme les utilisateurs[(via Excel’interface utilisateur).](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) Le slicer se trouve au-dessus de la feuille de calcul dans la couche de dessin, comme illustré dans la capture d’écran suivante.

![Un slicer filtrant des données sur un tableau croisé dynamique.](../images/excel-slicer.png)

> [!NOTE]
> Les techniques décrites dans cette section se concentrent sur l’utilisation de slicers connectés à des tableaux croisés dynamiques. Les mêmes techniques s’appliquent également à l’utilisation de slicers connectés à des tables.

#### <a name="create-a-slicer"></a>Créer un slicer

Vous pouvez créer un slicer dans un workbook ou une feuille de calcul à l’aide `Workbook.slicers.add` de la méthode ou de la `Worksheet.slicers.add` méthode. Cela ajoute un slicer à [la SlicerCollection](/javascript/api/excel/excel.slicercollection) de l’objet `Workbook` ou `Worksheet` spécifié. La `SlicerCollection.add` méthode a trois paramètres :

- `slicerSource`: source de données sur laquelle repose le nouveau slicer. Il peut s’agit d’une chaîne , ou d’une chaîne représentant `PivotTable` `Table` le nom ou l’ID d’un `PivotTable` ou `Table` .
- `sourceField`: champ dans la source de données par lequel filtrer. Il peut s’agit d’une chaîne , ou d’une chaîne représentant `PivotField` `TableColumn` le nom ou l’ID d’un `PivotField` ou `TableColumn` .
- `slicerDestination`: feuille de calcul dans laquelle le nouveau slicer sera créé. Il peut s’agit `Worksheet` d’un objet ou du nom ou de l’ID d’un `Worksheet` . Ce paramètre est inutile lorsque le `SlicerCollection` paramètre est accessible via `Worksheet.slicers` . Dans ce cas, la feuille de calcul de la collection est utilisée comme destination.

L’exemple de code suivant ajoute un nouveau slicer à la feuille de calcul **Pivot.** La source du slicer  est le tableau croisé dynamique ventes de batterie de serveurs et filtre à l’aide des **données type.** Le slicer est également nommé **Fruit Slicer pour** référence ultérieure.

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

#### <a name="filter-items-with-a-slicer"></a>Filtrer des éléments avec un slicer

Le slicer filtre le tableau croisé dynamique avec les éléments du `sourceField` . La `Slicer.selectItems` méthode définit les éléments qui restent dans le slicer. Ces éléments sont transmis à la méthode en tant `string[]` que , représentant les clés des éléments. Toutes les lignes contenant ces éléments restent dans l’agrégation du tableau croisé dynamique. Appels suivants `selectItems` pour définir la liste sur les touches spécifiées dans ces appels.

> [!NOTE]
> Si un élément qui ne se trouve pas dans la source de données est transmis, une `Slicer.selectItems` `InvalidArgument` erreur est lancée. Le contenu peut être vérifié par le biais de la propriété, qui est `Slicer.slicerItems` un [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).

L’exemple de code suivant montre trois éléments sélectionnés pour le slicer : **Sella,** **Tilleul** et **Orange**.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

Pour supprimer tous les filtres du slicer, utilisez la `Slicer.clearFilters` méthode, comme illustré dans l’exemple suivant.

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a>Style et mise en forme d’un slicer

Vous pouvez ajuster les paramètres d’affichage d’un slicer par le biais de `Slicer` propriétés. L’exemple de code suivant définit le style sur **SlicerStyleLight6,** définit le texte en haut du slicer sur **Types** de fruit, place le slicer à la position **(395, 15)** sur la couche de dessin et définit la taille du slicer à **135 x 150** pixels.

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

#### <a name="delete-a-slicer"></a>Supprimer un slicer

Pour supprimer un slicer, appelez la `Slicer.delete` méthode. L’exemple de code suivant supprime le premier slicer de la feuille de calcul actuelle.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a>Modifier la fonction d’agrégation

Les hiérarchies de données ont leurs valeurs agrégées. Pour les jeux de données de nombres, il s’agit d’une somme par défaut. La `summarizeBy` propriété définit ce comportement en fonction d’un type [AggregationFunction.](/javascript/api/excel/excel.aggregationfunction)

Les types de fonctions d’agrégation actuellement pris en charge sont `Sum` , , , , , , , , `Count` et `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` (par défaut).

Les exemples de code suivants modifient l’agrégation en moyenne des données.

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

## <a name="change-calculations-with-a-showasrule"></a>Modifier les calculs avec une méthode ShowAsRule

Par défaut, les tableaux croisés dynamiques agrègent les données de leurs hiérarchies de lignes et de colonnes indépendamment. Un [objet ShowAsRule](/javascript/api/excel/excel.showasrule) modifie la hiérarchie de données en valeurs de sortie basées sur d’autres éléments du tableau croisé dynamique.

`ShowAsRule`L’objet possède trois propriétés :

- `calculation`: type de calcul relatif à appliquer à la hiérarchie de données (la valeur par défaut est `none` ).
- `baseField`: [PivotField dans](/javascript/api/excel/excel.pivotfield) la hiérarchie contenant les données de base avant l’application du calcul. Étant Excel tableaux croisés dynamiques ont un mappage un-à-un de hiérarchie à champ, vous utiliserez le même nom pour accéder à la hiérarchie et au champ.
- `baseItem`: Tableau [croisé dynamique individuel comparé](/javascript/api/excel/excel.pivotitem) aux valeurs des champs de base en fonction du type de calcul. Tous les calculs ne nécessitent pas ce champ.

L’exemple suivant définit le calcul de la hiérarchie de données Somme des **caisses vendues** à la batterie de serveurs comme un pourcentage du total des colonnes.
Nous voulons toujours que la granularité s’étende au niveau du type de fruit. Nous allons donc utiliser la hiérarchie de ligne **Type** et son champ sous-jacent.
La batterie  de serveurs est également la première hiérarchie de ligne de l’exemple, de sorte que le nombre total d’entrées de la batterie de serveurs affiche également le pourcentage que chaque batterie de serveurs est responsable de la production.

![Tableau croisé dynamique montrant les pourcentages de ventes de fruit par rapport au total global des batteries de serveurs individuelles et des types de fruit individuels au sein de chaque batterie de serveurs.](../images/excel-pivots-showas-percentage.png)

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

L’exemple précédent a fixé le calcul à la colonne, par rapport au champ d’une hiérarchie de lignes individuelle. Lorsque le calcul est lié à un élément individuel, utilisez la `baseItem` propriété.

L’exemple suivant montre le `differenceFrom` calcul. Il affiche la différence entre les entrées de hiérarchie des données de ventes de la batterie de serveurs par rapport à celles des **batteries de serveurs A**.
Il s’agit d’une batterie de serveurs, ce qui nous permet de voir les différences entre les autres batteries de serveurs, ainsi que les répartitions pour chaque type de fruit comme ( Le type est également une hiérarchie de lignes dans `baseField` cet exemple). 

![Tableau croisé dynamique montrant les différences de ventes de fruit entre les « batteries de serveurs » et les autres. Cela montre à la fois la différence entre les ventes totales de fruit des batteries de serveurs et les ventes de types de fruit. Si « A Farms » n’a pas vendu un type particulier de fruit, « #N/A » s’affiche.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="see-also"></a>Voir aussi

- [Modèle d’objet JavaScript Excel dans les compléments Office](excel-add-ins-core-concepts.md)
- [Excel Référence de l’API JavaScript](/javascript/api/excel)
