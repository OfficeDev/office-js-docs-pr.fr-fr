---
title: Utilisation de tableaux croisés dynamiques à l’aide de l’API JavaScript pour Excel
description: Utilisez l'API JavaScript pour Excel afin de créer des tableaux croisés dynamiques et d’interagir avec leurs composants.
ms.date: 09/21/2018
ms.openlocfilehash: 00dd982d4ba4de0db34277cd546b572d4394e258
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459279"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="62c0c-103">Utilisation de tableaux croisés dynamiques à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="62c0c-103">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="62c0c-p101">Les tableaux croisés dynamiques rationalisent les jeux de données plus volumineux. Ils permettent la manipulation rapide des données groupées. L’API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d’interagir avec leurs composants.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p101">PivotTables streamline larger data sets. They allow the quick manipulation of grouped data. The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="62c0c-p102">Si vous ne connaissez pas les fonctionnalités des tableaux croisés dynamiques, envisagez de les découvrir en tant qu’utilisateur final. Consultez la rubrique [Créer un tableau croisé dynamique pour analyser les données d’une feuille de calcul](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) afin d’obtenir une présentation de ces outils.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p102">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user. See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="62c0c-p103">Cet article fournit des exemples de code pour des scénarios courants. Pour améliorer votre compréhension de l'API Tableau croisé dynamique, consultez [**Tableau croisé dynamique**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) et [**Collection Tableau croisé dynamique**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="62c0c-p103">This article provides code samples for common scenarios. To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="62c0c-111">Les tableaux croisés dynamiques créés avec OLAP ne sont pas actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="62c0c-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="62c0c-112">Hiérarchies</span><span class="sxs-lookup"><span data-stu-id="62c0c-112">Hierarchies</span></span>

<span data-ttu-id="62c0c-p104">Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie : ligne, colonne, données et filtre. Les données suivantes décrivant des ventes de fruits provenant de différentes fermes seront utilisées dans cet article.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p104">PivotTables are organized based on four hierarchy categories: row, column, data, and filter. The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Collection de ventes de fruits de différents types provenant de différentes fermes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="62c0c-p105">Ces données ont cinq hiérarchies : **Fermes**, **Type**, **Classification**, **Caisses vendues à la ferme** et **Caisses vendues en gros**. Chaque hiérarchie ne peut exister que dans l’une des quatre catégories. Si le \*\* Type\*\* est ajouté aux hiérarchies de colonnes puis aux hiérarchies de lignes, il ne reste que dans ces dernières.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p105">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**. Each hierarchy can only exist in one of the four categories. If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="62c0c-p106">Les hiérarchies de lignes et de colonnes définissent la façon dont les données sont regroupées. Par exemple, une hiérarchie de lignes de **Fermes** regroupe tous les jeux de données provenant de la même ferme. Le choix entre la hiérarchie de lignes et de colonnes définit l’orientation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p106">Row and column hierarchies define how data will be grouped. For example, a row hierarchy of **Farms** will group together all the data sets from the same farm. The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="62c0c-p107">Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de lignes et de colonnes. Un tableau croisé dynamique avec une hiérarchie de lignes de **Fermes** et une hiérarchie de données de **Caisses vendues en gros** affiche la somme totale (par défaut) de tous les différents fruits pour chaque ferme.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p107">Data hierarchies are the values to be aggregated based on the row and column hierarchies. A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="62c0c-p108">Les hiérarchies de filtres incluent ou excluent les données provenant du pivot en fonction des valeurs dans ce type filtré. Une hiérarchie de filtres de **Classification** avec le type **Biologique** sélectionné n’affiche que les données pour les fruits biologiques.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p108">Filter hierarchies include or exclude data from the pivot based on values within that filtered type. A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="62c0c-p109">Voici à nouveau les données des fermes, à côté d’un tableau croisé dynamique. Le tableau croisé dynamique utilise **Ferme** et **Type** en tant que hiérarchies de lignes, **Caisses vendues à la ferme** et **Caisses vendues en gros** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut de la somme) et **Classification** en tant que hiérarchie de filtres (avec **Biologique** sélectionné).</span><span class="sxs-lookup"><span data-stu-id="62c0c-p109">Here is the farm data again, alongside a PivotTable. The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Sélection de données de ventes de fruits à côté d'un tableau croisé dynamique avec les hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="62c0c-p110">Ce tableau croisé dynamique peut être généré via l’API JavaScript ou l’interface utilisateur d’Excel. Les deux options permettent une manipulation plus poussée via les compléments.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p110">This PivotTable could be generated through the JavaScript API or through the Excel UI. Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="62c0c-131">Créer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="62c0c-131">Create a PivotTable with Range objects</span></span>

<span data-ttu-id="62c0c-p111">Les tableaux croisés dynamiques nécessitent un nom, une source et une destination. La source peut être une adresse de plage ou un nom de table (passés comme un type `Range`, `string` ou `Table`). La destination est une adresse de plage (donnée sous forme de `Range`  ou `string`). Les exemples suivants illustrent diverses techniques de création de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p111">PivotTables need a name, source, and destination. The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type). The destination is a range address (given as either a `Range` or `string`). The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="62c0c-136">Créer un tableau croisé dynamique avec des adresses de plages</span><span class="sxs-lookup"><span data-stu-id="62c0c-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="62c0c-137">Créer un tableau croisé dynamique avec des objets Plage</span><span class="sxs-lookup"><span data-stu-id="62c0c-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="62c0c-138">Créer un tableau croisé dynamique au niveau classeur</span><span class="sxs-lookup"><span data-stu-id="62c0c-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="62c0c-139">Utiliser un tableau croisé dynamique existant</span><span class="sxs-lookup"><span data-stu-id="62c0c-139">Use an existing PivotTable</span></span>

<span data-ttu-id="62c0c-140">Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection Tableau croisé dynamique du classeur ou des feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="62c0c-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="62c0c-p112">Le code suivant récupère le premier tableau croisé dynamique du classeur. Il donne ensuite un nom à la table pour faciliter les références ultérieures.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p112">The following code gets the first PivotTable in the workbook. It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="62c0c-143">Ajouter des lignes et des colonnes à un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="62c0c-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="62c0c-144">Les lignes et les colonnes regroupent les données autour des valeurs de ces champs.</span><span class="sxs-lookup"><span data-stu-id="62c0c-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="62c0c-p113">L’ajout de la colonne **Ferme** regroupe toutes les ventes autour de chaque ferme. L'ajout des lignes **Type**  et **Classification**  décompose davantage les données en fonction du fruit vendu et de sa classification biologique ou non.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p113">Adding the **Farm** column pivots all the sales around each farm. Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="62c0c-148">Un tableau croisé dynamique peut également ne contenir que des lignes ou que des colonnes.</span><span class="sxs-lookup"><span data-stu-id="62c0c-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="62c0c-149">Ajouter des hiérarchies de données au tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="62c0c-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="62c0c-p114">Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes. L’ajout des hiérarchies de données de **Caisses vendues à la ferme** et **Caisses vendues en gros** donne les sommes de ces chiffres pour chaque ligne et chaque colonne.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p114">Data hierarchies fill the PivotTable with information to combine based on the rows and columns. Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="62c0c-152">Dans l’exemple, **Ferme** et **Type** sont des lignes, tandis que les ventes de caisses sont les données.</span><span class="sxs-lookup"><span data-stu-id="62c0c-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

![Tableau croisé dynamique affichant les ventes totales des différents fruits en fonction de la ferme d'où ils proviennent.](../images/excel-pivots-data-hierarchy.png)

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

## <a name="change-aggregation-function"></a><span data-ttu-id="62c0c-154">Modifier la fonction d’agrégation</span><span class="sxs-lookup"><span data-stu-id="62c0c-154">Change aggregation function</span></span>

<span data-ttu-id="62c0c-p115">Les hiérarchies de données voient leurs valeurs agrégées. Pour les jeux de données de nombres, il s’agit d’une somme par défaut .La propriété `summarizeBy` définit ce comportement en fonction d’un type `AggregrationFunction`.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p115">Data hierarchies have their values aggregated. For datasets of numbers, this is a sum by default. The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="62c0c-158">Les types de fonctions d’agrégation actuellement prises en charge sont `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` et `Automatic` (par défaut).</span><span class="sxs-lookup"><span data-stu-id="62c0c-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="62c0c-159">Les exemples de code suivants modifient l’agrégation en moyennes des données.</span><span class="sxs-lookup"><span data-stu-id="62c0c-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="62c0c-160">Modifier les calculs avec une propriété ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="62c0c-160">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="62c0c-p116">Les tableaux croisés dynamiques agrègent par défaut les données de leurs hiérarchies de ligne et de colonne de manière indépendante. Un objet `ShowAsRule` modifie la hiérarchie de données pour produire des valeurs en fonction des autres éléments du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p116">PivotTables, by default, aggregate the data of their row and column hierarchies independently. A `ShowAsRule` changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="62c0c-163">L’objet `ShowAsRule` contient possède trois propriétés :</span><span class="sxs-lookup"><span data-stu-id="62c0c-163">The `ShowAsRule` object has three properties:</span></span>
-   <span data-ttu-id="62c0c-164">`calculation`: le type de calcul relatif à appliquer à la hiérarchie des données (la valeur par défaut est `none`).</span><span class="sxs-lookup"><span data-stu-id="62c0c-164">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="62c0c-p117">`baseField`: le champ dans la hiérarchie contenant les données de base avant le calcul est appliqué. L’objet `PivotField` porte généralement le même nom que sa hiérarchie parent.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p117">`baseField`: The field within the hierarchy containing the base data before the calculation is applied. The `PivotField` usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="62c0c-p118">`baseItem`:  l’élément individuel comparé aux valeurs des champs de base en fonction du type de calcul. Tous les calculs ne nécessitent pas ce champ.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p118">`baseItem`: The individual item compared against the values of the base fields based on the calculation type. Not all calculations require this field.</span></span>

<span data-ttu-id="62c0c-p119">L’exemple suivant définit le calcul de la hiérarchie de données de la **Somme des caisses vendues à la ferme**  comme un pourcentage du total de colonne. Nous voulons quand même que la granularité s’étende au niveau du type de fruits, nous allons donc utiliser la hiérarchie de ligne \*\* Type\*\* et son champ sous-jacent. L’exemple a également \*\* Ferme\*\* comme hiérarchie de la première ligne, afin que les entrées de total de la ferme affichent également le pourcentage que chaque ferme a la responsabilité de produire.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p119">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total. We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field. The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

![Un tableau croisé dynamique affichant les pourcentages des ventes de fruits par rapport à un total général pour les fermes individuelles et les types des fruits dans chaque ferme.](../images/excel-pivots-showas-percentage.png)

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

<span data-ttu-id="62c0c-p120">L’exemple précédent définit le calcul de la colonne, par rapport à une hiérarchie de ligne individuelle. Lorsque le calcul se rapporte à un élément individuel, utilisez la propriété  `baseItem`.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p120">The previous example set the calculation to the column, relative to an individual row hierarchy. When the calculation relates to an individual item, use the `baseItem` property.</span></span> 

<span data-ttu-id="62c0c-p121">L’exemple ci-dessous illustre le calcul `differenceFrom`. Il affiche la différence des entrées de la hiérarchie de données relative aux ventes de caisses des fermes par rapport à celles des « Fermes A ». La propriété `baseField`  est **Ferme**, de sorte que nous voir les différences entre les autres fermes, ainsi que des répartitions pour chaque type de fruits comparables (**Type** est également une hiérarchie de ligne dans cet exemple).</span><span class="sxs-lookup"><span data-stu-id="62c0c-p121">The following example shows the `differenceFrom` calculation. It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”. The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Un tableau croisé dynamique affichant les différences des ventes de fruits entre les « Fermes A » et les autres. Il affiche à la fois la différence dans les ventes de fruits totales des fermes et les ventes des types de fruits. Si les « Fermes A » n’ont pas vendu un type de fruit particulier, « #N/A » s’affiche.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="62c0c-181">Dispositions des tableaux croisés dynamiques</span><span class="sxs-lookup"><span data-stu-id="62c0c-181">PivotTable layouts</span></span>

<span data-ttu-id="62c0c-p123">La disposition d’un tableau croisé dynamique définit le positionnement des hiérarchies et de leurs données. Accéder à la disposition permet de déterminer les plages de stockage des données.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p123">A PivotTable layout defines the placement of hierarchies and their data. You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="62c0c-184">Le diagramme suivant présente la correspondance des appels de fonction de disposition avec les plages du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="62c0c-184">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Diagramme présentant les sections d’un tableau croisé dynamique renvoyées par les fonctions de récupération de plage de la disposition.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="62c0c-p124">Le code suivant indique comment récupérer la dernière ligne des données de tableau croisé dynamique via la disposition. Ces valeurs sont ensuite additionnées pour obtenir un total général.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p124">The following code demonstrates how to get the last row of the PivotTable data by going through the layout. Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="62c0c-p125">Les tableaux croisés dynamiques ont trois styles de disposition : Compact, Plan et Tabulaire. Nous avons vu le style compact dans les exemples précédents.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p125">PivotTables have three layout styles: Compact, Outline, and Tabular. We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="62c0c-p126">Les exemples suivants utilisent respectivement le style plan et tabulaire. L’exemple de code montre comment passer d’une disposition à une autre.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p126">The following examples use the outline and tabular styles, respectively. The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="62c0c-192">Disposition Plan</span><span class="sxs-lookup"><span data-stu-id="62c0c-192">Outline layout</span></span>

![Tableau croisé dynamique utilisant la disposition plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="62c0c-194">Disposition Tabulaire</span><span class="sxs-lookup"><span data-stu-id="62c0c-194">Tabular layout</span></span>

![Tableau croisé dynamique utilisant la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="62c0c-196">Modifier les noms des hiérarchies</span><span class="sxs-lookup"><span data-stu-id="62c0c-196">Change hierarchy names</span></span>

<span data-ttu-id="62c0c-p127">Les champs des hiérarchies peuvent être modifiés. Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.</span><span class="sxs-lookup"><span data-stu-id="62c0c-p127">Hierarchy fields are editable. The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="62c0c-199">Supprimer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="62c0c-199">Delete a PivotTable</span></span>

<span data-ttu-id="62c0c-200">Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.</span><span class="sxs-lookup"><span data-stu-id="62c0c-200">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="62c0c-201">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="62c0c-201">See also</span></span>

- [<span data-ttu-id="62c0c-202">Concepts fondamentaux de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="62c0c-202">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="62c0c-203">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="62c0c-203">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
