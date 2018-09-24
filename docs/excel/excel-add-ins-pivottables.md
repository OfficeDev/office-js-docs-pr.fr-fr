---
title: Utilisation de tableaux croisés dynamiques à l’aide de l’API JavaScript pour Excel
description: Utilisez l'API JavaScript pour Excel afin de créer des tableaux croisés dynamiques et d’interagir avec leurs composants.
ms.date: 09/21/2018
ms.openlocfilehash: b8704389ced3686858f488b2a50f80c22b1b8bd6
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967668"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="971db-103">Utilisation de tableaux croisés dynamiques à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="971db-103">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="971db-104">Les tableaux croisés dynamiques rationalisent les jeux de données plus volumineux.</span><span class="sxs-lookup"><span data-stu-id="971db-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="971db-105">Ils permettent la manipulation rapide des données groupées.</span><span class="sxs-lookup"><span data-stu-id="971db-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="971db-106">L'API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d'interagir avec leurs composants.</span><span class="sxs-lookup"><span data-stu-id="971db-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="971db-107">Si vous ne connaissez pas les fonctionnalités des tableaux croisés dynamiques, envisagez de les découvrir en tant qu’utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="971db-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="971db-108">Consultez [Créer un tableau croisé dynamique pour analyser les données d'une feuille de calcul](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) afin d'obtenir une présentation de ces outils.</span><span class="sxs-lookup"><span data-stu-id="971db-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="971db-109">Cet article fournit des exemples de code pour des scénarios courants.</span><span class="sxs-lookup"><span data-stu-id="971db-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="971db-110">Pour améliorer votre compréhension de l'API Tableau croisé dynamique, consultez [**Tableau croisé dynamique**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) et [**Collection Tableau croisé dynamique**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="971db-110">To further your understanding of the PivotTable API, see [**PivotTable**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](https://docs.microsoft.com/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="971db-111">Les tableaux croisés dynamiques créés avec OLAP ne sont pas actuellement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="971db-111">PivotTables created with OLAP are not currently supported.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="971db-112">Hiérarchies</span><span class="sxs-lookup"><span data-stu-id="971db-112">Hierarchies</span></span>

<span data-ttu-id="971db-113">Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie : ligne, colonne, données et filtre.</span><span class="sxs-lookup"><span data-stu-id="971db-113">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="971db-114">Les données suivantes décrivant des ventes de fruits provenant de différentes fermes seront utilisées dans cet article.</span><span class="sxs-lookup"><span data-stu-id="971db-114">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Collection de ventes de fruits de différents types provenant de différentes fermes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="971db-116">Ces données ont cinq hiérarchies : **Fermes**, **Type**, **Classification**, **Caisses vendues à la ferme** et **Caisses vendues en gros**.</span><span class="sxs-lookup"><span data-stu-id="971db-116">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="971db-117">Chaque hiérarchie ne peut exister que dans l’une des quatre catégories.</span><span class="sxs-lookup"><span data-stu-id="971db-117">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="971db-118">Si le **Type** est ajouté aux hiérarchies de colonnes puis aux hiérarchies de lignes, il ne reste que dans ces dernières.</span><span class="sxs-lookup"><span data-stu-id="971db-118">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="971db-119">Les hiérarchies de lignes et de colonnes définissent la façon dont les données sont regroupées.</span><span class="sxs-lookup"><span data-stu-id="971db-119">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="971db-120">Par exemple, une hiérarchie de lignes de **Fermes** regroupe tous les jeux de données provenant de la même ferme.</span><span class="sxs-lookup"><span data-stu-id="971db-120">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="971db-121">Le choix entre la hiérarchie de lignes et de colonnes définit l'orientation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="971db-121">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="971db-122">Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de lignes et de colonnes.</span><span class="sxs-lookup"><span data-stu-id="971db-122">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="971db-123">Un tableau croisé dynamique avec une hiérarchie de lignes de **Fermes** et une hiérarchie de données de **Caisses vendues en gros** affiche la somme totale (par défaut) de tous les différents fruits pour chaque ferme.</span><span class="sxs-lookup"><span data-stu-id="971db-123">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="971db-124">Les hiérarchies de filtres incluent ou excluent les données provenant du pivot en fonction des valeurs dans ce type filtré.</span><span class="sxs-lookup"><span data-stu-id="971db-124">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="971db-125">Une hiérarchie de filtres de **Classification** avec le type **Biologique** sélectionné n'affiche que les données pour les fruits biologiques.</span><span class="sxs-lookup"><span data-stu-id="971db-125">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="971db-126">Voici à nouveau les données des fermes, à côté d’un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="971db-126">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="971db-127">Le tableau croisé dynamique utilise **Ferme** et **Type** en tant que hiérarchies de lignes, **Caisses vendues à la ferme** et **Caisses vendues en gros** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut de la somme) et **Classification** en tant que hiérarchie de filtres (avec **Biologique** sélectionné).</span><span class="sxs-lookup"><span data-stu-id="971db-127">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Sélection de données de ventes de fruits à côté d'un tableau croisé dynamique avec les hiérarches de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="971db-129">Ce tableau croisé dynamique peut être généré via l’API JavaScript ou l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="971db-129">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="971db-130">Les deux options permettent une manipulation plus poussée via les compléments.</span><span class="sxs-lookup"><span data-stu-id="971db-130">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="971db-131">Créer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="971db-131">Create a PivotTable with Range objects</span></span>

<span data-ttu-id="971db-132">Les tableaux croisés dynamiques nécessitent un nom, une source et une destination.</span><span class="sxs-lookup"><span data-stu-id="971db-132">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="971db-133">La source peut être une adresse de plage ou un nom de table (passés comme un type `Range`, `string` ou `Table`).</span><span class="sxs-lookup"><span data-stu-id="971db-133">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="971db-134">La destination est une adresse de plage (donnée sous forme de `Range` ou `string`).</span><span class="sxs-lookup"><span data-stu-id="971db-134">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="971db-135">Les exemples suivants présentent différentes techniques de création de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="971db-135">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="971db-136">Créer un tableau croisé dynamique avec des adresses de plages</span><span class="sxs-lookup"><span data-stu-id="971db-136">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="971db-137">Créer un tableau croisé dynamique avec des objets Plage</span><span class="sxs-lookup"><span data-stu-id="971db-137">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="971db-138">Créer un tableau croisé dynamique au niveau classeur</span><span class="sxs-lookup"><span data-stu-id="971db-138">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="971db-139">Utiliser un tableau croisé dynamique existant</span><span class="sxs-lookup"><span data-stu-id="971db-139">Use an existing PivotTable</span></span>

<span data-ttu-id="971db-140">Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection Tableau croisé dynamique du classeur ou des feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="971db-140">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="971db-141">Le code suivant récupère le premier tableau croisé dynamique du classeur.</span><span class="sxs-lookup"><span data-stu-id="971db-141">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="971db-142">Il donne ensuite un nom à la table pour faciliter les références ultérieures.</span><span class="sxs-lookup"><span data-stu-id="971db-142">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="971db-143">Ajouter des lignes et des colonnes à un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="971db-143">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="971db-144">Les lignes et les colonnes regroupent les données autour des valeurs de ces champs.</span><span class="sxs-lookup"><span data-stu-id="971db-144">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="971db-145">L'ajout de la colonne **Ferme** regroupe toutes les ventes autour de chaque ferme.</span><span class="sxs-lookup"><span data-stu-id="971db-145">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="971db-146">L'ajout des lignes **Type** et **Classification** décompose davantage les données en fonction du fruit vendu et de sa classification biologique ou non.</span><span class="sxs-lookup"><span data-stu-id="971db-146">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="971db-148">Un tableau croisé dynamique peut également ne contenir que des lignes ou que des colonnes.</span><span class="sxs-lookup"><span data-stu-id="971db-148">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="971db-149">Ajouter des hiérarchies de données au tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="971db-149">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="971db-150">Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes.</span><span class="sxs-lookup"><span data-stu-id="971db-150">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="971db-151">L'ajout des hiérarchies de données de **Caisses vendues à la ferme** et **Caisses vendues en gros** donne les sommes de ces chiffres pour chaque ligne et chaque colonne.</span><span class="sxs-lookup"><span data-stu-id="971db-151">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="971db-152">Dans l’exemple, **Ferme** et **Type** sont des lignes, tandis que les ventes de caisses sont les données.</span><span class="sxs-lookup"><span data-stu-id="971db-152">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="971db-154">Modifier la fonction d’agrégation</span><span class="sxs-lookup"><span data-stu-id="971db-154">Change aggregation function</span></span>

<span data-ttu-id="971db-155">Les hiérarchies de données voient leurs valeurs agrégées.</span><span class="sxs-lookup"><span data-stu-id="971db-155">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="971db-156">Pour les jeux de données de nombres, il s’agit d’une somme par défaut.</span><span class="sxs-lookup"><span data-stu-id="971db-156">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="971db-157">Le `summarizeBy` propriété définit ce comportement en fonction d'un `AggregrationFunction` type.</span><span class="sxs-lookup"><span data-stu-id="971db-157">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="971db-158">Les types de fonctions d’agrégation actuellement prises en charge sont `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` et `Automatic` (par défaut).</span><span class="sxs-lookup"><span data-stu-id="971db-158">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="971db-159">Les exemples de code suivants modifient l'agrégation en moyennes des données.</span><span class="sxs-lookup"><span data-stu-id="971db-159">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="pivottable-layouts"></a><span data-ttu-id="971db-160">Dispositions des tableaux croisés dynamiques</span><span class="sxs-lookup"><span data-stu-id="971db-160">PivotTable layouts</span></span>

<span data-ttu-id="971db-161">La disposition d'un tableau croisé dynamique définit le positionnement des hiérarchies et de leurs données.</span><span class="sxs-lookup"><span data-stu-id="971db-161">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="971db-162">Accéder à la disposition permet de déterminer les plages de stockage des données.</span><span class="sxs-lookup"><span data-stu-id="971db-162">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="971db-163">Le diagramme suivant présente la correspondance des appels de fonction de disposition avec les plages du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="971db-163">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Diagramme présentant les sections d’un tableau croisé dynamique renvoyées par les fonctions de récupération de plage de la disposition.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="971db-165">Le code suivant indique comment récupérer la dernière ligne des données de tableau croisé dynamique via la disposition.</span><span class="sxs-lookup"><span data-stu-id="971db-165">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="971db-166">Ces valeurs sont ensuite additionnées pour obtenir un total général.</span><span class="sxs-lookup"><span data-stu-id="971db-166">Those values are then summed together for a grand total.</span></span>


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

<span data-ttu-id="971db-167">Les tableaux croisés dynamiques ont trois styles de disposition : Compact, Plan et Tabulaire.</span><span class="sxs-lookup"><span data-stu-id="971db-167">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="971db-168">Nous avons vu le style compact dans les exemples précédents.</span><span class="sxs-lookup"><span data-stu-id="971db-168">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="971db-169">Les exemples suivants utilisent respectivement le style plan et tabulaire.</span><span class="sxs-lookup"><span data-stu-id="971db-169">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="971db-170">L’exemple de code montre comment passer d'une disposition à une autre.</span><span class="sxs-lookup"><span data-stu-id="971db-170">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="971db-171">Disposition Plan</span><span class="sxs-lookup"><span data-stu-id="971db-171">Outline layout</span></span>

![Tableau croisé dynamique utilisant la disposition plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="971db-173">Disposition Tabulaire</span><span class="sxs-lookup"><span data-stu-id="971db-173">Tabular layout</span></span>

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="971db-175">Modifier les noms des hiérarchies</span><span class="sxs-lookup"><span data-stu-id="971db-175">Change hierarchy names</span></span>

<span data-ttu-id="971db-176">Les champs des hiérarchies peuvent être modifiés.</span><span class="sxs-lookup"><span data-stu-id="971db-176">Hierarchy fields are editable.</span></span> <span data-ttu-id="971db-177">Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.</span><span class="sxs-lookup"><span data-stu-id="971db-177">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="971db-178">Supprimer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="971db-178">Delete a PivotTable</span></span>

<span data-ttu-id="971db-179">Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.</span><span class="sxs-lookup"><span data-stu-id="971db-179">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="971db-180">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="971db-180">See also</span></span>

- [<span data-ttu-id="971db-181">Concepts de base de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="971db-181">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="971db-182">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="971db-182">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/api/excel)
