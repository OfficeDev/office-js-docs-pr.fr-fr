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
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="433a1-103">Utilisation des tableaux croisés dynamiques avec l'API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="433a1-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="433a1-104">Les tableaux croisés dynamiques rationalisent les grands ensembles de données.</span><span class="sxs-lookup"><span data-stu-id="433a1-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="433a1-105">Ils permettent la manipulation rapide des données groupées.</span><span class="sxs-lookup"><span data-stu-id="433a1-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="433a1-106">L'API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d'interagir avec leurs composants.</span><span class="sxs-lookup"><span data-stu-id="433a1-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span>

<span data-ttu-id="433a1-107">Si vous n'êtes pas familiarisé avec la fonctionnalité de tableaux croisés dynamiques, envisagez de les explorer comme un utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="433a1-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span> <span data-ttu-id="433a1-108">RePortez-vous à la rubrique [créer un tableau croisé dynamique pour analyser les données de feuille de calcul](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) pour obtenir une introduction à ces outils.</span><span class="sxs-lookup"><span data-stu-id="433a1-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="433a1-109">Cet article fournit des exemples de code pour les scénarios courants.</span><span class="sxs-lookup"><span data-stu-id="433a1-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="433a1-110">Pour mieux comprendre l'API PivotTable, consultez la rubrique [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).</span><span class="sxs-lookup"><span data-stu-id="433a1-110">To further your understanding of the PivotTable API, see [**PivotTable**](/javascript/api/excel/excel.pivottable) and [**PivotTableCollection**](/javascript/api/excel/excel.pivottable).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="433a1-111">Les tableaux croisés dynamiques créés avec OLAP ne sont actuellement pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="433a1-111">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="433a1-112">Il n'existe pas non plus de prise en charge de PowerPivot.</span><span class="sxs-lookup"><span data-stu-id="433a1-112">There is also no support for Power Pivot.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="433a1-113">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="433a1-113">Hierarchies</span></span>

<span data-ttu-id="433a1-114">Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie: ligne, colonne, données et filtre.</span><span class="sxs-lookup"><span data-stu-id="433a1-114">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="433a1-115">Les données suivantes décrivant les ventes de fruit de différentes batteries de serveurs seront utilisées tout au long de cet article.</span><span class="sxs-lookup"><span data-stu-id="433a1-115">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Collection de ventes de fruit de différents types de batteries de serveurs différentes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="433a1-117">Ces données ont cinq hiérarchies **: batteries de serveurs**, **type**, **classification**, **caisses vendues à la batterie de serveurs**et **caisses vendues en gros**.</span><span class="sxs-lookup"><span data-stu-id="433a1-117">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="433a1-118">Chaque hiérarchie peut uniquement exister dans l'une des quatre catégories.</span><span class="sxs-lookup"><span data-stu-id="433a1-118">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="433a1-119">Si le **type** est ajouté aux hiérarchies de colonnes, puis ajouté aux hiérarchies de lignes, il n'est conservé que dans ce dernier.</span><span class="sxs-lookup"><span data-stu-id="433a1-119">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="433a1-120">Les hiérarchies de ligne et de colonne définissent le mode de regroupement des données.</span><span class="sxs-lookup"><span data-stu-id="433a1-120">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="433a1-121">Par exemple, une hiérarchie de lignes de **batteries de serveurs** regroupe tous les jeux de données de la même batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="433a1-121">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="433a1-122">Le choix entre la hiérarchie de ligne et de colonne définit l'orientation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="433a1-122">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="433a1-123">Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de ligne et de colonne.</span><span class="sxs-lookup"><span data-stu-id="433a1-123">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="433a1-124">Un tableau croisé dynamique avec une hiérarchie de lignes de **batteries de serveurs** et une hiérarchie de données de grossistes **vendus en gros** indique le total de tous les fruits de chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="433a1-124">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="433a1-125">Les hiérarchies de filtre incluent ou excluent les données du tableau croisé dynamique en fonction des valeurs contenues dans ce type filtré.</span><span class="sxs-lookup"><span data-stu-id="433a1-125">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="433a1-126">Une hiérarchie de filtrage de **classification** avec le type **Organic** Selected affiche uniquement les données pour les fruits organiques.</span><span class="sxs-lookup"><span data-stu-id="433a1-126">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="433a1-127">Voici les données de la batterie de serveurs à nouveau, ainsi qu'un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="433a1-127">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="433a1-128">Le tableau croisé dynamique utilise la **batterie de serveurs** et le **type** comme hiérarchies de lignes, les **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** en tant que hiérarchies de données (avec la fonction d'agrégation par défaut Sum) et une **classification** en tant que filtre hiérarchie (avec l'option **Organic** sélectionnée).</span><span class="sxs-lookup"><span data-stu-id="433a1-128">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Sélection de données sur les ventes de fruit en regard d'un tableau croisé dynamique avec des hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="433a1-130">Ce tableau croisé dynamique peut être généré via l'API JavaScript ou via l'interface utilisateur d'Excel.</span><span class="sxs-lookup"><span data-stu-id="433a1-130">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="433a1-131">Ces deux options permettent une manipulation supplémentaire via les compléments.</span><span class="sxs-lookup"><span data-stu-id="433a1-131">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="433a1-132">Créer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="433a1-132">Create a PivotTable</span></span>

<span data-ttu-id="433a1-133">Les tableaux croisés dynamiques nécessitent un nom, une source et une destination.</span><span class="sxs-lookup"><span data-stu-id="433a1-133">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="433a1-134">La source peut être une adresse de plage ou un nom de table ( `Range`transmis `string`en tant `Table` que type, ou type).</span><span class="sxs-lookup"><span data-stu-id="433a1-134">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="433a1-135">La destination est une adresse de plage (sous la forme `Range` a `string`ou).</span><span class="sxs-lookup"><span data-stu-id="433a1-135">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="433a1-136">Les exemples suivants illustrent différentes techniques de création de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="433a1-136">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="433a1-137">Créer un tableau croisé dynamique avec des adresses de plage</span><span class="sxs-lookup"><span data-stu-id="433a1-137">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="433a1-138">Création d'un tableau croisé dynamique avec des objets Range</span><span class="sxs-lookup"><span data-stu-id="433a1-138">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="433a1-139">Création d'un tableau croisé dynamique au niveau du classeur</span><span class="sxs-lookup"><span data-stu-id="433a1-139">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="433a1-140">Utiliser un tableau croisé dynamique existant</span><span class="sxs-lookup"><span data-stu-id="433a1-140">Use an existing PivotTable</span></span>

<span data-ttu-id="433a1-141">Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection de tableau croisé dynamique du classeur ou de feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="433a1-141">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="433a1-142">Le code suivant obtient le premier tableau croisé dynamique du classeur.</span><span class="sxs-lookup"><span data-stu-id="433a1-142">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="433a1-143">Il donne ensuite un nom à la table pour une référence facile.</span><span class="sxs-lookup"><span data-stu-id="433a1-143">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="433a1-144">Ajouter des lignes et des colonnes à un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="433a1-144">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="433a1-145">Lignes et colonnes tableau croisé dynamique des données autour de ces valeurs.</span><span class="sxs-lookup"><span data-stu-id="433a1-145">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="433a1-146">L'ajout de la colonne **batterie de serveurs** pivote toutes les ventes autour de chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="433a1-146">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="433a1-147">L'ajout des lignes de type et de **classification** répartit davantage les données en fonction des fruits vendus et s'il s'agit d'un **type** Organic ou non.</span><span class="sxs-lookup"><span data-stu-id="433a1-147">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="433a1-149">Vous pouvez également utiliser un tableau croisé dynamique avec uniquement des lignes ou des colonnes.</span><span class="sxs-lookup"><span data-stu-id="433a1-149">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="433a1-150">Ajouter des hiérarchies de données au tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="433a1-150">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="433a1-151">Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes.</span><span class="sxs-lookup"><span data-stu-id="433a1-151">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="433a1-152">L'ajout des hiérarchies de données des **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** fournit des sommes de ces chiffres pour chaque ligne et colonne.</span><span class="sxs-lookup"><span data-stu-id="433a1-152">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="433a1-153">Dans l'exemple, la **batterie de serveurs** et le **type** sont des lignes, avec le caisse ventes comme données.</span><span class="sxs-lookup"><span data-stu-id="433a1-153">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="433a1-155">Modifier la fonction d'agrégation</span><span class="sxs-lookup"><span data-stu-id="433a1-155">Change aggregation function</span></span>

<span data-ttu-id="433a1-156">Les hiérarchies de données ont leurs valeurs agrégées.</span><span class="sxs-lookup"><span data-stu-id="433a1-156">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="433a1-157">Pour les jeux de données de nombres, il s'agit d'une somme par défaut.</span><span class="sxs-lookup"><span data-stu-id="433a1-157">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="433a1-158">La `summarizeBy` propriété définit ce comportement en fonction d'un type [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .</span><span class="sxs-lookup"><span data-stu-id="433a1-158">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="433a1-159">Les types de fonction d'agrégation actuellement `Sum`pris `Count`en `Average`charge `Max`sont `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`,, `Automatic` ,, et (valeur par défaut).</span><span class="sxs-lookup"><span data-stu-id="433a1-159">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="433a1-160">Les exemples de code suivants modifient l'agrégation pour qu'elle soit la moyenne des données.</span><span class="sxs-lookup"><span data-stu-id="433a1-160">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="433a1-161">Modifier les calculs avec une ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="433a1-161">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="433a1-162">Les tableaux croisés dynamiques agrègent par défaut les données de leurs hiérarchies de ligne et de colonne indépendamment.</span><span class="sxs-lookup"><span data-stu-id="433a1-162">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="433a1-163">Un [ShowAsRule](/javascript/api/excel/excel.showasrule) modifie la hiérarchie des données en valeurs de sortie en fonction d'autres éléments du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="433a1-163">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="433a1-164">L' `ShowAsRule` objet possède trois propriétés:</span><span class="sxs-lookup"><span data-stu-id="433a1-164">The `ShowAsRule` object has three properties:</span></span>

-   <span data-ttu-id="433a1-165">`calculation`: Type de calcul relatif à appliquer à la hiérarchie de données (la valeur par `none`défaut est).</span><span class="sxs-lookup"><span data-stu-id="433a1-165">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
-   <span data-ttu-id="433a1-166">`baseField`: Champ au sein de la hiérarchie contenant les données de base avant l'application du calcul.</span><span class="sxs-lookup"><span data-stu-id="433a1-166">`baseField`: The field within the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="433a1-167">Le [champ PivotField](/javascript/api/excel/excel.pivotfield) a généralement le même nom que sa hiérarchie parente.</span><span class="sxs-lookup"><span data-stu-id="433a1-167">The [PivotField](/javascript/api/excel/excel.pivotfield) usually has the same name as its parent hierarchy.</span></span>
-   <span data-ttu-id="433a1-168">`baseItem`: La valeur de [PivotItem](/javascript/api/excel/excel.pivotitem) individuelle comparée aux valeurs des champs de base basés sur le type de calcul.</span><span class="sxs-lookup"><span data-stu-id="433a1-168">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="433a1-169">Tous les calculs ne nécessitent pas ce champ.</span><span class="sxs-lookup"><span data-stu-id="433a1-169">Not all calculations require this field.</span></span>

<span data-ttu-id="433a1-170">L'exemple suivant montre comment définir le calcul sur la **somme des caisses vendues dans** la hiérarchie des données de la batterie de serveurs pour qu'elle soit un pourcentage du total de la colonne.</span><span class="sxs-lookup"><span data-stu-id="433a1-170">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span> <span data-ttu-id="433a1-171">Nous souhaitons toujours que la granularité s'étende au niveau du type de fruit, c'est pourquoi nous allons utiliser la hiérarchie des lignes de **type** et le champ sous-jacent.</span><span class="sxs-lookup"><span data-stu-id="433a1-171">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span> <span data-ttu-id="433a1-172">L'exemple dispose également d'une **batterie de serveurs** comme première hiérarchie de lignes, de sorte que le nombre total d'entrées de batterie de serveurs affiche également le pourcentage de production de chaque batterie.</span><span class="sxs-lookup"><span data-stu-id="433a1-172">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="433a1-174">L'exemple précédent définit le calcul de la colonne, par rapport à une hiérarchie de lignes individuelle.</span><span class="sxs-lookup"><span data-stu-id="433a1-174">The previous example set the calculation to the column, relative to an individual row hierarchy.</span></span> <span data-ttu-id="433a1-175">Lorsque le calcul est lié à un élément individuel, utilisez `baseItem` la propriété.</span><span class="sxs-lookup"><span data-stu-id="433a1-175">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="433a1-176">L'exemple suivant montre le `differenceFrom` calcul.</span><span class="sxs-lookup"><span data-stu-id="433a1-176">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="433a1-177">Il affiche la différence entre les entrées de hiérarchie des données sur les ventes de la batterie de serveurs par rapport à celles des «batteries de serveurs».</span><span class="sxs-lookup"><span data-stu-id="433a1-177">It displays the difference of the farm crate sales data hierarchy entries relative to those of “A Farms”.</span></span>
<span data-ttu-id="433a1-178">La `baseField` **batterie de serveurs**is, de sorte que nous voyons les différences entre les autres batteries de serveurs, ainsi que les répartitions pour chaque type de fruit similaire (**type** est également une hiérarchie de lignes dans cet exemple).</span><span class="sxs-lookup"><span data-stu-id="433a1-178">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Tableau croisé dynamique montrant les différences entre les ventes de fruit et les autres.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="pivottable-layouts"></a><span data-ttu-id="433a1-182">Dispositions du tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="433a1-182">PivotTable layouts</span></span>

<span data-ttu-id="433a1-183">Un [PivotLayout](/javascript/api/excel/excel.pivotlayout) définit l'emplacement des hiérarchies et de leurs données.</span><span class="sxs-lookup"><span data-stu-id="433a1-183">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="433a1-184">Vous accédez à la disposition pour déterminer les plages dans lesquelles les données sont stockées.</span><span class="sxs-lookup"><span data-stu-id="433a1-184">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="433a1-185">Le diagramme suivant montre les appels de fonction de disposition qui correspondent aux plages du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="433a1-185">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Diagramme montrant les sections d'un tableau croisé dynamique renvoyées par les fonctions Get Range de la disposition.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="433a1-187">Le code suivant montre comment obtenir la dernière ligne des données du tableau croisé dynamique en parcourant la disposition.</span><span class="sxs-lookup"><span data-stu-id="433a1-187">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="433a1-188">Ces valeurs sont ensuite additionnées pour un total général.</span><span class="sxs-lookup"><span data-stu-id="433a1-188">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="433a1-189">Les tableaux croisés dynamiques ont trois styles de disposition: compact, plan et tabulaire.</span><span class="sxs-lookup"><span data-stu-id="433a1-189">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="433a1-190">Nous avons vu le style compact dans les exemples précédents.</span><span class="sxs-lookup"><span data-stu-id="433a1-190">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="433a1-191">Les exemples suivants utilisent respectivement les styles de plan et de tableau.</span><span class="sxs-lookup"><span data-stu-id="433a1-191">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="433a1-192">L'exemple de code montre comment effectuer un basculement entre les différentes dispositions.</span><span class="sxs-lookup"><span data-stu-id="433a1-192">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="433a1-193">Mise en page du plan</span><span class="sxs-lookup"><span data-stu-id="433a1-193">Outline layout</span></span>

![Tableau croisé dynamique à l'aide de la mise en forme du plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="433a1-195">Disposition tabulaire</span><span class="sxs-lookup"><span data-stu-id="433a1-195">Tabular layout</span></span>

![Un tableau croisé dynamique à l'aide de la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="433a1-197">Modifier les noms de hiérarchie</span><span class="sxs-lookup"><span data-stu-id="433a1-197">Change hierarchy names</span></span>

<span data-ttu-id="433a1-198">Les champs de hiérarchie sont modifiables.</span><span class="sxs-lookup"><span data-stu-id="433a1-198">Hierarchy fields are editable.</span></span> <span data-ttu-id="433a1-199">Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.</span><span class="sxs-lookup"><span data-stu-id="433a1-199">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="433a1-200">Supprimer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="433a1-200">Delete a PivotTable</span></span>

<span data-ttu-id="433a1-201">Les tableaux croisés dynamiques sont supprimés à l'aide de leur nom.</span><span class="sxs-lookup"><span data-stu-id="433a1-201">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="433a1-202">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="433a1-202">See also</span></span>

- [<span data-ttu-id="433a1-203">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="433a1-203">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="433a1-204">Référence de l'API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="433a1-204">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
