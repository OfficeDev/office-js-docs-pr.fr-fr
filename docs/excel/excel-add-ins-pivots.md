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
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="c6a04-103">Utilisation de tableaux croisés dynamiques à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="c6a04-103">Work with Events using the Excel JavaScript API</span></span>

<span data-ttu-id="c6a04-104">Les tableaux croisés dynamiques rationalisent les jeux de données plus volumineux.</span><span class="sxs-lookup"><span data-stu-id="c6a04-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="c6a04-105">Ils permettent la manipulation rapide des données groupées.</span><span class="sxs-lookup"><span data-stu-id="c6a04-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="c6a04-106">L'API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d'interagir avec leurs composants.</span><span class="sxs-lookup"><span data-stu-id="c6a04-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> 

<span data-ttu-id="c6a04-107">Si vous ne connaissez pas les fonctionnalités des tableaux croisés dynamiques, envisagez de les découvrir en tant qu’utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="c6a04-107">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end-user.</span></span> <span data-ttu-id="c6a04-108">Consultez [Créer un tableau croisé dynamique pour analyser les données d'une feuille de calcul](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) afin d'obtenir une présentation de ces outils.</span><span class="sxs-lookup"><span data-stu-id="c6a04-108">See [Create a PivotTable to analyze worksheet data](https://support.office.com/en-us/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span> 

<span data-ttu-id="c6a04-109">Cet article fournit des exemples de code pour des scénarios courants.</span><span class="sxs-lookup"><span data-stu-id="c6a04-109">This article provides code samples for common scenarios.</span></span> <span data-ttu-id="c6a04-110">[Excel OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel) fournit une documentation de référence complète pour cette fonction de préversion.</span><span class="sxs-lookup"><span data-stu-id="c6a04-110">The [Excel OpenSpec](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel) provides full reference documentation for this preview feature.</span></span> 

<span data-ttu-id="c6a04-111">Pour améliorer votre compréhension de l'API Tableau croisé dynamique, consultez [**Tableau croisé dynamique**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottable.md) et [**Collection Tableau croisé dynamique**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottablecollection.md).</span><span class="sxs-lookup"><span data-stu-id="c6a04-111">To further your understanding of the PivotTable API, see [**PivotTable**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottable.md) and [**PivotTableCollection**](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/pivottablecollection.md).</span></span>

> [!NOTE]
> <span data-ttu-id="c6a04-112">Ces exemples utilisent des API uniquement disponibles dans la préversion publique (bêta) actuellement.</span><span class="sxs-lookup"><span data-stu-id="c6a04-112">These samples use APIs currently available only in public preview (beta).</span></span> <span data-ttu-id="c6a04-113">Ces exemples nécessitent l'exécution de préversions.</span><span class="sxs-lookup"><span data-stu-id="c6a04-113">These samples require preview builds to run.</span></span> <span data-ttu-id="c6a04-114">Utilisez la bibliothèque bêta de [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ou participez au [programme Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="c6a04-114">Either use the beta library of the [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) or join the [Office Insider program](https://products.office.com/office-insider).</span></span> <span data-ttu-id="c6a04-115">Les fonctionnalités du tableau croisé dynamique sont actuellement disponibles dans la version 16.0.10801.20004.</span><span class="sxs-lookup"><span data-stu-id="c6a04-115">PivotTable features are currently available in build 16.0.10801.20004.</span></span>

## <a name="hierarchies"></a><span data-ttu-id="c6a04-116">Hiérarchies</span><span class="sxs-lookup"><span data-stu-id="c6a04-116">Hierarchies</span></span>

<span data-ttu-id="c6a04-117">Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie : ligne, colonne, données et filtre.</span><span class="sxs-lookup"><span data-stu-id="c6a04-117">PivotTables are organized based on four hierarchy categories: row, column, data, and filter.</span></span> <span data-ttu-id="c6a04-118">Les données suivantes décrivant des ventes de fruits provenant de différentes fermes seront utilisées dans cet article.</span><span class="sxs-lookup"><span data-stu-id="c6a04-118">The following data describing fruit sales from various farms will be used throughout this article.</span></span>

![Collection de ventes de fruits de différents types provenant de différentes fermes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="c6a04-120">Ces données ont cinq hiérarchies : **Fermes**, **Type**, **Classification**, **Caisses vendues à la ferme** et **Caisses vendues en gros**.</span><span class="sxs-lookup"><span data-stu-id="c6a04-120">This data has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="c6a04-121">Chaque hiérarchie ne peut exister que dans l’une des quatre catégories.</span><span class="sxs-lookup"><span data-stu-id="c6a04-121">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="c6a04-122">Si le **Type** est ajouté aux hiérarchies de colonnes puis aux hiérarchies de lignes, il ne reste que dans ces dernières.</span><span class="sxs-lookup"><span data-stu-id="c6a04-122">If **Type** is added to column hierarchies and then added to row hierarchies, it only remains in the latter.</span></span>

<span data-ttu-id="c6a04-123">Les hiérarchies de lignes et de colonnes définissent la façon dont les données sont regroupées.</span><span class="sxs-lookup"><span data-stu-id="c6a04-123">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="c6a04-124">Par exemple, une hiérarchie de lignes de **Fermes** regroupe tous les jeux de données provenant de la même ferme.</span><span class="sxs-lookup"><span data-stu-id="c6a04-124">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="c6a04-125">Le choix entre la hiérarchie de lignes et de colonnes définit l'orientation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="c6a04-125">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="c6a04-126">Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de lignes et de colonnes.</span><span class="sxs-lookup"><span data-stu-id="c6a04-126">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="c6a04-127">Un tableau croisé dynamique avec une hiérarchie de lignes de **Fermes** et une hiérarchie de données de **Caisses vendues en gros** affiche la somme totale (par défaut) de tous les différents fruits pour chaque ferme.</span><span class="sxs-lookup"><span data-stu-id="c6a04-127">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="c6a04-128">Les hiérarchies de filtres incluent ou excluent les données provenant du pivot en fonction des valeurs dans ce type filtré.</span><span class="sxs-lookup"><span data-stu-id="c6a04-128">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="c6a04-129">Une hiérarchie de filtres de **Classification** avec le type **Biologique** sélectionné n'affiche que les données pour les fruits biologiques.</span><span class="sxs-lookup"><span data-stu-id="c6a04-129">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="c6a04-130">Voici à nouveau les données des fermes, à côté d’un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="c6a04-130">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="c6a04-131">Le tableau croisé dynamique utilise **Ferme** et **Type** en tant que hiérarchies de lignes, **Caisses vendues à la ferme** et **Caisses vendues en gros** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut de la somme) et **Classification** en tant que hiérarchie de filtres (avec **Biologique** sélectionné).</span><span class="sxs-lookup"><span data-stu-id="c6a04-131">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span> 

![Sélection de données de ventes de fruits à côté d'un tableau croisé dynamique avec les hiérarches de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="c6a04-133">Ce tableau croisé dynamique peut être généré via l’API JavaScript ou l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="c6a04-133">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="c6a04-134">Les deux options permettent une manipulation plus poussée via les compléments.</span><span class="sxs-lookup"><span data-stu-id="c6a04-134">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="c6a04-135">Créer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="c6a04-135">Create a PivotTable or PivotChart report</span></span>

<span data-ttu-id="c6a04-136">Les tableaux croisés dynamiques nécessitent un nom, une source et une destination.</span><span class="sxs-lookup"><span data-stu-id="c6a04-136">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="c6a04-137">La source peut être une adresse de plage ou un nom de table (passés comme un type `Range`, `string` ou `Table`).</span><span class="sxs-lookup"><span data-stu-id="c6a04-137">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="c6a04-138">La destination est une adresse de plage (donnée sous forme de `Range` ou `string`).</span><span class="sxs-lookup"><span data-stu-id="c6a04-138">The destination is a range address (given as either a `Range` or `string`).</span></span> <span data-ttu-id="c6a04-139">Les exemples suivants présentent différentes techniques de création de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="c6a04-139">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="c6a04-140">Créer un tableau croisé dynamique avec des adresses de plages</span><span class="sxs-lookup"><span data-stu-id="c6a04-140">Create a PivotTable with range addresses</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" created on the current worksheet at cell A22 with data from the range A1:E21
    context.workbook.worksheets.getActiveWorksheet()
        .pivotTables.add("Farm Sales", "A1:E21", "A22");

    await context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="c6a04-141">Créer un tableau croisé dynamique avec des objets Plage</span><span class="sxs-lookup"><span data-stu-id="c6a04-141">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="c6a04-142">Créer un tableau croisé dynamique au niveau classeur</span><span class="sxs-lookup"><span data-stu-id="c6a04-142">Create a PivotTable at the workbook level</span></span>

```typescript
await Excel.run(async (context) => {
    // creating a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21
    context.workbook.pivotTables.add("Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    await context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="c6a04-143">Utiliser un tableau croisé dynamique existant</span><span class="sxs-lookup"><span data-stu-id="c6a04-143">Use an existing PivotTable</span></span>

<span data-ttu-id="c6a04-144">Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection Tableau croisé dynamique du classeur ou des feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="c6a04-144">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> 

<span data-ttu-id="c6a04-145">Le code suivant récupère le premier tableau croisé dynamique du classeur.</span><span class="sxs-lookup"><span data-stu-id="c6a04-145">The following code gets the first PivotTable in the workbook.</span></span> <span data-ttu-id="c6a04-146">Il donne ensuite un nom à la table pour faciliter les références ultérieures.</span><span class="sxs-lookup"><span data-stu-id="c6a04-146">It then gives the table a name for easy future reference.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    await context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="c6a04-147">Ajouter des lignes et des colonnes à un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="c6a04-147">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="c6a04-148">Les lignes et les colonnes regroupent les données autour des valeurs de ces champs.</span><span class="sxs-lookup"><span data-stu-id="c6a04-148">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="c6a04-149">L'ajout de la colonne **Ferme** regroupe toutes les ventes autour de chaque ferme.</span><span class="sxs-lookup"><span data-stu-id="c6a04-149">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="c6a04-150">L'ajout des lignes **Type** et **Classification** décompose davantage les données en fonction du fruit vendu et de sa classification biologique ou non.</span><span class="sxs-lookup"><span data-stu-id="c6a04-150">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="c6a04-152">Un tableau croisé dynamique peut également ne contenir que des lignes ou que des colonnes.</span><span class="sxs-lookup"><span data-stu-id="c6a04-152">You can also have a PivotTable with only rows or columns.</span></span>

```typescript
await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));
    
    await context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="c6a04-153">Ajouter des hiérarchies de données au tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="c6a04-153">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="c6a04-154">Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes.</span><span class="sxs-lookup"><span data-stu-id="c6a04-154">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="c6a04-155">L'ajout des hiérarchies de données de **Caisses vendues à la ferme** et **Caisses vendues en gros** donne les sommes de ces chiffres pour chaque ligne et chaque colonne.</span><span class="sxs-lookup"><span data-stu-id="c6a04-155">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span> 

<span data-ttu-id="c6a04-156">Dans l’exemple, **Ferme** et **Type** sont des lignes, tandis que les ventes de caisses sont les données.</span><span class="sxs-lookup"><span data-stu-id="c6a04-156">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span> 

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

## <a name="change-aggregation-function"></a><span data-ttu-id="c6a04-158">Modifier la fonction d’agrégation</span><span class="sxs-lookup"><span data-stu-id="c6a04-158">Change aggregation function</span></span>

<span data-ttu-id="c6a04-159">Les hiérarchies de données voient leurs valeurs agrégées.</span><span class="sxs-lookup"><span data-stu-id="c6a04-159">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="c6a04-160">Pour les jeux de données de nombres, il s’agit d’une somme par défaut.</span><span class="sxs-lookup"><span data-stu-id="c6a04-160">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="c6a04-161">Le `summarizeBy` propriété définit ce comportement en fonction d'un `AggregrationFunction` type.</span><span class="sxs-lookup"><span data-stu-id="c6a04-161">The `summarizeBy` property defines this behavior based on an `AggregrationFunction` type.</span></span> 

<span data-ttu-id="c6a04-162">Les types de fonctions d’agrégation actuellement prises en charge sont `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP` et `Automatic` (par défaut).</span><span class="sxs-lookup"><span data-stu-id="c6a04-162">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="c6a04-163">Les exemples de code suivants modifient l'agrégation en moyennes des données.</span><span class="sxs-lookup"><span data-stu-id="c6a04-163">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="pivottable-layouts"></a><span data-ttu-id="c6a04-164">Dispositions des tableaux croisés dynamiques</span><span class="sxs-lookup"><span data-stu-id="c6a04-164">PivotTable layouts</span></span>

<span data-ttu-id="c6a04-165">La disposition d'un tableau croisé dynamique définit le positionnement des hiérarchies et de leurs données.</span><span class="sxs-lookup"><span data-stu-id="c6a04-165">A PivotTable layout defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="c6a04-166">Accéder à la disposition permet de déterminer les plages de stockage des données.</span><span class="sxs-lookup"><span data-stu-id="c6a04-166">You access the layout to determine the ranges where data is stored.</span></span> 

<span data-ttu-id="c6a04-167">Le diagramme suivant présente la correspondance des appels de fonction de disposition avec les plages du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="c6a04-167">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Diagramme présentant les sections d’un tableau croisé dynamique renvoyées par les fonctions de récupération de plage de la disposition.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="c6a04-169">Le code suivant indique comment récupérer la dernière ligne des données de tableau croisé dynamique via la disposition.</span><span class="sxs-lookup"><span data-stu-id="c6a04-169">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="c6a04-170">Ces valeurs sont ensuite additionnées pour obtenir un total général.</span><span class="sxs-lookup"><span data-stu-id="c6a04-170">Those values are then summed together for a grand total.</span></span>


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

<span data-ttu-id="c6a04-171">Les tableaux croisés dynamiques ont trois styles de disposition : Compact, Plan et Tabulaire.</span><span class="sxs-lookup"><span data-stu-id="c6a04-171">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="c6a04-172">Nous avons vu le style compact dans les exemples précédents.</span><span class="sxs-lookup"><span data-stu-id="c6a04-172">We’ve seen the compact style in the previous examples.</span></span> 

<span data-ttu-id="c6a04-173">Les exemples suivants utilisent respectivement le style plan et tabulaire.</span><span class="sxs-lookup"><span data-stu-id="c6a04-173">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="c6a04-174">L’exemple de code montre comment passer d'une disposition à une autre.</span><span class="sxs-lookup"><span data-stu-id="c6a04-174">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="c6a04-175">Disposition Plan</span><span class="sxs-lookup"><span data-stu-id="c6a04-175">Outline layout</span></span>

![Tableau croisé dynamique utilisant la disposition plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="c6a04-177">Disposition Tabulaire</span><span class="sxs-lookup"><span data-stu-id="c6a04-177">Tabular layout</span></span>

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="c6a04-179">Modifier les noms des hiérarchies</span><span class="sxs-lookup"><span data-stu-id="c6a04-179">Change hierarchy names</span></span>

<span data-ttu-id="c6a04-180">Les champs des hiérarchies peuvent être modifiés.</span><span class="sxs-lookup"><span data-stu-id="c6a04-180">Hierarchy fields are editable.</span></span> <span data-ttu-id="c6a04-181">Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.</span><span class="sxs-lookup"><span data-stu-id="c6a04-181">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="c6a04-182">Supprimer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="c6a04-182">Delete a PivotTable</span></span>

<span data-ttu-id="c6a04-183">Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.</span><span class="sxs-lookup"><span data-stu-id="c6a04-183">PivotTables are deleted by using their name.</span></span>

```typescript
await Excel.run(async (context) => {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();

    await context.sync();
});
```

> [!NOTE]
> <span data-ttu-id="c6a04-184">Votre avis sur la conception de nos préversions est le bienvenu.</span><span class="sxs-lookup"><span data-stu-id="c6a04-184">We welcome feedback on our preview designs.</span></span> <span data-ttu-id="c6a04-185">Si vous avez des commentaires, des suggestions ou des problèmes avec la nouvelle API Tableau croisé dynamique, veuillez laisser vos commentaires sur [UserVoice](https://officespdev.uservoice.com/forums/224641-feature-requests-and-feedback?category_id=163563) ou dans le [répertoire OpenSpec GitHub](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec).</span><span class="sxs-lookup"><span data-stu-id="c6a04-185">If you have comments, suggestions, or issues with the new PivotTable API, please leave your comments on [UserVoice](https://officespdev.uservoice.com/forums/224641-feature-requests-and-feedback?category_id=163563) or on the [OpenSpec GitHub repo](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec).</span></span>
