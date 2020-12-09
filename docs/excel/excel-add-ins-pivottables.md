---
title: Utilisation des tableaux croisés dynamiques avec l’API JavaScript pour Excel
description: Utilisez l’API JavaScript pour Excel pour créer des tableaux croisés dynamiques et interagir avec leurs composants.
ms.date: 12/07/2020
localization_priority: Normal
ms.openlocfilehash: 0a1fefa6a855ab9ee1ccd71fd0dc60f282d2944b
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603798"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="50922-103">Utilisation des tableaux croisés dynamiques avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="50922-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="50922-104">Les tableaux croisés dynamiques rationalisent les grands ensembles de données.</span><span class="sxs-lookup"><span data-stu-id="50922-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="50922-105">Ils permettent la manipulation rapide des données groupées.</span><span class="sxs-lookup"><span data-stu-id="50922-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="50922-106">L’API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d’interagir avec leurs composants.</span><span class="sxs-lookup"><span data-stu-id="50922-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="50922-107">Cet article explique comment les tableaux croisés dynamiques sont représentés par l’API JavaScript Office et fournit des exemples de code pour les scénarios clés.</span><span class="sxs-lookup"><span data-stu-id="50922-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="50922-108">Si vous n’êtes pas familiarisé avec la fonctionnalité de tableaux croisés dynamiques, envisagez de les explorer comme un utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="50922-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="50922-109">Reportez-vous à la rubrique [créer un tableau croisé dynamique pour analyser les données de feuille de calcul](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) pour obtenir une introduction à ces outils.</span><span class="sxs-lookup"><span data-stu-id="50922-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="50922-110">Les tableaux croisés dynamiques créés avec OLAP ne sont actuellement pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="50922-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="50922-111">Il n’existe pas non plus de prise en charge de PowerPivot.</span><span class="sxs-lookup"><span data-stu-id="50922-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="50922-112">Modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="50922-112">Object model</span></span>

<span data-ttu-id="50922-113">Le [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) est l’objet central pour les tableaux croisés dynamiques de l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="50922-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="50922-114">`Workbook.pivotTables` et `Worksheet.pivotTables` sont [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) qui contiennent respectivement les [tableaux croisés dynamiques](/javascript/api/excel/excel.pivottable) dans le classeur et la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="50922-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="50922-115">Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) qui comporte plusieurs [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="50922-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="50922-116">Ces [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) peuvent être ajoutées à des collections de hiérarchies spécifiques pour définir le mode de tableau croisé dynamique des données (comme expliqué dans la [section suivante](#hierarchies)).</span><span class="sxs-lookup"><span data-stu-id="50922-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="50922-117">Un [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contient un [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) qui comporte exactement un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="50922-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="50922-118">Si la conception s’étend pour inclure des tableaux croisés dynamiques OLAP, cela peut changer.</span><span class="sxs-lookup"><span data-stu-id="50922-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="50922-119">Un champ [PivotField](/javascript/api/excel/excel.pivotfield) peut être appliqué à un ou plusieurs [PivotFilters](/javascript/api/excel/excel.pivotfilters) , tant que le [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) du champ est affecté à une catégorie de hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="50922-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span> 
- <span data-ttu-id="50922-120">Un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield) contient un [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) avec plusieurs [PivotItems](/javascript/api/excel/excel.pivotitem).</span><span class="sxs-lookup"><span data-stu-id="50922-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="50922-121">Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotLayout](/javascript/api/excel/excel.pivotlayout) qui définit où les [champs PivotFields](/javascript/api/excel/excel.pivotfield) et [PivotItems](/javascript/api/excel/excel.pivotitem) sont affichés dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="50922-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span>

<span data-ttu-id="50922-122">Examinons comment ces relations s’appliquent à certains exemples de données.</span><span class="sxs-lookup"><span data-stu-id="50922-122">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="50922-123">Les données suivantes décrivent les ventes de fruit de différentes batteries de serveurs.</span><span class="sxs-lookup"><span data-stu-id="50922-123">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="50922-124">Il s’agit de l’exemple de cet article.</span><span class="sxs-lookup"><span data-stu-id="50922-124">It will be the example throughout this article.</span></span>

![Collection de ventes de fruit de différents types de batteries de serveurs différentes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="50922-126">Les données de ventes de la batterie de fruits seront utilisées pour créer un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-126">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="50922-127">Chaque colonne, telle que **types**, est `PivotHierarchy` .</span><span class="sxs-lookup"><span data-stu-id="50922-127">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="50922-128">La hiérarchie de **types** contient le champ **types** .</span><span class="sxs-lookup"><span data-stu-id="50922-128">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="50922-129">Le champ **types** contient les éléments **Apple**, **Kiwi**, **citron**, **citron** et **orange**.</span><span class="sxs-lookup"><span data-stu-id="50922-129">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="50922-130">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="50922-130">Hierarchies</span></span>

<span data-ttu-id="50922-131">Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie : [ligne](/javascript/api/excel/excel.rowcolumnpivothierarchy), [colonne](/javascript/api/excel/excel.rowcolumnpivothierarchy), [données](/javascript/api/excel/excel.datapivothierarchy)et [filtre](/javascript/api/excel/excel.filterpivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="50922-131">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="50922-132">Les données de la batterie de serveurs affichées précédemment ont cinq hiérarchies : **batteries** de serveurs, **type**, **classification**, **caisses vendues à la batterie de serveurs** et **caisses vendues en gros**.</span><span class="sxs-lookup"><span data-stu-id="50922-132">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="50922-133">Chaque hiérarchie peut uniquement exister dans l’une des quatre catégories.</span><span class="sxs-lookup"><span data-stu-id="50922-133">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="50922-134">Si le **type** est ajouté aux hiérarchies de colonne, il ne peut pas également se trouver dans les hiérarchies de ligne, de données ou de filtre.</span><span class="sxs-lookup"><span data-stu-id="50922-134">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="50922-135">Si **type** est par la suite ajouté aux hiérarchies de lignes, il est supprimé des hiérarchies de colonne.</span><span class="sxs-lookup"><span data-stu-id="50922-135">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="50922-136">Ce comportement est le même, que l’attribution de hiérarchie soit réalisée via l’interface utilisateur Excel ou les API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="50922-136">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="50922-137">Les hiérarchies de ligne et de colonne définissent le mode de regroupement des données.</span><span class="sxs-lookup"><span data-stu-id="50922-137">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="50922-138">Par exemple, une hiérarchie de lignes de **batteries de serveurs** regroupe tous les jeux de données de la même batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="50922-138">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="50922-139">Le choix entre la hiérarchie de ligne et de colonne définit l’orientation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-139">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="50922-140">Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de ligne et de colonne.</span><span class="sxs-lookup"><span data-stu-id="50922-140">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="50922-141">Un tableau croisé dynamique avec une hiérarchie de lignes de **batteries de serveurs** et une hiérarchie de données de **grossistes vendus en gros** indique le total de tous les fruits de chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="50922-141">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="50922-142">Les hiérarchies de filtre incluent ou excluent les données du tableau croisé dynamique en fonction des valeurs contenues dans ce type filtré.</span><span class="sxs-lookup"><span data-stu-id="50922-142">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="50922-143">Une hiérarchie de filtrage de **classification** avec le type **Organic** Selected affiche uniquement les données pour les fruits organiques.</span><span class="sxs-lookup"><span data-stu-id="50922-143">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="50922-144">Voici les données de la batterie de serveurs à nouveau, ainsi qu’un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-144">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="50922-145">Le tableau croisé dynamique utilise la **batterie de serveurs** et le **type** comme hiérarchies de lignes, les **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut Sum) et une **classification** en tant que hiérarchie de filtres (avec l’option **Organic** sélectionnée).</span><span class="sxs-lookup"><span data-stu-id="50922-145">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![Sélection de données sur les ventes de fruit en regard d’un tableau croisé dynamique avec des hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="50922-147">Ce tableau croisé dynamique peut être généré via l’API JavaScript ou via l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="50922-147">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="50922-148">Ces deux options permettent une manipulation supplémentaire via les compléments.</span><span class="sxs-lookup"><span data-stu-id="50922-148">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="50922-149">Créer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="50922-149">Create a PivotTable</span></span>

<span data-ttu-id="50922-150">Les tableaux croisés dynamiques nécessitent un nom, une source et une destination.</span><span class="sxs-lookup"><span data-stu-id="50922-150">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="50922-151">La source peut être une adresse de plage ou un nom de table (transmis en tant que `Range` `string` type, ou `Table` type).</span><span class="sxs-lookup"><span data-stu-id="50922-151">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="50922-152">La destination est une adresse de plage (sous la forme a `Range` ou `string` ).</span><span class="sxs-lookup"><span data-stu-id="50922-152">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="50922-153">Les exemples suivants illustrent différentes techniques de création de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-153">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="50922-154">Créer un tableau croisé dynamique avec des adresses de plage</span><span class="sxs-lookup"><span data-stu-id="50922-154">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="50922-155">Création d’un tableau croisé dynamique avec des objets Range</span><span class="sxs-lookup"><span data-stu-id="50922-155">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="50922-156">Création d’un tableau croisé dynamique au niveau du classeur</span><span class="sxs-lookup"><span data-stu-id="50922-156">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="50922-157">Utiliser un tableau croisé dynamique existant</span><span class="sxs-lookup"><span data-stu-id="50922-157">Use an existing PivotTable</span></span>

<span data-ttu-id="50922-158">Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection de tableau croisé dynamique du classeur ou de feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="50922-158">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="50922-159">Le code suivant obtient un tableau croisé dynamique nommé **mon tableau croisé dynamique** à partir du classeur.</span><span class="sxs-lookup"><span data-stu-id="50922-159">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="50922-160">Ajouter des lignes et des colonnes à un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="50922-160">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="50922-161">Lignes et colonnes tableau croisé dynamique des données autour de ces valeurs.</span><span class="sxs-lookup"><span data-stu-id="50922-161">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="50922-162">L’ajout de la colonne **batterie de serveurs** pivote toutes les ventes autour de chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="50922-162">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="50922-163">L’ajout des lignes de type et de **classification** répartit davantage les données en fonction des fruits vendus et s’il s’agit d’un **type** Organic ou non.</span><span class="sxs-lookup"><span data-stu-id="50922-163">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="50922-165">Vous pouvez également utiliser un tableau croisé dynamique avec uniquement des lignes ou des colonnes.</span><span class="sxs-lookup"><span data-stu-id="50922-165">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="50922-166">Ajouter des hiérarchies de données au tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="50922-166">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="50922-167">Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes.</span><span class="sxs-lookup"><span data-stu-id="50922-167">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="50922-168">L’ajout des hiérarchies de données des **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** fournit des sommes de ces chiffres pour chaque ligne et colonne.</span><span class="sxs-lookup"><span data-stu-id="50922-168">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="50922-169">Dans l’exemple, la **batterie de serveurs** et le **type** sont des lignes, avec le caisse ventes comme données.</span><span class="sxs-lookup"><span data-stu-id="50922-169">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="50922-171">Dispositions de tableau croisé dynamique et obtention de données croisées dynamiques</span><span class="sxs-lookup"><span data-stu-id="50922-171">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="50922-172">Un [PivotLayout](/javascript/api/excel/excel.pivotlayout) définit l’emplacement des hiérarchies et de leurs données.</span><span class="sxs-lookup"><span data-stu-id="50922-172">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="50922-173">Vous accédez à la disposition pour déterminer les plages dans lesquelles les données sont stockées.</span><span class="sxs-lookup"><span data-stu-id="50922-173">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="50922-174">Le diagramme suivant montre les appels de fonction de disposition qui correspondent aux plages du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-174">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Diagramme montrant les sections d’un tableau croisé dynamique renvoyées par les fonctions Get Range de la disposition.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="50922-176">Obtenir des données à partir du tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="50922-176">Get data from the PivotTable</span></span>

<span data-ttu-id="50922-177">La disposition définit le mode d’affichage du tableau croisé dynamique dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="50922-177">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="50922-178">Cela signifie que l' `PivotLayout` objet contrôle les plages utilisées pour les éléments de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-178">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="50922-179">Utiliser les plages fournies par la disposition pour obtenir les données collectées et les agréger par le tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-179">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="50922-180">En particulier, utilisez `PivotLayout.getDataBodyRange` pour accéder à ce que génère le tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-180">In particular, use `PivotLayout.getDataBodyRange` to access what the PivotTable produces.</span></span>

<span data-ttu-id="50922-181">Le code suivant montre comment obtenir la dernière ligne des données du tableau croisé dynamique en parcourant la disposition ( **Total général** des colonnes de vente en gros des caisses **vendues au** sein de la batterie de serveurs et **de la somme des colonnes de grossiste vendues** dans l’exemple précédent).</span><span class="sxs-lookup"><span data-stu-id="50922-181">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="50922-182">Ces valeurs sont ensuite additionnées pour un total final, qui s’affiche dans la cellule **E30** (en dehors du tableau croisé dynamique).</span><span class="sxs-lookup"><span data-stu-id="50922-182">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

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

### <a name="layout-types"></a><span data-ttu-id="50922-183">Types de mise en page</span><span class="sxs-lookup"><span data-stu-id="50922-183">Layout types</span></span>

<span data-ttu-id="50922-184">Les tableaux croisés dynamiques ont trois styles de disposition : compact, plan et tabulaire.</span><span class="sxs-lookup"><span data-stu-id="50922-184">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="50922-185">Nous avons vu le style compact dans les exemples précédents.</span><span class="sxs-lookup"><span data-stu-id="50922-185">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="50922-186">Les exemples suivants utilisent respectivement les styles de plan et de tableau.</span><span class="sxs-lookup"><span data-stu-id="50922-186">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="50922-187">L’exemple de code montre comment effectuer un basculement entre les différentes dispositions.</span><span class="sxs-lookup"><span data-stu-id="50922-187">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="50922-188">Mise en page du plan</span><span class="sxs-lookup"><span data-stu-id="50922-188">Outline layout</span></span>

![Tableau croisé dynamique à l’aide de la mise en forme du plan.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="50922-190">Disposition tabulaire</span><span class="sxs-lookup"><span data-stu-id="50922-190">Tabular layout</span></span>

![Un tableau croisé dynamique à l’aide de la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a><span data-ttu-id="50922-192">Supprimer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="50922-192">Delete a PivotTable</span></span>

<span data-ttu-id="50922-193">Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.</span><span class="sxs-lookup"><span data-stu-id="50922-193">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="50922-194">Filtrer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="50922-194">Filter a PivotTable</span></span>

<span data-ttu-id="50922-195">La méthode principale de filtrage des données de tableau croisé dynamique est avec PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="50922-195">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="50922-196">Les segments offrent une méthode de filtrage alternative moins flexible.</span><span class="sxs-lookup"><span data-stu-id="50922-196">Slicers offer an alternate, less flexible filtering method.</span></span> 

<span data-ttu-id="50922-197">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filtre les données en fonction des quatre [catégories de hiérarchie](#hierarchies) d’un tableau croisé dynamique (filtres, colonnes, lignes et valeurs).</span><span class="sxs-lookup"><span data-stu-id="50922-197">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="50922-198">Il existe quatre types de PivotFilters, permettant le filtrage basé sur la date du calendrier, l’analyse des chaînes, la comparaison des nombres et le filtrage en fonction d’une entrée personnalisée.</span><span class="sxs-lookup"><span data-stu-id="50922-198">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span> 

<span data-ttu-id="50922-199">Les [segments](/javascript/api/excel/excel.slicer) peuvent être appliqués aux tableaux croisés dynamiques et aux tableaux Excel réguliers.</span><span class="sxs-lookup"><span data-stu-id="50922-199">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="50922-200">Lorsqu’elle est appliquée à un tableau croisé dynamique, les segments fonctionnent comme un [PivotManualFilter](#pivotmanualfilter) et permettent le filtrage basé sur une entrée personnalisée.</span><span class="sxs-lookup"><span data-stu-id="50922-200">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="50922-201">Contrairement à PivotFilters, les segments ont un [composant d’interface utilisateur Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span><span class="sxs-lookup"><span data-stu-id="50922-201">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="50922-202">Avec la `Slicer` classe, vous créez ce composant d’interface utilisateur, vous gérez le filtrage et vous contrôlez son apparence visuelle.</span><span class="sxs-lookup"><span data-stu-id="50922-202">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span> 

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="50922-203">Filtre avec PivotFilters</span><span class="sxs-lookup"><span data-stu-id="50922-203">Filter with PivotFilters</span></span>

<span data-ttu-id="50922-204">La fonction [PivotFilters](/javascript/api/excel/excel.pivotfilters) vous permet de filtrer les données de tableau croisé dynamique sur la base des quatre [catégories de hiérarchie](#hierarchies) (filtres, colonnes, lignes et valeurs).</span><span class="sxs-lookup"><span data-stu-id="50922-204">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="50922-205">Dans le modèle objet PivotTable, `PivotFilters` sont appliquées à un [champ](/javascript/api/excel/excel.pivotfield)de tableau croisé dynamique et chacune `PivotField` peut avoir une ou plusieurs affectations `PivotFilters` .</span><span class="sxs-lookup"><span data-stu-id="50922-205">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="50922-206">Pour appliquer PivotFilters à un champ de tableau croisé dynamique, les [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) correspondantes du champ doivent être affectées à une catégorie hiérarchique.</span><span class="sxs-lookup"><span data-stu-id="50922-206">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span> 

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="50922-207">Types de PivotFilters</span><span class="sxs-lookup"><span data-stu-id="50922-207">Types of PivotFilters</span></span>

| <span data-ttu-id="50922-208">Type de filtre</span><span class="sxs-lookup"><span data-stu-id="50922-208">Filter type</span></span> | <span data-ttu-id="50922-209">Objectif de filtrage</span><span class="sxs-lookup"><span data-stu-id="50922-209">Filter purpose</span></span> | <span data-ttu-id="50922-210">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="50922-210">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="50922-211">DateFilter</span><span class="sxs-lookup"><span data-stu-id="50922-211">DateFilter</span></span> | <span data-ttu-id="50922-212">Filtrage basé sur la date du calendrier.</span><span class="sxs-lookup"><span data-stu-id="50922-212">Calendar date-based filtering.</span></span> | [<span data-ttu-id="50922-213">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="50922-213">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="50922-214">LabelFilter</span><span class="sxs-lookup"><span data-stu-id="50922-214">LabelFilter</span></span> | <span data-ttu-id="50922-215">Filtrage de comparaison de texte.</span><span class="sxs-lookup"><span data-stu-id="50922-215">Text comparison filtering.</span></span> | [<span data-ttu-id="50922-216">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="50922-216">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="50922-217">ManualFilter</span><span class="sxs-lookup"><span data-stu-id="50922-217">ManualFilter</span></span> | <span data-ttu-id="50922-218">Filtrage de saisie personnalisé.</span><span class="sxs-lookup"><span data-stu-id="50922-218">Custom input filtering.</span></span> | [<span data-ttu-id="50922-219">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="50922-219">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="50922-220">ValueFilter</span><span class="sxs-lookup"><span data-stu-id="50922-220">ValueFilter</span></span> | <span data-ttu-id="50922-221">Filtrage de comparaison de nombres.</span><span class="sxs-lookup"><span data-stu-id="50922-221">Number comparison filtering.</span></span> | [<span data-ttu-id="50922-222">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="50922-222">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="50922-223">Créer un PivotFilter</span><span class="sxs-lookup"><span data-stu-id="50922-223">Create a PivotFilter</span></span>

<span data-ttu-id="50922-224">Pour filtrer les données de tableau croisé dynamique avec un filtre de tableau croisé dynamique \* (tel qu’un PivotDateFilter), appliquez le filtre à un [champ](/javascript/api/excel/excel.pivotfield)de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-224">To filter PivotTable data with a Pivot\*Filter (such as a PivotDateFilter), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="50922-225">Les quatre exemples de code suivants montrent comment utiliser chacun des quatre types de PivotFilters.</span><span class="sxs-lookup"><span data-stu-id="50922-225">The following four code samples show how to use each of the four types of PivotFilters.</span></span> 

##### <a name="pivotdatefilter"></a><span data-ttu-id="50922-226">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="50922-226">PivotDateFilter</span></span>

<span data-ttu-id="50922-227">Le premier exemple de code applique un [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) à la date champ PivotField **mis à jour** , en masquant toutes les données avant le **2020-08-01**.</span><span class="sxs-lookup"><span data-stu-id="50922-227">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span> 

> [!IMPORTANT] 
> <span data-ttu-id="50922-228">Un filtre de tableau croisé dynamique \* ne peut pas être appliqué à un champ PivotField sauf si le PivotHierarchy de ce champ est affecté à une catégorie hiérarchique.</span><span class="sxs-lookup"><span data-stu-id="50922-228">A Pivot\*Filter can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="50922-229">Dans l’exemple de code suivant, le `dateHierarchy` doit être ajouté à la catégorie du tableau croisé dynamique `rowHierarchies` pour pouvoir être utilisé pour le filtrage.</span><span class="sxs-lookup"><span data-stu-id="50922-229">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

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
> <span data-ttu-id="50922-230">Les trois extraits de code suivants affichent uniquement des extraits spécifiques au filtre, au lieu d' `Excel.run` appels complets.</span><span class="sxs-lookup"><span data-stu-id="50922-230">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="50922-231">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="50922-231">PivotLabelFilter</span></span>

<span data-ttu-id="50922-232">Le deuxième extrait de code montre comment appliquer un [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) au **type** PivotField, à l’aide de la `LabelFilterCondition.beginsWith` propriété pour exclure les étiquettes qui commencent par la lettre **L**.</span><span class="sxs-lookup"><span data-stu-id="50922-232">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span> 

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

##### <a name="pivotmanualfilter"></a><span data-ttu-id="50922-233">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="50922-233">PivotManualFilter</span></span>

<span data-ttu-id="50922-234">Le troisième extrait de code applique un filtre manuel avec [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) dans le champ **classification** , en filtrant les données qui n’incluent pas la classification **Organic**.</span><span class="sxs-lookup"><span data-stu-id="50922-234">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span> 

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="50922-235">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="50922-235">PivotValueFilter</span></span>

<span data-ttu-id="50922-236">Pour comparer des nombres, utilisez un filtre de valeur avec [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), comme illustré dans l’extrait de code final.</span><span class="sxs-lookup"><span data-stu-id="50922-236">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="50922-237">Le `PivotValueFilter` compare les données de la **batterie de serveurs** de la batterie aux données du champ PivotField de **grossiste des caisses vendues** , y compris celles dont la somme des caisses vendues dépasse la valeur **500**.</span><span class="sxs-lookup"><span data-stu-id="50922-237">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span> 

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

#### <a name="remove-pivotfilters"></a><span data-ttu-id="50922-238">Suppression de PivotFilters</span><span class="sxs-lookup"><span data-stu-id="50922-238">Remove PivotFilters</span></span>

<span data-ttu-id="50922-239">Pour supprimer tous les PivotFilters, appliquez la `clearAllFilters` méthode à chaque champ de tableau croisé dynamique, comme illustré dans l’exemple de code suivant.</span><span class="sxs-lookup"><span data-stu-id="50922-239">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span> 

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

### <a name="filter-with-slicers"></a><span data-ttu-id="50922-240">Filtre avec des segments</span><span class="sxs-lookup"><span data-stu-id="50922-240">Filter with slicers</span></span>

<span data-ttu-id="50922-241">Les [segments](/javascript/api/excel/excel.slicer) permettent aux données d’être filtrées à partir d’un tableau croisé dynamique ou d’un tableau Excel.</span><span class="sxs-lookup"><span data-stu-id="50922-241">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="50922-242">Un segment utilise des valeurs d’une colonne ou d’un champ PivotField spécifié pour filtrer les lignes correspondantes.</span><span class="sxs-lookup"><span data-stu-id="50922-242">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="50922-243">Ces valeurs sont stockées en tant qu’objets [SlicerItem](/javascript/api/excel/excel.sliceritem) dans le `Slicer` .</span><span class="sxs-lookup"><span data-stu-id="50922-243">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="50922-244">Votre complément peut ajuster ces filtres, comme les utilisateurs peuvent les[utiliser (par le biais de l’interface utilisateur Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span><span class="sxs-lookup"><span data-stu-id="50922-244">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="50922-245">Le segment se trouve au-dessus de la feuille de calcul de la couche de dessin, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="50922-245">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![Données de filtrage de segment sur un tableau croisé dynamique.](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="50922-247">Les techniques décrites dans cette section se concentrent sur l’utilisation des segments connectés aux tableaux croisés dynamiques.</span><span class="sxs-lookup"><span data-stu-id="50922-247">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="50922-248">Les mêmes techniques s’appliquent également à l’utilisation de segments connectés à des tables.</span><span class="sxs-lookup"><span data-stu-id="50922-248">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="50922-249">Créer un segment</span><span class="sxs-lookup"><span data-stu-id="50922-249">Create a slicer</span></span>

<span data-ttu-id="50922-250">Vous pouvez créer un segment dans un classeur ou une feuille de calcul à l’aide de la `Workbook.slicers.add` méthode ou de la `Worksheet.slicers.add` méthode.</span><span class="sxs-lookup"><span data-stu-id="50922-250">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="50922-251">Cette opération ajoute un Slicer au [SlicerCollection](/javascript/api/excel/excel.slicercollection) de l’objet spécifié `Workbook` ou `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="50922-251">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="50922-252">La `SlicerCollection.add` méthode comporte trois paramètres :</span><span class="sxs-lookup"><span data-stu-id="50922-252">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="50922-253">`slicerSource`: La source de données sur laquelle le nouveau segment est basé.</span><span class="sxs-lookup"><span data-stu-id="50922-253">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="50922-254">Il peut s’agir `PivotTable` d’un, `Table` , ou d’une chaîne représentant le nom ou l’ID d’un ou d’un `PivotTable` `Table` .</span><span class="sxs-lookup"><span data-stu-id="50922-254">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="50922-255">`sourceField`: Champ dans la source de données à utiliser pour filtrer.</span><span class="sxs-lookup"><span data-stu-id="50922-255">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="50922-256">Il peut s’agir `PivotField` d’un, `TableColumn` , ou d’une chaîne représentant le nom ou l’ID d’un ou d’un `PivotField` `TableColumn` .</span><span class="sxs-lookup"><span data-stu-id="50922-256">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="50922-257">`slicerDestination`: La feuille de calcul dans laquelle le nouveau segment sera créé.</span><span class="sxs-lookup"><span data-stu-id="50922-257">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="50922-258">Il peut s’agir `Worksheet` d’un objet ou du nom ou de l’ID d’un `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="50922-258">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="50922-259">Ce paramètre n’est pas nécessaire lorsque le `SlicerCollection` est accessible via `Worksheet.slicers` .</span><span class="sxs-lookup"><span data-stu-id="50922-259">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="50922-260">Dans ce cas, la feuille de calcul de la collection est utilisée comme destination.</span><span class="sxs-lookup"><span data-stu-id="50922-260">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="50922-261">L’exemple de code suivant ajoute un nouveau segment à la feuille de calcul de **tableau croisé dynamique** .</span><span class="sxs-lookup"><span data-stu-id="50922-261">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="50922-262">La source du Slicer est le tableau croisé dynamique de la **batterie de serveurs** et les filtres utilisant les données de **type** .</span><span class="sxs-lookup"><span data-stu-id="50922-262">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="50922-263">Le segment est également nommé **segment de fruit** pour référence ultérieure.</span><span class="sxs-lookup"><span data-stu-id="50922-263">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="50922-264">Filtrer des éléments avec un segment</span><span class="sxs-lookup"><span data-stu-id="50922-264">Filter items with a slicer</span></span>

<span data-ttu-id="50922-265">Le segment filtre le tableau croisé dynamique avec les éléments de la `sourceField` .</span><span class="sxs-lookup"><span data-stu-id="50922-265">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="50922-266">La `Slicer.selectItems` méthode définit les éléments qui restent dans le Slicer.</span><span class="sxs-lookup"><span data-stu-id="50922-266">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="50922-267">Ces éléments sont transmis à la méthode en tant que `string[]` , représentant les clés des éléments.</span><span class="sxs-lookup"><span data-stu-id="50922-267">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="50922-268">Toutes les lignes contenant ces éléments restent dans l’agrégation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-268">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="50922-269">Appels suivants permettant de `selectItems` définir la liste aux clés spécifiées dans ces appels.</span><span class="sxs-lookup"><span data-stu-id="50922-269">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="50922-270">Si reçoit `Slicer.selectItems` un élément qui ne se trouve pas dans la source de données, une `InvalidArgument` erreur est générée.</span><span class="sxs-lookup"><span data-stu-id="50922-270">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="50922-271">Le contenu peut être vérifié via la `Slicer.slicerItems` propriété, qui est une [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span><span class="sxs-lookup"><span data-stu-id="50922-271">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="50922-272">L’exemple de code suivant montre trois éléments sélectionnés pour le Slicer : **citron**, **citron** et **orange**.</span><span class="sxs-lookup"><span data-stu-id="50922-272">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="50922-273">Pour supprimer tous les filtres du segment, utilisez la `Slicer.clearFilters` méthode, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="50922-273">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="50922-274">Style et formatage d’un segment</span><span class="sxs-lookup"><span data-stu-id="50922-274">Style and format a slicer</span></span>

<span data-ttu-id="50922-275">Vous pouvez ajuster les paramètres d’affichage d’un segment par le biais de `Slicer` Propriétés.</span><span class="sxs-lookup"><span data-stu-id="50922-275">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="50922-276">L’exemple de code suivant définit le style sur **SlicerStyleLight6**, définit le texte en haut du Slicer sur **types de fruit**, place le segment à la position **(395, 15)** sur la couche de dessin et définit la taille du Slicer sur **135x150** pixels.</span><span class="sxs-lookup"><span data-stu-id="50922-276">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

#### <a name="delete-a-slicer"></a><span data-ttu-id="50922-277">Supprimer un segment</span><span class="sxs-lookup"><span data-stu-id="50922-277">Delete a slicer</span></span>

<span data-ttu-id="50922-278">Pour supprimer un segment, appelez la `Slicer.delete` méthode.</span><span class="sxs-lookup"><span data-stu-id="50922-278">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="50922-279">L’exemple de code suivant supprime le premier segment de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="50922-279">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="50922-280">Modifier la fonction d’agrégation</span><span class="sxs-lookup"><span data-stu-id="50922-280">Change aggregation function</span></span>

<span data-ttu-id="50922-281">Les hiérarchies de données ont leurs valeurs agrégées.</span><span class="sxs-lookup"><span data-stu-id="50922-281">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="50922-282">Pour les jeux de données de nombres, il s’agit d’une somme par défaut.</span><span class="sxs-lookup"><span data-stu-id="50922-282">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="50922-283">La `summarizeBy` propriété définit ce comportement en fonction d’un type [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .</span><span class="sxs-lookup"><span data-stu-id="50922-283">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="50922-284">Les types de fonction d’agrégation actuellement pris en charge sont `Sum` ,, `Count` `Average` , `Max` , `Min` ,,,,,, `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` et `Automatic` (valeur par défaut).</span><span class="sxs-lookup"><span data-stu-id="50922-284">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="50922-285">Les exemples de code suivants modifient l’agrégation pour qu’elle soit la moyenne des données.</span><span class="sxs-lookup"><span data-stu-id="50922-285">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="50922-286">Modifier les calculs avec une ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="50922-286">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="50922-287">Les tableaux croisés dynamiques agrègent par défaut les données de leurs hiérarchies de ligne et de colonne indépendamment.</span><span class="sxs-lookup"><span data-stu-id="50922-287">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="50922-288">Un [ShowAsRule](/javascript/api/excel/excel.showasrule) modifie la hiérarchie des données en valeurs de sortie en fonction d’autres éléments du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="50922-288">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="50922-289">L' `ShowAsRule` objet possède trois propriétés :</span><span class="sxs-lookup"><span data-stu-id="50922-289">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="50922-290">`calculation`: Type de calcul relatif à appliquer à la hiérarchie de données (la valeur par défaut est `none` ).</span><span class="sxs-lookup"><span data-stu-id="50922-290">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="50922-291">`baseField`: [Champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield) dans la hiérarchie contenant les données de base avant l’application du calcul.</span><span class="sxs-lookup"><span data-stu-id="50922-291">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="50922-292">Étant donné que les tableaux croisés dynamiques Excel ont un mappage un-à-un de la hiérarchie sur champ, vous utiliserez le même nom pour accéder à la hiérarchie et au champ.</span><span class="sxs-lookup"><span data-stu-id="50922-292">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="50922-293">`baseItem`: La valeur de [PivotItem](/javascript/api/excel/excel.pivotitem) individuelle comparée aux valeurs des champs de base basés sur le type de calcul.</span><span class="sxs-lookup"><span data-stu-id="50922-293">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="50922-294">Tous les calculs ne nécessitent pas ce champ.</span><span class="sxs-lookup"><span data-stu-id="50922-294">Not all calculations require this field.</span></span>

<span data-ttu-id="50922-295">L’exemple suivant montre comment définir le calcul sur la **somme des caisses vendues dans** la hiérarchie des données de la batterie de serveurs pour qu’elle soit un pourcentage du total de la colonne.</span><span class="sxs-lookup"><span data-stu-id="50922-295">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="50922-296">Nous souhaitons toujours que la granularité s’étende au niveau du type de fruit, c’est pourquoi nous allons utiliser la hiérarchie des lignes de **type** et le champ sous-jacent.</span><span class="sxs-lookup"><span data-stu-id="50922-296">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="50922-297">L’exemple dispose également d’une **batterie de serveurs** comme première hiérarchie de lignes, de sorte que le nombre total d’entrées de batterie de serveurs affiche également le pourcentage de production de chaque batterie.</span><span class="sxs-lookup"><span data-stu-id="50922-297">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="50922-299">L’exemple précédent définit le calcul sur la colonne, par rapport au champ d’une hiérarchie de lignes individuelle.</span><span class="sxs-lookup"><span data-stu-id="50922-299">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="50922-300">Lorsque le calcul est lié à un élément individuel, utilisez la `baseItem` propriété.</span><span class="sxs-lookup"><span data-stu-id="50922-300">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="50922-301">L’exemple suivant montre le `differenceFrom` calcul.</span><span class="sxs-lookup"><span data-stu-id="50922-301">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="50922-302">Il affiche la différence entre les entrées de hiérarchie de données ventes de la batterie de serveurs par rapport à celles d' **une** batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="50922-302">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="50922-303">La `baseField` **batterie de serveurs** is, de sorte que nous voyons les différences entre les autres batteries de serveurs, ainsi que les répartitions pour chaque type de fruit similaire (**type** est également une hiérarchie de lignes dans cet exemple).</span><span class="sxs-lookup"><span data-stu-id="50922-303">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Tableau croisé dynamique montrant les différences entre les ventes de fruit et les autres.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="50922-307">Modifier les noms de hiérarchie</span><span class="sxs-lookup"><span data-stu-id="50922-307">Change hierarchy names</span></span>

<span data-ttu-id="50922-308">Les champs de hiérarchie sont modifiables.</span><span class="sxs-lookup"><span data-stu-id="50922-308">Hierarchy fields are editable.</span></span> <span data-ttu-id="50922-309">Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.</span><span class="sxs-lookup"><span data-stu-id="50922-309">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="50922-310">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="50922-310">See also</span></span>

- [<span data-ttu-id="50922-311">Modèle objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="50922-311">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="50922-312">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="50922-312">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
