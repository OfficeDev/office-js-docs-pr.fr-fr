---
title: Utilisation des tableaux croisés dynamiques avec l’API JavaScript pour Excel
description: Utilisez l’API JavaScript pour Excel pour créer des tableaux croisés dynamiques et interagir avec leurs composants.
ms.date: 01/22/2020
localization_priority: Normal
ms.openlocfilehash: ec7d7ccd7f040185e31b59693827c31d5dab8372
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325093"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="b9b82-103">Utilisation des tableaux croisés dynamiques avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b9b82-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="b9b82-104">Les tableaux croisés dynamiques rationalisent les grands ensembles de données.</span><span class="sxs-lookup"><span data-stu-id="b9b82-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="b9b82-105">Ils permettent la manipulation rapide des données groupées.</span><span class="sxs-lookup"><span data-stu-id="b9b82-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="b9b82-106">L’API JavaScript pour Excel permet à votre complément de créer des tableaux croisés dynamiques et d’interagir avec leurs composants.</span><span class="sxs-lookup"><span data-stu-id="b9b82-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="b9b82-107">Cet article explique comment les tableaux croisés dynamiques sont représentés par l’API JavaScript Office et fournit des exemples de code pour les scénarios clés.</span><span class="sxs-lookup"><span data-stu-id="b9b82-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="b9b82-108">Si vous n’êtes pas familiarisé avec la fonctionnalité de tableaux croisés dynamiques, envisagez de les explorer comme un utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="b9b82-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="b9b82-109">Reportez-vous à la rubrique [créer un tableau croisé dynamique pour analyser les données de feuille de calcul](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) pour obtenir une introduction à ces outils.</span><span class="sxs-lookup"><span data-stu-id="b9b82-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9b82-110">Les tableaux croisés dynamiques créés avec OLAP ne sont actuellement pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="b9b82-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="b9b82-111">Il n’existe pas non plus de prise en charge de PowerPivot.</span><span class="sxs-lookup"><span data-stu-id="b9b82-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="b9b82-112">Modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="b9b82-112">Object model</span></span>

<span data-ttu-id="b9b82-113">Le [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) est l’objet central pour les tableaux croisés dynamiques de l’API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="b9b82-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="b9b82-114">`Workbook.pivotTables`et `Worksheet.pivotTables` sont [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) qui contiennent respectivement les [tableaux croisés dynamiques](/javascript/api/excel/excel.pivottable) dans le classeur et la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b9b82-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="b9b82-115">Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) qui comporte plusieurs [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="b9b82-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="b9b82-116">Un [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contient un [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) qui comporte exactement un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="b9b82-116">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="b9b82-117">Si la conception s’étend pour inclure des tableaux croisés dynamiques OLAP, cela peut changer.</span><span class="sxs-lookup"><span data-stu-id="b9b82-117">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="b9b82-118">Un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield) contient un [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) avec plusieurs [PivotItems](/javascript/api/excel/excel.pivotitem).</span><span class="sxs-lookup"><span data-stu-id="b9b82-118">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="b9b82-119">Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotLayout](/javascript/api/excel/excel.pivotlayout) qui définit où les [champs PivotFields](/javascript/api/excel/excel.pivotfield) et [PivotItems](/javascript/api/excel/excel.pivotitem) sont affichés dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b9b82-119">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span>

<span data-ttu-id="b9b82-120">Examinons comment ces relations s’appliquent à certains exemples de données.</span><span class="sxs-lookup"><span data-stu-id="b9b82-120">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="b9b82-121">Les données suivantes décrivent les ventes de fruit de différentes batteries de serveurs.</span><span class="sxs-lookup"><span data-stu-id="b9b82-121">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="b9b82-122">Il s’agit de l’exemple de cet article.</span><span class="sxs-lookup"><span data-stu-id="b9b82-122">It will be the example throughout this article.</span></span>

![Collection de ventes de fruit de différents types de batteries de serveurs différentes.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="b9b82-124">Les données de ventes de la batterie de fruits seront utilisées pour créer un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="b9b82-124">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="b9b82-125">Chaque colonne, telle que **types**, est `PivotHierarchy`.</span><span class="sxs-lookup"><span data-stu-id="b9b82-125">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="b9b82-126">La hiérarchie de **types** contient le champ **types** .</span><span class="sxs-lookup"><span data-stu-id="b9b82-126">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="b9b82-127">Le champ **types** contient les éléments **Apple**, **Kiwi**, **citron**, **citron**et **orange**.</span><span class="sxs-lookup"><span data-stu-id="b9b82-127">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="b9b82-128">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="b9b82-128">Hierarchies</span></span>

<span data-ttu-id="b9b82-129">Les tableaux croisés dynamiques sont organisés en fonction de quatre catégories de hiérarchie : [ligne](/javascript/api/excel/excel.rowcolumnpivothierarchy), [colonne](/javascript/api/excel/excel.rowcolumnpivothierarchy), [données](/javascript/api/excel/excel.datapivothierarchy)et [filtre](/javascript/api/excel/excel.filterpivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="b9b82-129">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="b9b82-130">Les données de la batterie de serveurs affichées précédemment ont cinq hiérarchies : **batteries**de serveurs, **type**, **classification**, **caisses vendues à la batterie de serveurs**et **caisses vendues en gros**.</span><span class="sxs-lookup"><span data-stu-id="b9b82-130">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="b9b82-131">Chaque hiérarchie peut uniquement exister dans l’une des quatre catégories.</span><span class="sxs-lookup"><span data-stu-id="b9b82-131">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="b9b82-132">Si le **type** est ajouté aux hiérarchies de colonne, il ne peut pas également se trouver dans les hiérarchies de ligne, de données ou de filtre.</span><span class="sxs-lookup"><span data-stu-id="b9b82-132">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="b9b82-133">Si **type** est par la suite ajouté aux hiérarchies de lignes, il est supprimé des hiérarchies de colonne.</span><span class="sxs-lookup"><span data-stu-id="b9b82-133">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="b9b82-134">Ce comportement est le même, que l’attribution de hiérarchie soit réalisée via l’interface utilisateur Excel ou les API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="b9b82-134">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="b9b82-135">Les hiérarchies de ligne et de colonne définissent le mode de regroupement des données.</span><span class="sxs-lookup"><span data-stu-id="b9b82-135">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="b9b82-136">Par exemple, une hiérarchie de lignes de **batteries de serveurs** regroupe tous les jeux de données de la même batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="b9b82-136">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="b9b82-137">Le choix entre la hiérarchie de ligne et de colonne définit l’orientation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="b9b82-137">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="b9b82-138">Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de ligne et de colonne.</span><span class="sxs-lookup"><span data-stu-id="b9b82-138">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="b9b82-139">Un tableau croisé dynamique avec une hiérarchie de lignes de **batteries de serveurs** et une hiérarchie de données de **grossistes vendus en gros** indique le total de tous les fruits de chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="b9b82-139">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="b9b82-140">Les hiérarchies de filtre incluent ou excluent les données du tableau croisé dynamique en fonction des valeurs contenues dans ce type filtré.</span><span class="sxs-lookup"><span data-stu-id="b9b82-140">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="b9b82-141">Une hiérarchie de filtrage de **classification** avec le type **Organic** Selected affiche uniquement les données pour les fruits organiques.</span><span class="sxs-lookup"><span data-stu-id="b9b82-141">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="b9b82-142">Voici les données de la batterie de serveurs à nouveau, ainsi qu’un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="b9b82-142">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="b9b82-143">Le tableau croisé dynamique utilise la **batterie de serveurs** et le **type** comme hiérarchies de lignes, les **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut Sum) et une **classification** en tant que hiérarchie de filtres (avec l’option **Organic** sélectionnée).</span><span class="sxs-lookup"><span data-stu-id="b9b82-143">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![Sélection de données sur les ventes de fruit en regard d’un tableau croisé dynamique avec des hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="b9b82-145">Ce tableau croisé dynamique peut être généré via l’API JavaScript ou via l’interface utilisateur d’Excel.</span><span class="sxs-lookup"><span data-stu-id="b9b82-145">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="b9b82-146">Ces deux options permettent une manipulation supplémentaire via les compléments.</span><span class="sxs-lookup"><span data-stu-id="b9b82-146">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="b9b82-147">Créer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="b9b82-147">Create a PivotTable</span></span>

<span data-ttu-id="b9b82-148">Les tableaux croisés dynamiques nécessitent un nom, une source et une destination.</span><span class="sxs-lookup"><span data-stu-id="b9b82-148">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="b9b82-149">La source peut être une adresse de plage ou un nom de table ( `Range`transmis `string`en tant `Table` que type, ou type).</span><span class="sxs-lookup"><span data-stu-id="b9b82-149">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="b9b82-150">La destination est une adresse de plage (sous la forme `Range` a `string`ou).</span><span class="sxs-lookup"><span data-stu-id="b9b82-150">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="b9b82-151">Les exemples suivants illustrent différentes techniques de création de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="b9b82-151">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="b9b82-152">Créer un tableau croisé dynamique avec des adresses de plage</span><span class="sxs-lookup"><span data-stu-id="b9b82-152">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="b9b82-153">Création d’un tableau croisé dynamique avec des objets Range</span><span class="sxs-lookup"><span data-stu-id="b9b82-153">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="b9b82-154">Création d’un tableau croisé dynamique au niveau du classeur</span><span class="sxs-lookup"><span data-stu-id="b9b82-154">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="b9b82-155">Utiliser un tableau croisé dynamique existant</span><span class="sxs-lookup"><span data-stu-id="b9b82-155">Use an existing PivotTable</span></span>

<span data-ttu-id="b9b82-156">Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection de tableau croisé dynamique du classeur ou de feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="b9b82-156">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="b9b82-157">Le code suivant obtient un tableau croisé dynamique nommé **mon tableau croisé dynamique** à partir du classeur.</span><span class="sxs-lookup"><span data-stu-id="b9b82-157">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="b9b82-158">Ajouter des lignes et des colonnes à un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="b9b82-158">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="b9b82-159">Lignes et colonnes tableau croisé dynamique des données autour de ces valeurs.</span><span class="sxs-lookup"><span data-stu-id="b9b82-159">Rows and columns pivot the data around those fields’ values.</span></span>

<span data-ttu-id="b9b82-160">L’ajout de la colonne **batterie de serveurs** pivote toutes les ventes autour de chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="b9b82-160">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="b9b82-161">L’ajout des lignes de type et de **classification** répartit davantage les données en fonction des fruits vendus et s’il s’agit d’un **type** Organic ou non.</span><span class="sxs-lookup"><span data-stu-id="b9b82-161">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="b9b82-163">Vous pouvez également utiliser un tableau croisé dynamique avec uniquement des lignes ou des colonnes.</span><span class="sxs-lookup"><span data-stu-id="b9b82-163">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="b9b82-164">Ajouter des hiérarchies de données au tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="b9b82-164">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="b9b82-165">Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes.</span><span class="sxs-lookup"><span data-stu-id="b9b82-165">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="b9b82-166">L’ajout des hiérarchies de données des **caisses vendues au niveau** de la batterie de serveurs et des **caisses vendus en gros** fournit des sommes de ces chiffres pour chaque ligne et colonne.</span><span class="sxs-lookup"><span data-stu-id="b9b82-166">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="b9b82-167">Dans l’exemple, la **batterie de serveurs** et le **type** sont des lignes, avec le caisse ventes comme données.</span><span class="sxs-lookup"><span data-stu-id="b9b82-167">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

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

## <a name="slicers"></a><span data-ttu-id="b9b82-169">Slicers</span><span class="sxs-lookup"><span data-stu-id="b9b82-169">Slicers</span></span>

<span data-ttu-id="b9b82-170">Les [segments](/javascript/api/excel/excel.slicer) permettent aux données d’être filtrées à partir d’un tableau croisé dynamique ou d’un tableau Excel.</span><span class="sxs-lookup"><span data-stu-id="b9b82-170">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="b9b82-171">Un segment utilise des valeurs d’une colonne ou d’un champ PivotField spécifié pour filtrer les lignes correspondantes.</span><span class="sxs-lookup"><span data-stu-id="b9b82-171">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="b9b82-172">Ces valeurs sont stockées en [](/javascript/api/excel/excel.sliceritem) tant qu’objets SlicerItem `Slicer`dans le.</span><span class="sxs-lookup"><span data-stu-id="b9b82-172">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="b9b82-173">Votre complément peut ajuster ces filtres, comme les utilisateurs peuvent les[utiliser (par le biais de l’interface utilisateur Excel](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span><span class="sxs-lookup"><span data-stu-id="b9b82-173">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="b9b82-174">Le segment se trouve au-dessus de la feuille de calcul de la couche de dessin, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="b9b82-174">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![Données de filtrage de segment sur un tableau croisé dynamique.](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="b9b82-176">Les techniques décrites dans cette section se concentrent sur l’utilisation des segments connectés aux tableaux croisés dynamiques.</span><span class="sxs-lookup"><span data-stu-id="b9b82-176">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="b9b82-177">Les mêmes techniques s’appliquent également à l’utilisation de segments connectés à des tables.</span><span class="sxs-lookup"><span data-stu-id="b9b82-177">The same techniques also apply to using slicers connected to tables.</span></span>

### <a name="create-a-slicer"></a><span data-ttu-id="b9b82-178">Créer un segment</span><span class="sxs-lookup"><span data-stu-id="b9b82-178">Create a slicer</span></span>

<span data-ttu-id="b9b82-179">Vous pouvez créer un segment dans un classeur ou une feuille de calcul `Workbook.slicers.add` à l' `Worksheet.slicers.add` aide de la méthode ou de la méthode.</span><span class="sxs-lookup"><span data-stu-id="b9b82-179">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="b9b82-180">Cette opération ajoute un Slicer au [SlicerCollection](/javascript/api/excel/excel.slicercollection) de l’objet spécifié `Workbook` ou `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="b9b82-180">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="b9b82-181">La `SlicerCollection.add` méthode comporte trois paramètres :</span><span class="sxs-lookup"><span data-stu-id="b9b82-181">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="b9b82-182">`slicerSource`: La source de données sur laquelle le nouveau segment est basé.</span><span class="sxs-lookup"><span data-stu-id="b9b82-182">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="b9b82-183">Il peut s’agir `PivotTable`d' `Table`un,, ou d’une chaîne représentant le nom `PivotTable` ou `Table`l’ID d’un ou d’un.</span><span class="sxs-lookup"><span data-stu-id="b9b82-183">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="b9b82-184">`sourceField`: Champ dans la source de données à utiliser pour filtrer.</span><span class="sxs-lookup"><span data-stu-id="b9b82-184">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="b9b82-185">Il peut s’agir `PivotField`d' `TableColumn`un,, ou d’une chaîne représentant le nom `PivotField` ou `TableColumn`l’ID d’un ou d’un.</span><span class="sxs-lookup"><span data-stu-id="b9b82-185">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="b9b82-186">`slicerDestination`: La feuille de calcul dans laquelle le nouveau segment sera créé.</span><span class="sxs-lookup"><span data-stu-id="b9b82-186">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="b9b82-187">Il peut s’agir `Worksheet` d’un objet ou du nom ou de `Worksheet`l’ID d’un.</span><span class="sxs-lookup"><span data-stu-id="b9b82-187">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="b9b82-188">Ce paramètre n’est pas nécessaire `SlicerCollection` lorsque le est `Worksheet.slicers`accessible via.</span><span class="sxs-lookup"><span data-stu-id="b9b82-188">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="b9b82-189">Dans ce cas, la feuille de calcul de la collection est utilisée comme destination.</span><span class="sxs-lookup"><span data-stu-id="b9b82-189">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="b9b82-190">L’exemple de code suivant ajoute un nouveau segment à la feuille de calcul de **tableau croisé dynamique** .</span><span class="sxs-lookup"><span data-stu-id="b9b82-190">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="b9b82-191">La source du Slicer est le tableau croisé dynamique de la **batterie de serveurs** et les filtres utilisant les données de **type** .</span><span class="sxs-lookup"><span data-stu-id="b9b82-191">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="b9b82-192">Le segment est également nommé **segment de fruit** pour référence ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b9b82-192">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="b9b82-193">Filtrer des éléments avec un segment</span><span class="sxs-lookup"><span data-stu-id="b9b82-193">Filter items with a slicer</span></span>

<span data-ttu-id="b9b82-194">Le segment filtre le tableau croisé dynamique avec les éléments `sourceField`de la.</span><span class="sxs-lookup"><span data-stu-id="b9b82-194">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="b9b82-195">La `Slicer.selectItems` méthode définit les éléments qui restent dans le Slicer.</span><span class="sxs-lookup"><span data-stu-id="b9b82-195">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="b9b82-196">Ces éléments sont transmis à la méthode en tant `string[]`que, représentant les clés des éléments.</span><span class="sxs-lookup"><span data-stu-id="b9b82-196">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="b9b82-197">Toutes les lignes contenant ces éléments restent dans l’agrégation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="b9b82-197">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="b9b82-198">Appels suivants permettant `selectItems` de définir la liste aux clés spécifiées dans ces appels.</span><span class="sxs-lookup"><span data-stu-id="b9b82-198">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="b9b82-199">Si `Slicer.selectItems` reçoit un élément qui ne se trouve pas dans la source de données `InvalidArgument` , une erreur est générée.</span><span class="sxs-lookup"><span data-stu-id="b9b82-199">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="b9b82-200">Le contenu peut être vérifié via la `Slicer.slicerItems` propriété, qui est une [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span><span class="sxs-lookup"><span data-stu-id="b9b82-200">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="b9b82-201">L’exemple de code suivant montre trois éléments sélectionnés pour le Slicer : **citron**, **citron**et **orange**.</span><span class="sxs-lookup"><span data-stu-id="b9b82-201">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="b9b82-202">Pour supprimer tous les filtres du segment, utilisez la `Slicer.clearFilters` méthode, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="b9b82-202">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

### <a name="style-and-format-a-slicer"></a><span data-ttu-id="b9b82-203">Style et formatage d’un segment</span><span class="sxs-lookup"><span data-stu-id="b9b82-203">Style and format a slicer</span></span>

<span data-ttu-id="b9b82-204">Vous pouvez ajuster les paramètres d’affichage d’un segment par le biais `Slicer` de propriétés.</span><span class="sxs-lookup"><span data-stu-id="b9b82-204">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="b9b82-205">L’exemple de code suivant définit le style sur **SlicerStyleLight6**, définit le texte en haut du Slicer sur **types de fruit**, place le segment à la position **(395, 15)** sur la couche de dessin et définit la taille du Slicer sur **135x150** pixels.</span><span class="sxs-lookup"><span data-stu-id="b9b82-205">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

### <a name="delete-a-slicer"></a><span data-ttu-id="b9b82-206">Supprimer un segment</span><span class="sxs-lookup"><span data-stu-id="b9b82-206">Delete a slicer</span></span>

<span data-ttu-id="b9b82-207">Pour supprimer un segment, appelez la `Slicer.delete` méthode.</span><span class="sxs-lookup"><span data-stu-id="b9b82-207">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="b9b82-208">L’exemple de code suivant supprime le premier segment de la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="b9b82-208">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="b9b82-209">Modifier la fonction d’agrégation</span><span class="sxs-lookup"><span data-stu-id="b9b82-209">Change aggregation function</span></span>

<span data-ttu-id="b9b82-210">Les hiérarchies de données ont leurs valeurs agrégées.</span><span class="sxs-lookup"><span data-stu-id="b9b82-210">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="b9b82-211">Pour les jeux de données de nombres, il s’agit d’une somme par défaut.</span><span class="sxs-lookup"><span data-stu-id="b9b82-211">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="b9b82-212">La `summarizeBy` propriété définit ce comportement en fonction d’un type [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) .</span><span class="sxs-lookup"><span data-stu-id="b9b82-212">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="b9b82-213">Les types de fonction d’agrégation actuellement `Sum`pris `Count`en `Average`charge `Max`sont `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`,, `Automatic` ,, et (valeur par défaut).</span><span class="sxs-lookup"><span data-stu-id="b9b82-213">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="b9b82-214">Les exemples de code suivants modifient l’agrégation pour qu’elle soit la moyenne des données.</span><span class="sxs-lookup"><span data-stu-id="b9b82-214">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="b9b82-215">Modifier les calculs avec une ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="b9b82-215">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="b9b82-216">Les tableaux croisés dynamiques agrègent par défaut les données de leurs hiérarchies de ligne et de colonne indépendamment.</span><span class="sxs-lookup"><span data-stu-id="b9b82-216">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="b9b82-217">Un [ShowAsRule](/javascript/api/excel/excel.showasrule) modifie la hiérarchie des données en valeurs de sortie en fonction d’autres éléments du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="b9b82-217">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="b9b82-218">L' `ShowAsRule` objet possède trois propriétés :</span><span class="sxs-lookup"><span data-stu-id="b9b82-218">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="b9b82-219">`calculation`: Type de calcul relatif à appliquer à la hiérarchie de données (la valeur par `none`défaut est).</span><span class="sxs-lookup"><span data-stu-id="b9b82-219">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="b9b82-220">`baseField`: [Champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield) dans la hiérarchie contenant les données de base avant l’application du calcul.</span><span class="sxs-lookup"><span data-stu-id="b9b82-220">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="b9b82-221">Étant donné que les tableaux croisés dynamiques Excel ont un mappage un-à-un de la hiérarchie sur champ, vous utiliserez le même nom pour accéder à la hiérarchie et au champ.</span><span class="sxs-lookup"><span data-stu-id="b9b82-221">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="b9b82-222">`baseItem`: La valeur de [PivotItem](/javascript/api/excel/excel.pivotitem) individuelle comparée aux valeurs des champs de base basés sur le type de calcul.</span><span class="sxs-lookup"><span data-stu-id="b9b82-222">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="b9b82-223">Tous les calculs ne nécessitent pas ce champ.</span><span class="sxs-lookup"><span data-stu-id="b9b82-223">Not all calculations require this field.</span></span>

<span data-ttu-id="b9b82-224">L’exemple suivant montre comment définir le calcul sur la **somme des caisses vendues dans** la hiérarchie des données de la batterie de serveurs pour qu’elle soit un pourcentage du total de la colonne.</span><span class="sxs-lookup"><span data-stu-id="b9b82-224">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="b9b82-225">Nous souhaitons toujours que la granularité s’étende au niveau du type de fruit, c’est pourquoi nous allons utiliser la hiérarchie des lignes de **type** et le champ sous-jacent.</span><span class="sxs-lookup"><span data-stu-id="b9b82-225">We still want the granularity to extend to the fruit type level, so we’ll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="b9b82-226">L’exemple dispose également d’une **batterie de serveurs** comme première hiérarchie de lignes, de sorte que le nombre total d’entrées de batterie de serveurs affiche également le pourcentage de production de chaque batterie.</span><span class="sxs-lookup"><span data-stu-id="b9b82-226">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="b9b82-228">L’exemple précédent définit le calcul sur la colonne, par rapport au champ d’une hiérarchie de lignes individuelle.</span><span class="sxs-lookup"><span data-stu-id="b9b82-228">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="b9b82-229">Lorsque le calcul est lié à un élément individuel, utilisez `baseItem` la propriété.</span><span class="sxs-lookup"><span data-stu-id="b9b82-229">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="b9b82-230">L’exemple suivant montre le `differenceFrom` calcul.</span><span class="sxs-lookup"><span data-stu-id="b9b82-230">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="b9b82-231">Il affiche la différence entre les entrées de hiérarchie de données ventes de la batterie de serveurs par rapport à celles d' **une**batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="b9b82-231">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="b9b82-232">La `baseField` **batterie de serveurs**is, de sorte que nous voyons les différences entre les autres batteries de serveurs, ainsi que les répartitions pour chaque type de fruit similaire (**type** est également une hiérarchie de lignes dans cet exemple).</span><span class="sxs-lookup"><span data-stu-id="b9b82-232">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

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

## <a name="pivottable-layouts"></a><span data-ttu-id="b9b82-236">Dispositions du tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="b9b82-236">PivotTable layouts</span></span>

<span data-ttu-id="b9b82-237">Un [PivotLayout](/javascript/api/excel/excel.pivotlayout) définit l’emplacement des hiérarchies et de leurs données.</span><span class="sxs-lookup"><span data-stu-id="b9b82-237">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="b9b82-238">Vous accédez à la disposition pour déterminer les plages dans lesquelles les données sont stockées.</span><span class="sxs-lookup"><span data-stu-id="b9b82-238">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="b9b82-239">Le diagramme suivant montre les appels de fonction de disposition qui correspondent aux plages du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="b9b82-239">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Diagramme montrant les sections d’un tableau croisé dynamique renvoyées par les fonctions Get Range de la disposition.](../images/excel-pivots-layout-breakdown.png)

<span data-ttu-id="b9b82-241">Le code suivant montre comment obtenir la dernière ligne des données du tableau croisé dynamique en parcourant la disposition.</span><span class="sxs-lookup"><span data-stu-id="b9b82-241">The following code demonstrates how to get the last row of the PivotTable data by going through the layout.</span></span> <span data-ttu-id="b9b82-242">Ces valeurs sont ensuite additionnées pour un total général.</span><span class="sxs-lookup"><span data-stu-id="b9b82-242">Those values are then summed together for a grand total.</span></span>

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

<span data-ttu-id="b9b82-243">Les tableaux croisés dynamiques ont trois styles de disposition : compact, plan et tabulaire.</span><span class="sxs-lookup"><span data-stu-id="b9b82-243">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="b9b82-244">Nous avons vu le style compact dans les exemples précédents.</span><span class="sxs-lookup"><span data-stu-id="b9b82-244">We’ve seen the compact style in the previous examples.</span></span>

<span data-ttu-id="b9b82-245">Les exemples suivants utilisent respectivement les styles de plan et de tableau.</span><span class="sxs-lookup"><span data-stu-id="b9b82-245">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="b9b82-246">L’exemple de code montre comment effectuer un basculement entre les différentes dispositions.</span><span class="sxs-lookup"><span data-stu-id="b9b82-246">The code sample shows how to cycle between the different layouts.</span></span>

### <a name="outline-layout"></a><span data-ttu-id="b9b82-247">Mise en page du plan</span><span class="sxs-lookup"><span data-stu-id="b9b82-247">Outline layout</span></span>

![Tableau croisé dynamique à l’aide de la mise en forme du plan.](../images/excel-pivots-outline-layout.png)

### <a name="tabular-layout"></a><span data-ttu-id="b9b82-249">Disposition tabulaire</span><span class="sxs-lookup"><span data-stu-id="b9b82-249">Tabular layout</span></span>

![Un tableau croisé dynamique à l’aide de la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

## <a name="change-hierarchy-names"></a><span data-ttu-id="b9b82-251">Modifier les noms de hiérarchie</span><span class="sxs-lookup"><span data-stu-id="b9b82-251">Change hierarchy names</span></span>

<span data-ttu-id="b9b82-252">Les champs de hiérarchie sont modifiables.</span><span class="sxs-lookup"><span data-stu-id="b9b82-252">Hierarchy fields are editable.</span></span> <span data-ttu-id="b9b82-253">Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.</span><span class="sxs-lookup"><span data-stu-id="b9b82-253">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="b9b82-254">Supprimer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="b9b82-254">Delete a PivotTable</span></span>

<span data-ttu-id="b9b82-255">Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.</span><span class="sxs-lookup"><span data-stu-id="b9b82-255">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="b9b82-256">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b9b82-256">See also</span></span>

- [<span data-ttu-id="b9b82-257">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b9b82-257">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="b9b82-258">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b9b82-258">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
