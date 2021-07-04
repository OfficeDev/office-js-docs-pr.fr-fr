---
title: Utiliser des tableaux croisés dynamiques à l’aide Excel API JavaScript
description: Utilisez l Excel API JavaScript pour créer des tableaux croisés dynamiques et interagir avec leurs composants.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 8c8917f57b7546694e12380fc4369847be24ceac
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290739"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="821fc-103">Utiliser des tableaux croisés dynamiques à l’aide Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="821fc-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="821fc-104">Les tableaux croisés dynamiques simplifient les jeux de données plus volumineux.</span><span class="sxs-lookup"><span data-stu-id="821fc-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="821fc-105">Elles permettent la manipulation rapide des données groupées.</span><span class="sxs-lookup"><span data-stu-id="821fc-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="821fc-106">L Excel API JavaScript permet à votre application de créer des tableaux croisés dynamiques et d’interagir avec leurs composants.</span><span class="sxs-lookup"><span data-stu-id="821fc-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="821fc-107">Cet article décrit comment les tableaux croisés dynamiques sont représentés par Office API JavaScript et fournit des exemples de code pour les scénarios clés.</span><span class="sxs-lookup"><span data-stu-id="821fc-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="821fc-108">Si vous ne connaissez pas la fonctionnalité des tableaux croisés dynamiques, envisagez de les explorer en tant qu’utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="821fc-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="821fc-109">Voir [Créer un tableau croisé dynamique pour analyser les](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) données de feuille de calcul afin d’obtenir une bonne base sur ces outils.</span><span class="sxs-lookup"><span data-stu-id="821fc-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="821fc-110">Les tableaux croisés dynamiques créés avec OLAP ne sont actuellement pas pris en charge.</span><span class="sxs-lookup"><span data-stu-id="821fc-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="821fc-111">Il n’existe pas non plus de prise en charge de Power Pivot.</span><span class="sxs-lookup"><span data-stu-id="821fc-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="821fc-112">Modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="821fc-112">Object model</span></span>

<span data-ttu-id="821fc-113">Le [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) est l’objet central des tableaux croisés dynamiques dans l Office API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="821fc-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="821fc-114">`Workbook.pivotTables` et `Worksheet.pivotTables` sont [des PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) qui contiennent respectivement les tableaux [croisés dynamiques](/javascript/api/excel/excel.pivottable) dans le workbook et la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="821fc-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="821fc-115">Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) qui possède plusieurs [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span><span class="sxs-lookup"><span data-stu-id="821fc-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="821fc-116">Ces [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) peuvent être ajoutées à des collections de hiérarchies spécifiques pour définir la façon dont le tableau croisé dynamique analyse les données (comme expliqué dans la [section suivante).](#hierarchies)</span><span class="sxs-lookup"><span data-stu-id="821fc-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="821fc-117">Une [PivotHierarchy contient](/javascript/api/excel/excel.pivothierarchy) un [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) qui possède exactement un [champ de tableau croisé dynamique](/javascript/api/excel/excel.pivotfield).</span><span class="sxs-lookup"><span data-stu-id="821fc-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="821fc-118">Si la conception est étendue pour inclure des tableaux croisés dynamiques OLAP, cela peut changer.</span><span class="sxs-lookup"><span data-stu-id="821fc-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="821fc-119">Un [champ de](/javascript/api/excel/excel.pivotfield) tableau croisé dynamique peut avoir un ou plusieurs filtres de tableau croisé dynamique [appliqués,](/javascript/api/excel/excel.pivotfilters) tant que la [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) du champ est affectée à une catégorie de hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="821fc-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span>
- <span data-ttu-id="821fc-120">Un [champ de](/javascript/api/excel/excel.pivotfield) tableau croisé dynamique contient un [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) qui a plusieurs [pivotItems](/javascript/api/excel/excel.pivotitem).</span><span class="sxs-lookup"><span data-stu-id="821fc-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="821fc-121">Un [tableau croisé dynamique](/javascript/api/excel/excel.pivottable) contient un [pivotLayout](/javascript/api/excel/excel.pivotlayout) qui définit l’endroit où les [pivotFields](/javascript/api/excel/excel.pivotfield) et [pivotItems](/javascript/api/excel/excel.pivotitem) sont affichés dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="821fc-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span> <span data-ttu-id="821fc-122">La disposition contrôle également certains paramètres d’affichage du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-122">The layout also controls some display settings for the PivotTable.</span></span>

<span data-ttu-id="821fc-123">Examinons comment ces relations s’appliquent à certains exemples de données.</span><span class="sxs-lookup"><span data-stu-id="821fc-123">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="821fc-124">Les données suivantes décrivent les ventes de fruit de différentes batteries de serveurs.</span><span class="sxs-lookup"><span data-stu-id="821fc-124">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="821fc-125">Ce sera l’exemple tout au long de cet article.</span><span class="sxs-lookup"><span data-stu-id="821fc-125">It will be the example throughout this article.</span></span>

![Collection de ventes de fruit de différents types de batteries de serveurs.](../images/excel-pivots-raw-data.png)

<span data-ttu-id="821fc-127">Les données de ventes de cette batterie de serveurs de fruit seront utilisées pour la production d’un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-127">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="821fc-128">Chaque colonne, telle que **Types,** est une `PivotHierarchy` .</span><span class="sxs-lookup"><span data-stu-id="821fc-128">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="821fc-129">La **hiérarchie Types** contient le champ **Types.**</span><span class="sxs-lookup"><span data-stu-id="821fc-129">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="821fc-130">Le **champ Types** contient les éléments **Apple**, **Domaine,** **Domaine,** **Vert** vert et **Orange**.</span><span class="sxs-lookup"><span data-stu-id="821fc-130">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="821fc-131">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="821fc-131">Hierarchies</span></span>

<span data-ttu-id="821fc-132">Les tableaux croisés dynamiques sont organisés en quatre catégories hiérarchiques : [ligne,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [colonne,](/javascript/api/excel/excel.rowcolumnpivothierarchy) [données](/javascript/api/excel/excel.datapivothierarchy)et [filtre.](/javascript/api/excel/excel.filterpivothierarchy)</span><span class="sxs-lookup"><span data-stu-id="821fc-132">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="821fc-133">Les données de batterie de serveurs indiquées précédemment disposent de cinq hiérarchies : Batteries **de** serveurs, **Type**, **Classification**, **Caisses vendues** à la batterie de serveurs et **Caisses vendues en commun**.</span><span class="sxs-lookup"><span data-stu-id="821fc-133">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="821fc-134">Chaque hiérarchie ne peut exister que dans l’une des quatre catégories.</span><span class="sxs-lookup"><span data-stu-id="821fc-134">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="821fc-135">Si **type** est ajouté aux hiérarchies de colonnes, il ne peut pas non plus se trouver dans les hiérarchies de lignes, de données ou de filtres.</span><span class="sxs-lookup"><span data-stu-id="821fc-135">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="821fc-136">Si **Type** est ensuite ajouté aux hiérarchies de lignes, il est supprimé des hiérarchies de colonnes.</span><span class="sxs-lookup"><span data-stu-id="821fc-136">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="821fc-137">Ce comportement est le même si l’affectation de hiérarchie est effectuée via l’interface Excel’interface utilisateur ou Excel api JavaScript.</span><span class="sxs-lookup"><span data-stu-id="821fc-137">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="821fc-138">Les hiérarchies de lignes et de colonnes définissent le regroupement des données.</span><span class="sxs-lookup"><span data-stu-id="821fc-138">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="821fc-139">Par exemple, une hiérarchie de lignes **de** batteries de serveurs rassemble tous les ensembles de données de la même batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="821fc-139">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="821fc-140">Le choix entre la hiérarchie de lignes et de colonnes définit l’orientation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-140">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="821fc-141">Les hiérarchies de données sont les valeurs à agréger en fonction des hiérarchies de lignes et de colonnes.</span><span class="sxs-lookup"><span data-stu-id="821fc-141">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="821fc-142">Un tableau croisé dynamique avec  une hiérarchie de lignes de batteries de serveurs et une hiérarchie de données de l’ordre des **caisses vendues indique** le total total (par défaut) de tous les différents produits pour chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="821fc-142">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="821fc-143">Les hiérarchies de filtres incluent ou excluent des données du tableau croisé dynamique en fonction des valeurs de ce type filtré.</span><span class="sxs-lookup"><span data-stu-id="821fc-143">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="821fc-144">Une hiérarchie de filtres de **classification** avec le type **organique** sélectionné affiche uniquement les données pour les fruit organiques.</span><span class="sxs-lookup"><span data-stu-id="821fc-144">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="821fc-145">Voici à nouveau les données de la batterie de serveurs, ainsi qu’un tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-145">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="821fc-146">Le tableau croisé dynamique utilise  Farm **and** **Type** comme hiérarchies de lignes, La vente des **caisses** sur la batterie de serveurs et la vente **de caisses** en tant que hiérarchies de données (avec la fonction d’agrégation par défaut de somme) et **classification** en tant que hiérarchie de filtre (avec l’alimentation organique sélectionnée).</span><span class="sxs-lookup"><span data-stu-id="821fc-146">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![Sélection de données de ventes de fruit à côté d’un tableau croisé dynamique avec des hiérarchies de lignes, de données et de filtres.](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="821fc-148">Ce tableau croisé dynamique peut être généré via l’API JavaScript ou par le biais Excel’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="821fc-148">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="821fc-149">Les deux options permettent d’autres manipulations par le biais de leurs add-ins.</span><span class="sxs-lookup"><span data-stu-id="821fc-149">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="821fc-150">Créer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-150">Create a PivotTable</span></span>

<span data-ttu-id="821fc-151">Les tableaux croisés dynamiques ont besoin d’un nom, d’une source et d’une destination.</span><span class="sxs-lookup"><span data-stu-id="821fc-151">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="821fc-152">La source peut être une adresse de plage ou un nom de table (transmis en tant `Range` `string` que , ou `Table` type).</span><span class="sxs-lookup"><span data-stu-id="821fc-152">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="821fc-153">La destination est une adresse de plage (donnée en tant que a `Range` ou `string` ).</span><span class="sxs-lookup"><span data-stu-id="821fc-153">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="821fc-154">Les exemples suivants montrent différentes techniques de création de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-154">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="821fc-155">Créer un tableau croisé dynamique avec des adresses de plage</span><span class="sxs-lookup"><span data-stu-id="821fc-155">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="821fc-156">Créer un tableau croisé dynamique avec des objets Range</span><span class="sxs-lookup"><span data-stu-id="821fc-156">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="821fc-157">Créer un tableau croisé dynamique au niveau du workbook</span><span class="sxs-lookup"><span data-stu-id="821fc-157">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="821fc-158">Utiliser un tableau croisé dynamique existant</span><span class="sxs-lookup"><span data-stu-id="821fc-158">Use an existing PivotTable</span></span>

<span data-ttu-id="821fc-159">Les tableaux croisés dynamiques créés manuellement sont également accessibles via la collection de tableaux croisés dynamiques du manuel ou des feuilles de calcul individuelles.</span><span class="sxs-lookup"><span data-stu-id="821fc-159">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="821fc-160">Le code suivant obtient un tableau croisé dynamique nommé **My Pivot** à partir du workbook.</span><span class="sxs-lookup"><span data-stu-id="821fc-160">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="821fc-161">Ajouter des lignes et des colonnes à un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-161">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="821fc-162">Les lignes et les colonnes pivotent les données autour des valeurs de ces champs.</span><span class="sxs-lookup"><span data-stu-id="821fc-162">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="821fc-163">L’ajout **de la** colonne Batterie de serveurs pivote toutes les ventes autour de chaque batterie de serveurs.</span><span class="sxs-lookup"><span data-stu-id="821fc-163">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="821fc-164">L’ajout **des lignes Type** et **Classification** décompose davantage les données en fonction des fruit vendus et selon qu’il s’agit d’un produit organique ou non.</span><span class="sxs-lookup"><span data-stu-id="821fc-164">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="821fc-166">Vous pouvez également avoir un tableau croisé dynamique avec uniquement des lignes ou des colonnes.</span><span class="sxs-lookup"><span data-stu-id="821fc-166">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="821fc-167">Ajouter des hiérarchies de données au tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-167">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="821fc-168">Les hiérarchies de données remplissent le tableau croisé dynamique avec des informations à combiner en fonction des lignes et des colonnes.</span><span class="sxs-lookup"><span data-stu-id="821fc-168">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="821fc-169">L’ajout des hiérarchies de données **des caisses vendues** à la batterie de serveurs et des **caisses vendues permet** d’obtenir les sommes de ces chiffres pour chaque ligne et colonne.</span><span class="sxs-lookup"><span data-stu-id="821fc-169">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="821fc-170">Dans l’exemple, **la batterie de** serveurs et le **type** sont des lignes, avec les ventes de caisse en tant que données.</span><span class="sxs-lookup"><span data-stu-id="821fc-170">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="821fc-172">Dispositions de tableau croisé dynamique et obtention de données pivotées</span><span class="sxs-lookup"><span data-stu-id="821fc-172">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="821fc-173">Un [pivotLayout](/javascript/api/excel/excel.pivotlayout) définit le placement des hiérarchies et leurs données.</span><span class="sxs-lookup"><span data-stu-id="821fc-173">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="821fc-174">Vous accédez à la disposition pour déterminer les plages où les données sont stockées.</span><span class="sxs-lookup"><span data-stu-id="821fc-174">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="821fc-175">Le diagramme suivant montre les appels de fonction de disposition qui correspondent aux plages du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-175">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![Diagramme montrant quelles sections d’un tableau croisé dynamique sont renvoyées par les fonctions get range de la disposition.](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="821fc-177">Obtenir des données à partir du tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-177">Get data from the PivotTable</span></span>

<span data-ttu-id="821fc-178">La disposition définit la façon dont le tableau croisé dynamique est affiché dans la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="821fc-178">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="821fc-179">Cela signifie que `PivotLayout` l’objet contrôle les plages utilisées pour les éléments de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-179">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="821fc-180">Utilisez les plages fournies par la disposition pour obtenir des données collectées et agrégées par le tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-180">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="821fc-181">En particulier, utilisez `PivotLayout.getDataBodyRange` cette information pour accéder aux données produites par le tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-181">In particular, use `PivotLayout.getDataBodyRange` to access the data produced by the PivotTable.</span></span>

<span data-ttu-id="821fc-182">Le code suivant montre comment obtenir la dernière ligne des données du tableau croisé dynamique en passant par la disposition (le **total total des** **montants** vendus à la batterie de serveurs et la somme des **caisses vendues dans** l’exemple précédent).</span><span class="sxs-lookup"><span data-stu-id="821fc-182">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="821fc-183">Ces valeurs sont ensuite additionées pour un total final, qui est affiché dans la cellule **E30** (en dehors du tableau croisé dynamique).</span><span class="sxs-lookup"><span data-stu-id="821fc-183">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

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

### <a name="layout-types"></a><span data-ttu-id="821fc-184">Types de disposition</span><span class="sxs-lookup"><span data-stu-id="821fc-184">Layout types</span></span>

<span data-ttu-id="821fc-185">Les tableaux croisés dynamiques ont trois styles de disposition : Compact, Outline et Tabulaire.</span><span class="sxs-lookup"><span data-stu-id="821fc-185">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="821fc-186">Nous avons vu le style compact dans les exemples précédents.</span><span class="sxs-lookup"><span data-stu-id="821fc-186">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="821fc-187">Les exemples suivants utilisent respectivement les styles plan et tabulaire.</span><span class="sxs-lookup"><span data-stu-id="821fc-187">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="821fc-188">L’exemple de code montre comment faire un cycle entre les différentes dispositions.</span><span class="sxs-lookup"><span data-stu-id="821fc-188">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="821fc-189">Disposition du plan</span><span class="sxs-lookup"><span data-stu-id="821fc-189">Outline layout</span></span>

![Tableau croisé dynamique utilisant la disposition du plan.](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="821fc-191">Disposition tabulaire</span><span class="sxs-lookup"><span data-stu-id="821fc-191">Tabular layout</span></span>

![Tableau croisé dynamique utilisant la disposition tabulaire.](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a><span data-ttu-id="821fc-193">Exemple de code de commutateur de type PivotLayout</span><span class="sxs-lookup"><span data-stu-id="821fc-193">PivotLayout type switch code sample</span></span>

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

### <a name="other-pivotlayout-functions"></a><span data-ttu-id="821fc-194">Autres fonctions PivotLayout</span><span class="sxs-lookup"><span data-stu-id="821fc-194">Other PivotLayout functions</span></span>

<span data-ttu-id="821fc-195">Par défaut, les tableaux croisés dynamiques ajustent les tailles de lignes et de colonnes selon les besoins.</span><span class="sxs-lookup"><span data-stu-id="821fc-195">By default, PivotTables adjust row and column sizes as needed.</span></span> <span data-ttu-id="821fc-196">Cette chose est effectuée lorsque le tableau croisé dynamique est actualisé.</span><span class="sxs-lookup"><span data-stu-id="821fc-196">This is done when the PivotTable is refreshed.</span></span> <span data-ttu-id="821fc-197">`PivotLayout.autoFormat` spécifie ce comportement.</span><span class="sxs-lookup"><span data-stu-id="821fc-197">`PivotLayout.autoFormat` specifies that behavior.</span></span> <span data-ttu-id="821fc-198">Les modifications de taille de ligne ou de colonne apportées par votre add-in sont persistantes `autoFormat` lorsqu’elles le `false` sont.</span><span class="sxs-lookup"><span data-stu-id="821fc-198">Any row or column size changes made by your add-in persist when `autoFormat` is `false`.</span></span> <span data-ttu-id="821fc-199">En outre, les paramètres par défaut d’un tableau croisé dynamique conservent toute mise en forme personnalisée dans le tableau croisé dynamique (par exemple, les remplissages et les modifications de police).</span><span class="sxs-lookup"><span data-stu-id="821fc-199">Additionally, the default settings of a PivotTable keep any custom formatting in the PivotTable (such as fills and font changes).</span></span> <span data-ttu-id="821fc-200">Définir `PivotLayout.preserveFormatting` pour appliquer le format par défaut lors de `false` l’actualisation.</span><span class="sxs-lookup"><span data-stu-id="821fc-200">Set `PivotLayout.preserveFormatting` to `false` to apply the default format when refreshed.</span></span>

<span data-ttu-id="821fc-201">A `PivotLayout` contrôle également les paramètres d’en-tête et de ligne totale, la façon dont les cellules de données vides sont affichées et les options de texte [de](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) alt.</span><span class="sxs-lookup"><span data-stu-id="821fc-201">A `PivotLayout` also controls header and total row settings, how empty data cells are displayed, and [alt text](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) options.</span></span> <span data-ttu-id="821fc-202">La [référence PivotLayout](/javascript/api/excel/excel.pivotlayout) fournit une liste complète de ces fonctionnalités.</span><span class="sxs-lookup"><span data-stu-id="821fc-202">The [PivotLayout](/javascript/api/excel/excel.pivotlayout) reference provides a complete list of these features.</span></span>

<span data-ttu-id="821fc-203">L’exemple de code suivant permet aux cellules de données vides d’afficher la chaîne, met en forme la plage de corps avec un alignement horizontal cohérent et garantit que les modifications de mise en forme restent même après l’actualisation du tableau croisé `"--"` dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-203">The following code sample makes empty data cells display the string `"--"`, formats the body range to a consistent horizontal alignment, and ensures that the formatting changes remain even after the PivotTable is refreshed.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="821fc-204">Supprimer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-204">Delete a PivotTable</span></span>

<span data-ttu-id="821fc-205">Les tableaux croisés dynamiques sont supprimés à l’aide de leur nom.</span><span class="sxs-lookup"><span data-stu-id="821fc-205">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="821fc-206">Filtrer un tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-206">Filter a PivotTable</span></span>

<span data-ttu-id="821fc-207">La méthode principale de filtrage des données de tableau croisé dynamique est avec des filtres de tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-207">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="821fc-208">Les slicers offrent une autre méthode de filtrage moins flexible.</span><span class="sxs-lookup"><span data-stu-id="821fc-208">Slicers offer an alternate, less flexible filtering method.</span></span>

<span data-ttu-id="821fc-209">[Les filtres de tableau](/javascript/api/excel/excel.pivotfilters) croisé dynamique filtrent les données en fonction des quatre [catégories hiérarchiques](#hierarchies) d’un tableau croisé dynamique (filtres, colonnes, lignes et valeurs).</span><span class="sxs-lookup"><span data-stu-id="821fc-209">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="821fc-210">Il existe quatre types de filtres de tableau croisé dynamique, ce qui permet le filtrage basé sur les dates du calendrier, l’comparaison des chaînes, la comparaison des nombres et le filtrage en fonction d’une entrée personnalisée.</span><span class="sxs-lookup"><span data-stu-id="821fc-210">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span>

<span data-ttu-id="821fc-211">[Les slicers](/javascript/api/excel/excel.slicer) peuvent être appliqués à la fois aux tableaux croisés dynamiques et aux tableaux Excel tableaux.</span><span class="sxs-lookup"><span data-stu-id="821fc-211">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="821fc-212">Lorsqu’ils sont appliqués à un tableau croisé dynamique, les slicers fonctionnent comme un [pivotManualFilter](#pivotmanualfilter) et autorisent le filtrage basé sur une entrée personnalisée.</span><span class="sxs-lookup"><span data-stu-id="821fc-212">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="821fc-213">Contrairement aux filtres de tableau croisé dynamique, les slicers ont [un Excel’interface utilisateur.](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)</span><span class="sxs-lookup"><span data-stu-id="821fc-213">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="821fc-214">Avec la `Slicer` classe, vous créez ce composant d’interface utilisateur, gérez le filtrage et contrôlez son apparence visuelle.</span><span class="sxs-lookup"><span data-stu-id="821fc-214">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span>

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="821fc-215">Filtrer avec des filtres de tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-215">Filter with PivotFilters</span></span>

<span data-ttu-id="821fc-216">[Les filtres de tableau](/javascript/api/excel/excel.pivotfilters) croisé dynamique vous permettent de filtrer les données de tableau croisé dynamique en fonction des quatre [catégories hiérarchiques (filtres,](#hierarchies) colonnes, lignes et valeurs).</span><span class="sxs-lookup"><span data-stu-id="821fc-216">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="821fc-217">Dans le modèle objet de tableau croisé dynamique, sont appliqués à un champ de tableau croisé dynamique , et chacun peut `PivotFilters` avoir un ou plusieurs [](/javascript/api/excel/excel.pivotfield) `PivotField` `PivotFilters` attribués .</span><span class="sxs-lookup"><span data-stu-id="821fc-217">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="821fc-218">Pour appliquer des filtres de tableau croisé dynamique à un champ de tableau croisé dynamique, la [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) correspondante du champ doit être affectée à une catégorie de hiérarchie.</span><span class="sxs-lookup"><span data-stu-id="821fc-218">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span>

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="821fc-219">Types de filtres de tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-219">Types of PivotFilters</span></span>

| <span data-ttu-id="821fc-220">Type de filtre</span><span class="sxs-lookup"><span data-stu-id="821fc-220">Filter type</span></span> | <span data-ttu-id="821fc-221">Objectif du filtre</span><span class="sxs-lookup"><span data-stu-id="821fc-221">Filter purpose</span></span> | <span data-ttu-id="821fc-222">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="821fc-222">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="821fc-223">DateFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-223">DateFilter</span></span> | <span data-ttu-id="821fc-224">Filtrage basé sur les dates du calendrier.</span><span class="sxs-lookup"><span data-stu-id="821fc-224">Calendar date-based filtering.</span></span> | [<span data-ttu-id="821fc-225">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-225">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="821fc-226">LabelFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-226">LabelFilter</span></span> | <span data-ttu-id="821fc-227">Filtrage de comparaison de texte.</span><span class="sxs-lookup"><span data-stu-id="821fc-227">Text comparison filtering.</span></span> | [<span data-ttu-id="821fc-228">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-228">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="821fc-229">ManualFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-229">ManualFilter</span></span> | <span data-ttu-id="821fc-230">Filtrage des entrées personnalisé.</span><span class="sxs-lookup"><span data-stu-id="821fc-230">Custom input filtering.</span></span> | [<span data-ttu-id="821fc-231">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-231">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="821fc-232">ValueFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-232">ValueFilter</span></span> | <span data-ttu-id="821fc-233">Filtrage de comparaison de nombres.</span><span class="sxs-lookup"><span data-stu-id="821fc-233">Number comparison filtering.</span></span> | [<span data-ttu-id="821fc-234">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-234">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="821fc-235">Créer un filtre de tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-235">Create a PivotFilter</span></span>

<span data-ttu-id="821fc-236">Pour filtrer des données de tableau croisé dynamique avec `Pivot*Filter` un (par `PivotDateFilter` exemple, un ), appliquez le filtre à un champ [de tableau croisé dynamique.](/javascript/api/excel/excel.pivotfield)</span><span class="sxs-lookup"><span data-stu-id="821fc-236">To filter PivotTable data with a `Pivot*Filter` (such as a `PivotDateFilter`), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="821fc-237">Les quatre exemples de code suivants montrent comment utiliser chacun des quatre types de filtres croisés dynamiques.</span><span class="sxs-lookup"><span data-stu-id="821fc-237">The following four code samples show how to use each of the four types of PivotFilters.</span></span>

##### <a name="pivotdatefilter"></a><span data-ttu-id="821fc-238">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-238">PivotDateFilter</span></span>

<span data-ttu-id="821fc-239">Le premier exemple de code applique un [pivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) au champ de tableau croisé dynamique **date-mise** à jour, masquant les données antérieures au **08-2020-08-01**.</span><span class="sxs-lookup"><span data-stu-id="821fc-239">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="821fc-240">A `Pivot*Filter` can’t be applied to a PivotField unless that field’s PivotHierarchy is assigned to a hierarchy category.</span><span class="sxs-lookup"><span data-stu-id="821fc-240">A `Pivot*Filter` can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="821fc-241">Dans l’exemple de code suivant, le tableau croisé dynamique doit être ajouté à la catégorie du tableau croisé dynamique avant de pouvoir être `dateHierarchy` `rowHierarchies` utilisé pour le filtrage.</span><span class="sxs-lookup"><span data-stu-id="821fc-241">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

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
> <span data-ttu-id="821fc-242">Les trois extraits de code suivants affichent uniquement des extraits spécifiques au filtre, au lieu d’appels `Excel.run` complets.</span><span class="sxs-lookup"><span data-stu-id="821fc-242">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="821fc-243">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-243">PivotLabelFilter</span></span>

<span data-ttu-id="821fc-244">Le deuxième extrait de code montre comment appliquer un [pivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) au champ de tableau croisé dynamique de **type,** en utilisant la propriété pour exclure les étiquettes qui commencent par la lettre `LabelFilterCondition.beginsWith` **L**.</span><span class="sxs-lookup"><span data-stu-id="821fc-244">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span>

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

##### <a name="pivotmanualfilter"></a><span data-ttu-id="821fc-245">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-245">PivotManualFilter</span></span>

<span data-ttu-id="821fc-246">Le troisième extrait de code applique un filtre manuel avec [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) au champ **Classification,** en filtrant les données qui n’incluent pas la classification **Organique**.</span><span class="sxs-lookup"><span data-stu-id="821fc-246">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span>

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="821fc-247">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="821fc-247">PivotValueFilter</span></span>

<span data-ttu-id="821fc-248">Pour comparer des nombres, utilisez un filtre de valeurs avec [PivotValueFilter,](/javascript/api/excel/excel.pivotvaluefilter)comme illustré dans l’extrait de code final.</span><span class="sxs-lookup"><span data-stu-id="821fc-248">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="821fc-249">Le tableau croisé dynamique compare les données du champ pivot de la batterie de serveurs aux données du champ PivotField ventes de caisses, y compris uniquement les batteries dont la somme des caisses vendues dépasse la valeur `PivotValueFilter` **500**.  </span><span class="sxs-lookup"><span data-stu-id="821fc-249">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span>

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

#### <a name="remove-pivotfilters"></a><span data-ttu-id="821fc-250">Supprimer des filtres de tableau croisé dynamique</span><span class="sxs-lookup"><span data-stu-id="821fc-250">Remove PivotFilters</span></span>

<span data-ttu-id="821fc-251">Pour supprimer tous les filtres de tableau croisé dynamique, appliquez la méthode à chaque champ de tableau croisé dynamique, comme `clearAllFilters` illustré dans l’exemple de code suivant.</span><span class="sxs-lookup"><span data-stu-id="821fc-251">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span>

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

### <a name="filter-with-slicers"></a><span data-ttu-id="821fc-252">Filtrer avec des slicers</span><span class="sxs-lookup"><span data-stu-id="821fc-252">Filter with slicers</span></span>

<span data-ttu-id="821fc-253">[Les slicers](/javascript/api/excel/excel.slicer) permettent de filtrer les données à partir d’Excel tableau croisé dynamique ou d’un tableau.</span><span class="sxs-lookup"><span data-stu-id="821fc-253">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="821fc-254">Un slicer utilise les valeurs d’une colonne spécifiée ou d’un champ de tableau croisé dynamique pour filtrer les lignes correspondantes.</span><span class="sxs-lookup"><span data-stu-id="821fc-254">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="821fc-255">Ces valeurs sont stockées en tant [qu’objets SlicerItem](/javascript/api/excel/excel.sliceritem) dans `Slicer` le .</span><span class="sxs-lookup"><span data-stu-id="821fc-255">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="821fc-256">Votre add-in peut ajuster ces filtres, tout comme les utilisateurs[(via l’interface Excel’interface utilisateur).](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)</span><span class="sxs-lookup"><span data-stu-id="821fc-256">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="821fc-257">Le slicer se trouve au-dessus de la feuille de calcul dans la couche de dessin, comme illustré dans la capture d’écran suivante.</span><span class="sxs-lookup"><span data-stu-id="821fc-257">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![Un slicer filtrant des données sur un tableau croisé dynamique.](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="821fc-259">Les techniques décrites dans cette section se concentrent sur l’utilisation de slicers connectés à des tableaux croisés dynamiques.</span><span class="sxs-lookup"><span data-stu-id="821fc-259">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="821fc-260">Les mêmes techniques s’appliquent également à l’utilisation de slicers connectés à des tables.</span><span class="sxs-lookup"><span data-stu-id="821fc-260">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="821fc-261">Créer un slicer</span><span class="sxs-lookup"><span data-stu-id="821fc-261">Create a slicer</span></span>

<span data-ttu-id="821fc-262">Vous pouvez créer un slicer dans un workbook ou une feuille de calcul à l’aide `Workbook.slicers.add` de la méthode ou de la `Worksheet.slicers.add` méthode.</span><span class="sxs-lookup"><span data-stu-id="821fc-262">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="821fc-263">Cela ajoute un slicer à [la SlicerCollection](/javascript/api/excel/excel.slicercollection) de l’objet `Workbook` ou `Worksheet` spécifié.</span><span class="sxs-lookup"><span data-stu-id="821fc-263">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="821fc-264">La `SlicerCollection.add` méthode a trois paramètres :</span><span class="sxs-lookup"><span data-stu-id="821fc-264">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="821fc-265">`slicerSource`: source de données sur laquelle repose le nouveau slicer.</span><span class="sxs-lookup"><span data-stu-id="821fc-265">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="821fc-266">Il peut s’agit d’une chaîne , ou d’une chaîne représentant `PivotTable` `Table` le nom ou l’ID d’un `PivotTable` ou `Table` .</span><span class="sxs-lookup"><span data-stu-id="821fc-266">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="821fc-267">`sourceField`: champ dans la source de données par lequel filtrer.</span><span class="sxs-lookup"><span data-stu-id="821fc-267">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="821fc-268">Il peut s’agit d’une chaîne , ou d’une chaîne représentant `PivotField` `TableColumn` le nom ou l’ID d’un `PivotField` ou `TableColumn` .</span><span class="sxs-lookup"><span data-stu-id="821fc-268">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="821fc-269">`slicerDestination`: feuille de calcul dans laquelle le nouveau slicer sera créé.</span><span class="sxs-lookup"><span data-stu-id="821fc-269">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="821fc-270">Il peut s’agit `Worksheet` d’un objet ou du nom ou de l’ID d’un `Worksheet` .</span><span class="sxs-lookup"><span data-stu-id="821fc-270">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="821fc-271">Ce paramètre est inutile lorsque le `SlicerCollection` paramètre est accessible via `Worksheet.slicers` .</span><span class="sxs-lookup"><span data-stu-id="821fc-271">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="821fc-272">Dans ce cas, la feuille de calcul de la collection est utilisée comme destination.</span><span class="sxs-lookup"><span data-stu-id="821fc-272">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="821fc-273">L’exemple de code suivant ajoute un nouveau slicer à la feuille de calcul **Pivot.**</span><span class="sxs-lookup"><span data-stu-id="821fc-273">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="821fc-274">La source du slicer  est le tableau croisé dynamique ventes de batterie de serveurs et filtre à l’aide des **données type.**</span><span class="sxs-lookup"><span data-stu-id="821fc-274">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="821fc-275">Le slicer est également nommé **Fruit Slicer pour** référence ultérieure.</span><span class="sxs-lookup"><span data-stu-id="821fc-275">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="821fc-276">Filtrer des éléments avec un slicer</span><span class="sxs-lookup"><span data-stu-id="821fc-276">Filter items with a slicer</span></span>

<span data-ttu-id="821fc-277">Le slicer filtre le tableau croisé dynamique avec les éléments du `sourceField` .</span><span class="sxs-lookup"><span data-stu-id="821fc-277">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="821fc-278">La `Slicer.selectItems` méthode définit les éléments qui restent dans le slicer.</span><span class="sxs-lookup"><span data-stu-id="821fc-278">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="821fc-279">Ces éléments sont transmis à la méthode en tant `string[]` que , représentant les clés des éléments.</span><span class="sxs-lookup"><span data-stu-id="821fc-279">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="821fc-280">Toutes les lignes contenant ces éléments restent dans l’agrégation du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-280">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="821fc-281">Appels suivants `selectItems` pour définir la liste sur les touches spécifiées dans ces appels.</span><span class="sxs-lookup"><span data-stu-id="821fc-281">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="821fc-282">Si un élément qui ne se trouve pas dans la source de données est transmis, une `Slicer.selectItems` `InvalidArgument` erreur est lancée.</span><span class="sxs-lookup"><span data-stu-id="821fc-282">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="821fc-283">Le contenu peut être vérifié par le biais de la propriété, qui est `Slicer.slicerItems` un [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span><span class="sxs-lookup"><span data-stu-id="821fc-283">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="821fc-284">L’exemple de code suivant montre trois éléments sélectionnés pour le slicer : **Sella,** **Tilleul** et **Orange**.</span><span class="sxs-lookup"><span data-stu-id="821fc-284">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="821fc-285">Pour supprimer tous les filtres du slicer, utilisez la `Slicer.clearFilters` méthode, comme illustré dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="821fc-285">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="821fc-286">Style et mise en forme d’un slicer</span><span class="sxs-lookup"><span data-stu-id="821fc-286">Style and format a slicer</span></span>

<span data-ttu-id="821fc-287">Vous pouvez ajuster les paramètres d’affichage d’un slicer par le biais de `Slicer` propriétés.</span><span class="sxs-lookup"><span data-stu-id="821fc-287">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="821fc-288">L’exemple de code suivant définit le style sur **SlicerStyleLight6,** définit le texte en haut du slicer sur **Types** de fruit, place le slicer à la position **(395, 15)** sur la couche de dessin et définit la taille du slicer à **135 x 150** pixels.</span><span class="sxs-lookup"><span data-stu-id="821fc-288">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

#### <a name="delete-a-slicer"></a><span data-ttu-id="821fc-289">Supprimer un slicer</span><span class="sxs-lookup"><span data-stu-id="821fc-289">Delete a slicer</span></span>

<span data-ttu-id="821fc-290">Pour supprimer un slicer, appelez la `Slicer.delete` méthode.</span><span class="sxs-lookup"><span data-stu-id="821fc-290">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="821fc-291">L’exemple de code suivant supprime le premier slicer de la feuille de calcul actuelle.</span><span class="sxs-lookup"><span data-stu-id="821fc-291">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="821fc-292">Modifier la fonction d’agrégation</span><span class="sxs-lookup"><span data-stu-id="821fc-292">Change aggregation function</span></span>

<span data-ttu-id="821fc-293">Les hiérarchies de données ont leurs valeurs agrégées.</span><span class="sxs-lookup"><span data-stu-id="821fc-293">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="821fc-294">Pour les jeux de données de nombres, il s’agit d’une somme par défaut.</span><span class="sxs-lookup"><span data-stu-id="821fc-294">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="821fc-295">La `summarizeBy` propriété définit ce comportement en fonction d’un type [AggregationFunction.](/javascript/api/excel/excel.aggregationfunction)</span><span class="sxs-lookup"><span data-stu-id="821fc-295">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="821fc-296">Les types de fonctions d’agrégation actuellement pris en charge sont `Sum` , , , , , , , , `Count` et `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` (par défaut).</span><span class="sxs-lookup"><span data-stu-id="821fc-296">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="821fc-297">Les exemples de code suivants modifient l’agrégation en moyenne des données.</span><span class="sxs-lookup"><span data-stu-id="821fc-297">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="821fc-298">Modifier les calculs avec une méthode ShowAsRule</span><span class="sxs-lookup"><span data-stu-id="821fc-298">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="821fc-299">Par défaut, les tableaux croisés dynamiques agrègent les données de leurs hiérarchies de lignes et de colonnes indépendamment.</span><span class="sxs-lookup"><span data-stu-id="821fc-299">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="821fc-300">Un [objet ShowAsRule](/javascript/api/excel/excel.showasrule) modifie la hiérarchie de données en valeurs de sortie basées sur d’autres éléments du tableau croisé dynamique.</span><span class="sxs-lookup"><span data-stu-id="821fc-300">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="821fc-301">`ShowAsRule`L’objet possède trois propriétés :</span><span class="sxs-lookup"><span data-stu-id="821fc-301">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="821fc-302">`calculation`: type de calcul relatif à appliquer à la hiérarchie de données (la valeur par défaut est `none` ).</span><span class="sxs-lookup"><span data-stu-id="821fc-302">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="821fc-303">`baseField`: [PivotField dans](/javascript/api/excel/excel.pivotfield) la hiérarchie contenant les données de base avant l’application du calcul.</span><span class="sxs-lookup"><span data-stu-id="821fc-303">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="821fc-304">Étant Excel tableaux croisés dynamiques ont un mappage un-à-un de hiérarchie à champ, vous utiliserez le même nom pour accéder à la hiérarchie et au champ.</span><span class="sxs-lookup"><span data-stu-id="821fc-304">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="821fc-305">`baseItem`: Tableau [croisé dynamique individuel comparé](/javascript/api/excel/excel.pivotitem) aux valeurs des champs de base en fonction du type de calcul.</span><span class="sxs-lookup"><span data-stu-id="821fc-305">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="821fc-306">Tous les calculs ne nécessitent pas ce champ.</span><span class="sxs-lookup"><span data-stu-id="821fc-306">Not all calculations require this field.</span></span>

<span data-ttu-id="821fc-307">L’exemple suivant définit le calcul de la hiérarchie de données Somme des **caisses vendues** à la batterie de serveurs comme un pourcentage du total des colonnes.</span><span class="sxs-lookup"><span data-stu-id="821fc-307">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="821fc-308">Nous voulons toujours que la granularité s’étende au niveau du type de fruit. Nous allons donc utiliser la hiérarchie de ligne **Type** et son champ sous-jacent.</span><span class="sxs-lookup"><span data-stu-id="821fc-308">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="821fc-309">La batterie  de serveurs est également la première hiérarchie de ligne de l’exemple, de sorte que le nombre total d’entrées de la batterie de serveurs affiche également le pourcentage que chaque batterie de serveurs est responsable de la production.</span><span class="sxs-lookup"><span data-stu-id="821fc-309">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="821fc-311">L’exemple précédent a fixé le calcul à la colonne, par rapport au champ d’une hiérarchie de lignes individuelle.</span><span class="sxs-lookup"><span data-stu-id="821fc-311">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="821fc-312">Lorsque le calcul est lié à un élément individuel, utilisez la `baseItem` propriété.</span><span class="sxs-lookup"><span data-stu-id="821fc-312">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="821fc-313">L’exemple suivant montre le `differenceFrom` calcul.</span><span class="sxs-lookup"><span data-stu-id="821fc-313">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="821fc-314">Il affiche la différence entre les entrées de hiérarchie des données de ventes de la batterie de serveurs par rapport à celles des **batteries de serveurs A**.</span><span class="sxs-lookup"><span data-stu-id="821fc-314">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="821fc-315">Il s’agit d’une batterie de serveurs, ce qui nous permet de voir les différences entre les autres batteries de serveurs, ainsi que les répartitions pour chaque type de fruit comme ( Le type est également une hiérarchie de lignes dans `baseField` cet exemple). </span><span class="sxs-lookup"><span data-stu-id="821fc-315">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![Tableau croisé dynamique montrant les différences de ventes de fruit entre les « batteries de serveurs » et les autres.](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="821fc-319">Modifier les noms de hiérarchie</span><span class="sxs-lookup"><span data-stu-id="821fc-319">Change hierarchy names</span></span>

<span data-ttu-id="821fc-320">Les champs de hiérarchie sont modifiables.</span><span class="sxs-lookup"><span data-stu-id="821fc-320">Hierarchy fields are editable.</span></span> <span data-ttu-id="821fc-321">Le code suivant montre comment modifier les noms affichés de deux hiérarchies de données.</span><span class="sxs-lookup"><span data-stu-id="821fc-321">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="821fc-322">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="821fc-322">See also</span></span>

- [<span data-ttu-id="821fc-323">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="821fc-323">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="821fc-324">Excel Référence de l’API JavaScript</span><span class="sxs-lookup"><span data-stu-id="821fc-324">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
