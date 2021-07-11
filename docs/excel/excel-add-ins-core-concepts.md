---
title: Modèle d’objet JavaScript Excel dans les compléments Office
description: Découvrez les types d’objets clés dans les API JavaScript Excel et comment les utiliser pour créer des compléments Excel.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6c88dc84796d9fd898bee880035ed964ab6cd7c8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349559"
---
# <a name="excel-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="dbd0d-103">Modèle d’objet JavaScript Excel dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="dbd0d-103">Excel JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="dbd0d-104">Cet article décrit comment utiliser l’[API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md) afin de créer des compléments pour Excel 2016 ou versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="dbd0d-105">Il présente les concepts fondamentaux de l’utilisation des API et fournit des conseils pour effectuer des tâches spécifiques, comme la lecture ou l’écriture d’une grande plage, la mise à jour de toutes les cellules d’une plage, et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dbd0d-106">Pour en savoir plus sur la nature asynchrone des API Excel et la manière dont elles fonctionnent avec le classeur, voir [Utilisation du modèle d’API spécifique à l’application](../develop/application-specific-api-model.md).</span><span class="sxs-lookup"><span data-stu-id="dbd0d-106">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.</span></span>  

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="dbd0d-107">API Office.js pour Excel</span><span class="sxs-lookup"><span data-stu-id="dbd0d-107">Office.js APIs for Excel</span></span>

<span data-ttu-id="dbd0d-108">Un complément Excel interagit avec des objets dans Excel en utilisant l’API Office JavaScript, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="dbd0d-108">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="dbd0d-109">**API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](../reference/overview/excel-add-ins-reference-overview.md) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-109">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="dbd0d-110">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-110">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="dbd0d-p102">Vous utiliserez probablement l’API JavaScript Excel pour développer la majorité des fonctionnalités des compléments destinés à Excel 2016 ou version ultérieure, vous utiliserez également des objets dans l’API commune. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="dbd0d-p102">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API. For example:</span></span>

* <span data-ttu-id="dbd0d-113">[Context](/javascript/api/office/office.context) :le `Context` représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-113">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="dbd0d-114">Il se compose de détails sur la configuration du classeur comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-114">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="dbd0d-115">En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-115">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="dbd0d-116">[Document](/javascript/api/office/office.document) : le `Document` fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-116">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="dbd0d-117">L’image suivante illustre les situations dans lesquelles vous pouvez utiliser l’API JavaScript Excel ou les API communes.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-117">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Différences entre l’API JS Excel et les API courantes.](../images/excel-js-api-common-api.png)

## <a name="excel-specific-object-model"></a><span data-ttu-id="dbd0d-119">Modèle d’objet spécifique à Excel</span><span class="sxs-lookup"><span data-stu-id="dbd0d-119">Excel-specific object model</span></span>

<span data-ttu-id="dbd0d-120">Pour comprendre les API Excel, vous devez connaître la manière dont les composants d’un classeur sont liés les uns aux autres.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-120">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

* <span data-ttu-id="dbd0d-121">Un **classeur** contient une ou plusieurs **feuilles de calcul**.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-121">A **Workbook** contains one or more **Worksheets**.</span></span>
* <span data-ttu-id="dbd0d-122">Une **Feuille de calcul** contient les collections de ces objets de données présents sur la feuille individuelle, et donne accès aux cellules via des objets de la **Plage**.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-122">A **Worksheet** contains collections of those data objects that are present in the individual sheet, and gives access to cells through **Range** objects.</span></span>
* <span data-ttu-id="dbd0d-123">Une **plage** représente un groupe de cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-123">A **Range** represents a group of contiguous cells.</span></span>
* <span data-ttu-id="dbd0d-124">Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-124">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
* <span data-ttu-id="dbd0d-125">Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-125">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

### <a name="ranges"></a><span data-ttu-id="dbd0d-126">Plages</span><span class="sxs-lookup"><span data-stu-id="dbd0d-126">Ranges</span></span>

<span data-ttu-id="dbd0d-127">Une plage est un groupe de cellules contiguës dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-127">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="dbd0d-128">Les compléments utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la colonne **B** et de la ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et des lignes **2** à **4**) pour définir les plages.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-128">Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="dbd0d-129">Les plages comportent trois propriétés principales : `values`, `formulas`et `format`.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-129">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="dbd0d-130">Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-130">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="dbd0d-131">Exemple de plage</span><span class="sxs-lookup"><span data-stu-id="dbd0d-131">Range sample</span></span>

<span data-ttu-id="dbd0d-132">L’exemple de code suivant montre comment créer des registres des ventes.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-132">The following sample shows how to create sales records.</span></span> <span data-ttu-id="dbd0d-133">Cette fonction utilise les objets `Range` pour déterminer les valeurs, les formules et les formats.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-133">This function uses `Range` objects to set the values, formulas, and formats.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Create the headers and format them to stand out.
    var headers = [
      ["Product", "Quantity", "Unit Price", "Totals"]
    ];
    var headerRange = sheet.getRange("B2:E2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";

    // Create the product data rows.
    var productData = [
      ["Almonds", 6, 7.5],
      ["Coffee", 20, 34.5],
      ["Chocolate", 10, 9.56],
    ];
    var dataRange = sheet.getRange("B3:D5");
    dataRange.values = productData;

    // Create the formulas to total the amounts sold.
    var totalFormulas = [
      ["=C3 * D3"],
      ["=C4 * D4"],
      ["=C5 * D5"],
      ["=SUM(E3:E5)"]
    ];
    var totalRange = sheet.getRange("E3:E6");
    totalRange.formulas = totalFormulas;
    totalRange.format.font.bold = true;

    // Display the totals as US dollar amounts.
    totalRange.numberFormat = [["$0.00"]];

    return context.sync();
});
```

<span data-ttu-id="dbd0d-134">Cet exemple crée les données suivantes dans la feuille de calcul active.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-134">This sample creates the following data in the current worksheet.</span></span>

![Un registre des ventes affiche des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/excel-overview-range-sample.png)

<span data-ttu-id="dbd0d-136">Pour plus d’informations, consultez [Définir et obtenir des valeurs de plage, un texte, ou des formules à l’aide de l’API JavaScript Excel](excel-add-ins-ranges-set-get-values.md).</span><span class="sxs-lookup"><span data-stu-id="dbd0d-136">For more information, see [Set and get range values, text, or formulas using the Excel JavaScript API](excel-add-ins-ranges-set-get-values.md).</span></span>

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="dbd0d-137">Graphiques, tableaux et autres objets de données</span><span class="sxs-lookup"><span data-stu-id="dbd0d-137">Charts, tables, and other data objects</span></span>

<span data-ttu-id="dbd0d-138">Les API JavaScript Excel peuvent créer et manipuler les structures de données et les visualisations dans Excel.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-138">The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="dbd0d-139">Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-139">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="dbd0d-140">Création d’un tableau</span><span class="sxs-lookup"><span data-stu-id="dbd0d-140">Creating a table</span></span>

<span data-ttu-id="dbd0d-p108">Créez des tableaux à l’aide de plages remplies de données. Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-p108">Create tables by using data-filled ranges. Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="dbd0d-143">L’exemple suivant crée un tableau à l’aide des plages de l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-143">The following sample creates a table using the ranges from the previous sample.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

<span data-ttu-id="dbd0d-144">L’exécution de cet exemple de code sur la feuille de calcul avec les données précédentes crée le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-144">Using this sample code on the worksheet with the previous data creates the following table.</span></span>

![Un tableau créée à partir du registre des ventes précédent.](../images/excel-overview-table-sample.png)

<span data-ttu-id="dbd0d-146">Pour plus d’informations, consultez [Travailler avec des tableaux l’aide de l’API JavaScript Excel](excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="dbd0d-146">For more information, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>

#### <a name="creating-a-chart"></a><span data-ttu-id="dbd0d-147">Création d’un graphique</span><span class="sxs-lookup"><span data-stu-id="dbd0d-147">Creating a chart</span></span>

<span data-ttu-id="dbd0d-148">Vous pouvez créer un graphique pour visualiser les données d’une plage.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-148">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="dbd0d-149">Les API prennent en charge des dizaines de variétés de graphiques, chacun pouvant être personnalisé selon vos besoins.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-149">The APIs support dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="dbd0d-150">L’exemple suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-150">The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

<span data-ttu-id="dbd0d-151">L’exécution de cet exemple sur la feuille de calcul avec le tableau précédent crée le graphique suivant.</span><span class="sxs-lookup"><span data-stu-id="dbd0d-151">Running this sample on the worksheet with the previous table creates the following chart.</span></span>

![Histogramme montrant les quantités pour trois des articles présents dans le registre des ventes précédent.](../images/excel-overview-chart-sample.png)

<span data-ttu-id="dbd0d-153">Pour plus d’informations, consultez [Travailler avec des graphiques l’aide de l’API JavaScript Excel](excel-add-ins-charts.md).</span><span class="sxs-lookup"><span data-stu-id="dbd0d-153">For more information, see [Work with charts using the Excel JavaScript API](excel-add-ins-charts.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="dbd0d-154">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dbd0d-154">See also</span></span>

* [<span data-ttu-id="dbd0d-155">Création de votre premier complément Excel</span><span class="sxs-lookup"><span data-stu-id="dbd0d-155">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
* [<span data-ttu-id="dbd0d-156">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="dbd0d-156">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="dbd0d-157">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="dbd0d-157">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
* [<span data-ttu-id="dbd0d-158">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="dbd0d-158">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
