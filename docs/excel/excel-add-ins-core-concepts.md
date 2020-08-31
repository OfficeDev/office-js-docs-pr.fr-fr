---
title: Concepts fondamentaux de programmation avec l’API JavaScript pour Excel
description: Utilisez l’API JavaScript pour Excel afin de créer des compléments pour Excel.
ms.date: 07/28/2020
localization_priority: Priority
ms.openlocfilehash: dde7dc66e0746fc4d9cf91ed3df824fab05c109d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292592"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="b18d5-103">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b18d5-103">Fundamental programming concepts with the Excel JavaScript API</span></span>

<span data-ttu-id="b18d5-104">Cet article décrit comment utiliser l’[API JavaScript pour Excel](../reference/overview/excel-add-ins-reference-overview.md) afin de créer des compléments pour Excel 2016 ou versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="b18d5-104">This article describes how to use the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="b18d5-105">Il présente les concepts fondamentaux de l’utilisation des API et fournit des conseils pour effectuer des tâches spécifiques, comme la lecture ou l’écriture d’une grande plage, la mise à jour de toutes les cellules d’une plage, et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="b18d5-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b18d5-106">Pour en savoir plus sur la nature asynchrone des API Excel et la manière dont elles fonctionnent avec le classeur, voir [Utilisation du modèle d’API spécifique à l’application](../develop/application-specific-api-model.md).</span><span class="sxs-lookup"><span data-stu-id="b18d5-106">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Excel APIs and how they work with the workbook.</span></span>  

## <a name="officejs-apis-for-excel"></a><span data-ttu-id="b18d5-107">API Office.js pour Excel</span><span class="sxs-lookup"><span data-stu-id="b18d5-107">Office.js APIs for Excel</span></span>

<span data-ttu-id="b18d5-108">Un complément Excel interagit avec des objets dans Excel en utilisant l’API Office JavaScript, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="b18d5-108">An Excel add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="b18d5-109">**API JavaScript pour Excel** : inclut dans Office 2016, l’[API JavaScript Excel](../reference/overview/excel-add-ins-reference-overview.md) fournit des objets fortement typés que vous pouvez utiliser pour accéder à des feuilles de calcul, des plages, des tableaux, des graphiques et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="b18d5-109">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span>

* <span data-ttu-id="b18d5-110">**API communes** : incluses dans Office 2013, les [API communes](/javascript/api/office) peuvent être utilisées pour accéder à des fonctionnalités telles qu’une interface utilisateur, des boîtes de dialogue et des paramètres du client, qui sont communes à plusieurs types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="b18d5-110">**Common APIs**: Introduced with Office 2013, the [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="b18d5-111">Vous utiliserez probablement l’API JavaScript Excel pour développer la majorité des fonctionnalités des compléments destinés à Excel 2016 ou version ultérieure, vous utiliserez également des objets dans l’API commune.</span><span class="sxs-lookup"><span data-stu-id="b18d5-111">While you'll likely use the Excel JavaScript API to develop the majority of functionality in add-ins that target Excel 2016 or later, you'll also use objects in the Common API.</span></span> <span data-ttu-id="b18d5-112">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="b18d5-112">For example:</span></span>

* <span data-ttu-id="b18d5-113">[Context](/javascript/api/office/office.context) :le `Context` représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API.</span><span class="sxs-lookup"><span data-stu-id="b18d5-113">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="b18d5-114">Il se compose de détails sur la configuration du classeur comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`.</span><span class="sxs-lookup"><span data-stu-id="b18d5-114">It consists of workbook configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="b18d5-115">En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si l’ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="b18d5-115">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether the specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="b18d5-116">[Document](/javascript/api/office/office.document) : le `Document` fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Excel dans lequel le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="b18d5-116">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Excel file where the add-in is running.</span></span>

<span data-ttu-id="b18d5-117">L’image suivante illustre les situations dans lesquelles vous pouvez utiliser l’API JavaScript Excel ou les API communes.</span><span class="sxs-lookup"><span data-stu-id="b18d5-117">The following image illustrates when you might use the Excel JavaScript API or the Common APIs.</span></span>

![Image des différences entre l’API Excel et les API communes](../images/excel-js-api-common-api.png)

## <a name="object-model"></a><span data-ttu-id="b18d5-119">Modèle d’objet</span><span class="sxs-lookup"><span data-stu-id="b18d5-119">Object model</span></span>

<span data-ttu-id="b18d5-120">Pour comprendre les API Excel, vous devez connaître la manière dont les composants d’un classeur sont liés les uns aux autres.</span><span class="sxs-lookup"><span data-stu-id="b18d5-120">To understand the Excel APIs, you must understand how the components of a workbook are related to one another.</span></span>

* <span data-ttu-id="b18d5-121">Un **classeur** contient une ou plusieurs **feuilles de calcul**.</span><span class="sxs-lookup"><span data-stu-id="b18d5-121">A **Workbook** contains one or more **Worksheets**.</span></span>
* <span data-ttu-id="b18d5-122">Une **feuille de calcul** donne accès à des cellules via **plage** objets.</span><span class="sxs-lookup"><span data-stu-id="b18d5-122">A **Worksheet** gives access to cells through **Range** objects.</span></span>
* <span data-ttu-id="b18d5-123">Une **plage** représente un groupe de cellules contiguës.</span><span class="sxs-lookup"><span data-stu-id="b18d5-123">A **Range** represents a group of contiguous cells.</span></span>
* <span data-ttu-id="b18d5-124">Les **plages** sont utilisées pour créer et placer des **tableaux**, des **graphiques**, des **formes** et d’autres objets d’organisation ou de visualisation de données.</span><span class="sxs-lookup"><span data-stu-id="b18d5-124">**Ranges** are used to create and place **Tables**, **Charts**, **Shapes**, and other data visualization or organization objects.</span></span>
* <span data-ttu-id="b18d5-125">Une **feuille de calcul** contient des collections d’objets de données présents dans la feuille individuelle.</span><span class="sxs-lookup"><span data-stu-id="b18d5-125">A **Worksheet** contains collections of those data objects that are present in the individual sheet.</span></span>
* <span data-ttu-id="b18d5-126">Les **classeurs** contiennent des collections de certains de ces objets de données (par exemple : les **tableaux**) pour l'ensemble du **classeur**.</span><span class="sxs-lookup"><span data-stu-id="b18d5-126">**Workbooks** contain collections of some of those data objects (such as **Tables**) for the entire **Workbook**.</span></span>

### <a name="ranges"></a><span data-ttu-id="b18d5-127">Plages</span><span class="sxs-lookup"><span data-stu-id="b18d5-127">Ranges</span></span>

<span data-ttu-id="b18d5-128">Une plage est un groupe de cellules contiguës dans le classeur.</span><span class="sxs-lookup"><span data-stu-id="b18d5-128">A range is a group of contiguous cells in the workbook.</span></span> <span data-ttu-id="b18d5-129">Les compléments utilisent généralement la notation de style A1 (par exemple : **B3** pour la cellule unique de la colonne **B** et de la ligne **3** ou **C2:F4** pour les cellules des colonnes **C** à **F** et des lignes **2** à **4**) pour définir les plages.</span><span class="sxs-lookup"><span data-stu-id="b18d5-129">Add-ins typically use A1-style notation (e.g. **B3** for the single cell in column **B** and row **3** or **C2:F4** for the cells from columns **C** through **F** and rows **2** through **4**) to define ranges.</span></span>

<span data-ttu-id="b18d5-130">Les plages comportent trois propriétés principales : `values`, `formulas`et `format`.</span><span class="sxs-lookup"><span data-stu-id="b18d5-130">Ranges have three core properties: `values`, `formulas`, and `format`.</span></span> <span data-ttu-id="b18d5-131">Ces propriétés obtiennent ou définissent les valeurs des cellules, les formules à évaluer et la mise en forme visuelle des cellules.</span><span class="sxs-lookup"><span data-stu-id="b18d5-131">These properties get or set the cell values, formulas to be evaluated, and the visual formatting of the cells.</span></span>

#### <a name="range-sample"></a><span data-ttu-id="b18d5-132">Exemple de plage</span><span class="sxs-lookup"><span data-stu-id="b18d5-132">Range sample</span></span>

<span data-ttu-id="b18d5-133">L’exemple de code suivant montre comment créer des registres des ventes.</span><span class="sxs-lookup"><span data-stu-id="b18d5-133">The following sample shows how to create sales records.</span></span> <span data-ttu-id="b18d5-134">Cette fonction utilise les objets `Range` pour déterminer les valeurs, les formules et les formats.</span><span class="sxs-lookup"><span data-stu-id="b18d5-134">This function uses `Range` objects to set the values, formulas, and formats.</span></span>

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

<span data-ttu-id="b18d5-135">Cet exemple crée les données suivantes dans la feuille de calcul active :</span><span class="sxs-lookup"><span data-stu-id="b18d5-135">This sample creates the following data in the current worksheet:</span></span>

![Un registre des ventes affiche des lignes de valeur, une colonne de formule et des en-têtes mis en forme.](../images/excel-overview-range-sample.png)

### <a name="charts-tables-and-other-data-objects"></a><span data-ttu-id="b18d5-137">Graphiques, tableaux et autres objets de données</span><span class="sxs-lookup"><span data-stu-id="b18d5-137">Charts, tables, and other data objects</span></span>

<span data-ttu-id="b18d5-138">Les API JavaScript Excel peuvent créer et manipuler les structures de données et les visualisations dans Excel.</span><span class="sxs-lookup"><span data-stu-id="b18d5-138">The Excel JavaScript APIs can create and manipulate the data structures and visualizations within Excel.</span></span> <span data-ttu-id="b18d5-139">Les tableaux et les graphiques sont deux des objets les plus fréquemment utilisés, mais les API prennent en charge les tableaux croisés dynamiques, les formes, les images et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="b18d5-139">Tables and charts are two of the more commonly used objects, but the APIs support PivotTables, shapes, images, and more.</span></span>

#### <a name="creating-a-table"></a><span data-ttu-id="b18d5-140">Création d’un tableau</span><span class="sxs-lookup"><span data-stu-id="b18d5-140">Creating a table</span></span>

<span data-ttu-id="b18d5-141">Créez des tableaux à l’aide des plages de données remplies.</span><span class="sxs-lookup"><span data-stu-id="b18d5-141">Create tables by using data-filled ranges.</span></span> <span data-ttu-id="b18d5-142">Les contrôles de mise en forme et du tableau (par exemple, les filtres) sont automatiquement appliqués à la plage.</span><span class="sxs-lookup"><span data-stu-id="b18d5-142">Formatting and table controls (such as filters) are automatically applied to the range.</span></span>

<span data-ttu-id="b18d5-143">L’exemple suivant crée un tableau à l’aide des plages de l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="b18d5-143">The following sample creates a table using the ranges from the previous sample.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.tables.add("B2:E5", true);
    return context.sync();
});
```

<span data-ttu-id="b18d5-144">L’exécution de cet exemple de code sur la feuille de calcul avec les données précédentes crée le tableau suivant :</span><span class="sxs-lookup"><span data-stu-id="b18d5-144">Using this sample code on the worksheet with the previous data creates the following table:</span></span>

![Un tableau créée à partir du registre des ventes précédent.](../images/excel-overview-table-sample.png)

#### <a name="creating-a-chart"></a><span data-ttu-id="b18d5-146">Création d’un graphique</span><span class="sxs-lookup"><span data-stu-id="b18d5-146">Creating a chart</span></span>

<span data-ttu-id="b18d5-147">Vous pouvez créer un graphique pour visualiser les données d’une plage.</span><span class="sxs-lookup"><span data-stu-id="b18d5-147">Create charts to visualize the data in a range.</span></span> <span data-ttu-id="b18d5-148">Les API prennent en charge des dizaines de variétés de graphiques, chacun pouvant être personnalisé selon vos besoins.</span><span class="sxs-lookup"><span data-stu-id="b18d5-148">The APIs support dozens of chart varieties, each of which can be customized to suit your needs.</span></span>

<span data-ttu-id="b18d5-149">L’exemple suivant crée un histogramme pour trois éléments et place celui-ci 100 pixels en dessous de la partie supérieure de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="b18d5-149">The following sample creates a simple column chart for three items and places it 100 pixels below the top of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var chart = sheet.charts.add(Excel.ChartType.columnStacked, sheet.getRange("B3:C5"));
    chart.top = 100;
    return context.sync();
});
```

<span data-ttu-id="b18d5-150">L’exécution de cet exemple sur la feuille de calcul avec le tableau précédent crée le graphique suivant :</span><span class="sxs-lookup"><span data-stu-id="b18d5-150">Running this sample on the worksheet with the previous table creates the following chart:</span></span>

![Histogramme montrant les quantités pour trois des articles présents dans le registre des ventes précédent.](../images/excel-overview-chart-sample.png)

## <a name="run-options"></a><span data-ttu-id="b18d5-152">Options d’exécution</span><span class="sxs-lookup"><span data-stu-id="b18d5-152">Run options</span></span>

<span data-ttu-id="b18d5-153">`Excel.run` est associé à une surcharge liée à un objet [RunOptions](/javascript/api/excel/excel.runoptions).</span><span class="sxs-lookup"><span data-stu-id="b18d5-153">`Excel.run` has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="b18d5-154">Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="b18d5-154">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="b18d5-155">La propriété suivante est actuellement prise en charge :</span><span class="sxs-lookup"><span data-stu-id="b18d5-155">The following property is currently supported:</span></span>

* <span data-ttu-id="b18d5-156">`delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule.</span><span class="sxs-lookup"><span data-stu-id="b18d5-156">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="b18d5-157">Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule.</span><span class="sxs-lookup"><span data-stu-id="b18d5-157">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="b18d5-158">Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur).</span><span class="sxs-lookup"><span data-stu-id="b18d5-158">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="b18d5-159">Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.</span><span class="sxs-lookup"><span data-stu-id="b18d5-159">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="null-or-blank-property-values"></a><span data-ttu-id="b18d5-160">Valeurs de propriété null ou vides</span><span class="sxs-lookup"><span data-stu-id="b18d5-160">null or blank property values</span></span>

<span data-ttu-id="b18d5-161">`null` et les chaînes vides ont des implications particulières dans les API JavaScript Excel.</span><span class="sxs-lookup"><span data-stu-id="b18d5-161">`null` and empty strings have special implications in the Excel JavaScript APIs.</span></span> <span data-ttu-id="b18d5-162">Elles sont utilisées pour représenter les cellules vides, l’absence de mise en forme ou les valeurs par défaut.</span><span class="sxs-lookup"><span data-stu-id="b18d5-162">They're used to represent empty cells, no formatting, or default values.</span></span> <span data-ttu-id="b18d5-163">Cette section décrit l’utilisation de `null` et d’une chaîne vide lors de l’obtention et de la définition de propriétés.</span><span class="sxs-lookup"><span data-stu-id="b18d5-163">This section details the use of `null` and empty string when getting and setting properties.</span></span>

### <a name="null-input-in-2-d-array"></a><span data-ttu-id="b18d5-164">entrée de valeurs null dans un tableau 2D</span><span class="sxs-lookup"><span data-stu-id="b18d5-164">null input in 2-D Array</span></span>

<span data-ttu-id="b18d5-p113">Dans Excel, une plage est représentée par un tableau 2D, où les lignes représentent la première dimension et les colonnes la deuxième. Pour définir des valeurs, un format de nombre ou une formule uniquement pour des cellules spécifiques dans une plage, spécifiez des valeurs, un format de nombre ou une formule pour ces cellules dans le tableau 2D, et indiquez `null` pour toutes les autres cellules du tableau 2D.</span><span class="sxs-lookup"><span data-stu-id="b18d5-p113">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns. To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>

<span data-ttu-id="b18d5-p114">Par exemple, pour mettre à jour le format de nombre pour une seule cellule dans une plage et conserver le format de nombre existant pour toutes les autres cellules de la plage, spécifiez le nouveau format de nombre de la cellule à mettre à jour, puis spécifiez `null` pour toutes les autres cellules. L’extrait de code suivant définit un nouveau format de nombre pour la quatrième cellule de la plage et ne modifie pas le format de nombre pour les trois premières cellules de la plage.</span><span class="sxs-lookup"><span data-stu-id="b18d5-p114">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells. The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

### <a name="null-input-for-a-property"></a><span data-ttu-id="b18d5-169">Entrée null pour une propriété</span><span class="sxs-lookup"><span data-stu-id="b18d5-169">null input for a property</span></span>

<span data-ttu-id="b18d5-p115">`null` n’est pas une entrée valide pour une propriété unique. Par exemple, l’extrait de code suivant n’est pas valide, car la propriété `values` de la plage ne peut pas être définie sur `null`.</span><span class="sxs-lookup"><span data-stu-id="b18d5-p115">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the `values` property of the range cannot be set to `null`.</span></span>

```js
range.values = null;
```

<span data-ttu-id="b18d5-172">De même, l’extrait de code suivant n’est pas valide, car `null` n’est pas une valeur valide pour la propriété `color`.</span><span class="sxs-lookup"><span data-stu-id="b18d5-172">Likewise, the following code snippet is not valid, as `null` is not a valid value for the `color` property.</span></span>

```js
range.format.fill.color =  null;
```

### <a name="null-property-values-in-the-response"></a><span data-ttu-id="b18d5-173">valeurs de la propriété Null dans la réponse</span><span class="sxs-lookup"><span data-stu-id="b18d5-173">null property values in the response</span></span>

<span data-ttu-id="b18d5-p116">Les propriétés de mise en forme comme `size` et `color` contiendront des valeurs `null` dans la réponse lorsque différentes valeurs existent dans la plage spécifiée. Par exemple, si vous récupérez une plage et chargez sa propriété `format.font.color`:</span><span class="sxs-lookup"><span data-stu-id="b18d5-p116">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range. For example, if you retrieve a range and load its `format.font.color` property:</span></span>

* <span data-ttu-id="b18d5-176">Si toutes les cellules de la plage ont la même couleur de police, `range.format.font.color` spécifie cette couleur.</span><span class="sxs-lookup"><span data-stu-id="b18d5-176">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="b18d5-177">Si plusieurs couleurs de police sont présentes dans la plage, `range.format.font.color` est `null`.</span><span class="sxs-lookup"><span data-stu-id="b18d5-177">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>

### <a name="blank-input-for-a-property"></a><span data-ttu-id="b18d5-178">Entrée vide pour une propriété</span><span class="sxs-lookup"><span data-stu-id="b18d5-178">Blank input for a property</span></span>

<span data-ttu-id="b18d5-p117">Lorsque vous spécifiez une valeur vide pour une propriété (c’est-à-dire deux guillemets droits sans espace entre `''`), cela est interprété comme une instruction d’effacement ou de réinitialisation de la propriété. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="b18d5-p117">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>

* <span data-ttu-id="b18d5-181">Si vous spécifiez une valeur vide pour la propriété `values` d’une plage, le contenu de la plage est effacé.</span><span class="sxs-lookup"><span data-stu-id="b18d5-181">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
* <span data-ttu-id="b18d5-182">Si vous spécifiez une valeur vide pour la propriété `numberFormat`, le format de nombre est réinitialisé sur `General`.</span><span class="sxs-lookup"><span data-stu-id="b18d5-182">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
* <span data-ttu-id="b18d5-183">Si vous spécifiez une valeur vide pour les propriétés `formula` et `formulaLocale`, les valeurs de la formule sont effacées.</span><span class="sxs-lookup"><span data-stu-id="b18d5-183">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>

### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="b18d5-184">Valeurs de propriété vides dans la réponse</span><span class="sxs-lookup"><span data-stu-id="b18d5-184">Blank property values in the response</span></span>

<span data-ttu-id="b18d5-p118">Pour les opérations de lecture, une valeur de propriété vide dans la réponse (c'est-à-dire, deux guillemets droits sans espace entre `''`) indique que la cellule ne contient pas de donnée ni de valeur. Dans le premier exemple ci-dessous, la première et la dernière cellules de la plage ne contiennent pas de donnée. Dans le deuxième exemple, les deux premières cellules de la plage ne contiennent pas de formule.</span><span class="sxs-lookup"><span data-stu-id="b18d5-p118">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value. In the first example below, the first and last cell in the range contain no data. In the second example, the first two cells in the range do not contain a formula.</span></span>

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```

## <a name="requirement-sets"></a><span data-ttu-id="b18d5-188">Ensembles de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b18d5-188">Requirement sets</span></span>

<span data-ttu-id="b18d5-189">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="b18d5-189">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="b18d5-190">Le complément Office peut effectuer une vérification à l’exécution ou utiliser des ensembles de conditions requises spécifiés dans le manifeste pour déterminer si une application Office prend en charge les API requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="b18d5-190">An Office Add-in can perform a runtime check or use requirement sets specified in the manifest to determine whether an Office application supports the APIs that the add-in needs.</span></span> <span data-ttu-id="b18d5-191">Pour identifier les ensembles de conditions requises spécifiques disponibles sur chaque plateforme prise en charge, reportez-vous à [Ensembles de conditions requises de l’API JavaScript pour Excel](../reference/requirement-sets/excel-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b18d5-191">To identify the specific requirement sets that are available on each supported platform, see [Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md).</span></span>

### <a name="checking-for-requirement-set-support-at-runtime"></a><span data-ttu-id="b18d5-192">Vérification de la prise en charge de l’ensemble de conditions requises à l’exécution</span><span class="sxs-lookup"><span data-stu-id="b18d5-192">Checking for requirement set support at runtime</span></span>

<span data-ttu-id="b18d5-193">L’exemple de code suivant montre comment déterminer si l’application Office dans laquelle le complément est en cours d’exécution prend en charge l’ensemble spécifié de conditions requises pour l’API.</span><span class="sxs-lookup"><span data-stu-id="b18d5-193">The following code sample shows how to determine whether the Office application where the add-in is running supports the specified API requirement set.</span></span>

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a><span data-ttu-id="b18d5-194">Définition de la prise en charge de l’ensemble de conditions requises dans le manifeste</span><span class="sxs-lookup"><span data-stu-id="b18d5-194">Defining requirement set support in the manifest</span></span>

<span data-ttu-id="b18d5-195">Vous pouvez utiliser l’[élément Requirements](../reference/manifest/requirements.md) dans le manifeste de complément pour spécifier les ensembles de conditions requises minimales et/ou les méthodes d’API que votre complément doit activer.</span><span class="sxs-lookup"><span data-stu-id="b18d5-195">You can use the [Requirements element](../reference/manifest/requirements.md) in the add-in manifest to specify the minimal requirement sets and/or API methods that your add-in requires to activate.</span></span> <span data-ttu-id="b18d5-196">Si la plateforme ou l’application Office ne prend pas en charge les ensembles de conditions requises ou les méthodes d’API spécifiées dans l’élément `Requirements` du manifeste, le complément ne s’exécute pas dans cette application ou plateforme et ne s’affiche pas dans la liste de compléments dans **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="b18d5-196">If the Office application or platform doesn't support the requirement sets or API methods that are specified in the `Requirements` element of the manifest, the add-in won't run in that application or platform, and it won't display in the list of add-ins that are shown in **My Add-ins**.</span></span>

<span data-ttu-id="b18d5-197">L’exemple de code suivant montre l’élément `Requirements` dans un manifeste indiquant que le complément doit être chargé dans toutes les applications clientes Office prenant en charge l’ensemble de conditions requises ExcelApi version 1.3 ou ultérieure.</span><span class="sxs-lookup"><span data-stu-id="b18d5-197">The following code sample shows the `Requirements` element in an add-in manifest which specifies that the add-in should load in all Office client applications that support ExcelApi requirement set version 1.3 or greater.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

> [!NOTE]
> <span data-ttu-id="b18d5-198">Pour rendre votre complément disponible sur toutes les plateformes d’une application Office, comme Excel sur le web, Windows et iPad, nous vous recommandons de vérifier la prise en charge des conditions requises lors de l’exécution au lieu de définir la prise en charge d’ensemble de conditions requises dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="b18d5-198">To make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad, we recommend that you check for requirement support at runtime instead of defining requirement set support in the manifest.</span></span>

### <a name="requirement-sets-for-the-officejs-common-api"></a><span data-ttu-id="b18d5-199">Ensembles de conditions requises pour l’API commune Office.js</span><span class="sxs-lookup"><span data-stu-id="b18d5-199">Requirement sets for the Office.js Common API</span></span>

<span data-ttu-id="b18d5-200">Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](../reference/requirement-sets/office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b18d5-200">For information about Common API requirement sets, see [Office Common API requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

## <a name="handle-errors"></a><span data-ttu-id="b18d5-201">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="b18d5-201">Handle errors</span></span>

<span data-ttu-id="b18d5-202">Lorsqu’une erreur d’API se produit, l’API renvoie un objet `error` qui contient un code et un message.</span><span class="sxs-lookup"><span data-stu-id="b18d5-202">When an API error occurs, the API returns an `error` object that contains a code and a message.</span></span> <span data-ttu-id="b18d5-203">Pour plus d’informations sur la gestion des erreurs, notamment la liste des erreurs d’API, consultez la rubrique [Gestion des erreurs](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="b18d5-203">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b18d5-204">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="b18d5-204">See also</span></span>

* [<span data-ttu-id="b18d5-205">Création de votre premier complément Excel</span><span class="sxs-lookup"><span data-stu-id="b18d5-205">Build your first Excel add-in</span></span>](../quickstarts/excel-quickstart-jquery.md)
* [<span data-ttu-id="b18d5-206">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="b18d5-206">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="b18d5-207">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="b18d5-207">Excel JavaScript API performance optimization</span></span>](../excel/performance.md)
* [<span data-ttu-id="b18d5-208">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="b18d5-208">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
* [<span data-ttu-id="b18d5-209">Problèmes courants liés au code et comportements de plateforme inattendus</span><span class="sxs-lookup"><span data-stu-id="b18d5-209">Common coding issues and unexpected platform behaviors</span></span>](../develop/common-coding-issues.md)
