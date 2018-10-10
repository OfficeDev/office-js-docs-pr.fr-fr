---
title: Utiliser des graphiques à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 80b537ec66caf6e173dfe4453a257c5963156e6f
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459300"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="d5117-102">Utiliser des graphiques à l’aide de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="d5117-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="d5117-p101">Cet article fournit des exemples de code qui montrent comment effectuer des tâches courantes avec des graphiques à l’aide de l’API JavaScript Excel. Pour obtenir la liste complète des propriétés et méthodes prises en charge par les objets **Chart** et **ChartCollection**, voir l’[Objet Chart (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart?view=office-js) et l’[Objet Collection Graphique (API JavaScript pour Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="d5117-p101">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API. For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart?view=office-js) and [Chart Collection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection?view=office-js).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="d5117-105">Créer un graphique</span><span class="sxs-lookup"><span data-stu-id="d5117-105">Create a chart</span></span>

<span data-ttu-id="d5117-p102">L’exemple de code suivant crée un graphique dans la feuille de calcul nommée **Sample**. Le graphique est un graphique en **courbes** basé sur des données de la plage **A1:B13**.</span><span class="sxs-lookup"><span data-stu-id="d5117-p102">The following code sample creates a chart in the worksheet named **Sample**. The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var dataRange = sheet.getRange("A1:B13");
    var chart = sheet.charts.add("Line", dataRange, "auto");

    chart.title.text = "Sales Data";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-108">**Nouveau graphique en courbes**</span><span class="sxs-lookup"><span data-stu-id="d5117-108">**New line chart**</span></span>

![Nouveau graphique en courbes dans Excel](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="d5117-110">Ajouter une série de données à un graphique</span><span class="sxs-lookup"><span data-stu-id="d5117-110">Add a data series to a chart</span></span>

<span data-ttu-id="d5117-p103">L’exemple de code suivant ajoute une série de données au premier graphique de la feuille de calcul. La nouvelle série de données correspond à la colonne nommée **2016** et repose sur les données de la plage **D2:D5**.</span><span class="sxs-lookup"><span data-stu-id="d5117-p103">The following code sample adds a data series to the first chart in the worksheet. The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var chart = sheet.charts.getItemAt(0);
    var dataRange = sheet.getRange("D2:D5");

    var newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-113">**Graphique avant l’ajout de la série de données 2016**</span><span class="sxs-lookup"><span data-stu-id="d5117-113">**Chart before the 2016 data series is added**</span></span>

![Graphique dans Excel avant l’ajout de la série de données 2016](../images/excel-charts-data-series-before.png)

<span data-ttu-id="d5117-115">**Graphique après l’ajout de la série de données 2016**</span><span class="sxs-lookup"><span data-stu-id="d5117-115">**Chart after the 2016 data series is added**</span></span>

![Graphique dans Excel après l’ajout de la série de données 2016](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="d5117-117">Définir le titre du graphique</span><span class="sxs-lookup"><span data-stu-id="d5117-117">Set chart title</span></span>

<span data-ttu-id="d5117-118">L’exemple de code suivant définit le titre du premier graphique dans la feuille de calcul sur **Sales Data by Year**.</span><span class="sxs-lookup"><span data-stu-id="d5117-118">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-119">**Graphique après la définition du titre**</span><span class="sxs-lookup"><span data-stu-id="d5117-119">**Chart after title is set**</span></span>

![Graphique avec un titre dans Excel](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="d5117-121">Définir les propriétés d’un axe d’un graphique</span><span class="sxs-lookup"><span data-stu-id="d5117-121">Set properties of an axis in a chart</span></span>

<span data-ttu-id="d5117-122">Les graphiques qui utilisent le [système de coordonnées cartésiennes](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), comme les histogrammes, les graphiques à barres et les nuages de points, ont un axe des abscisses et un axe des ordonnées.</span><span class="sxs-lookup"><span data-stu-id="d5117-122">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis.</span></span> <span data-ttu-id="d5117-123">Ces exemples montrent comment définir le titre et afficher les unités d’un axe dans un graphique.</span><span class="sxs-lookup"><span data-stu-id="d5117-123">These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="d5117-124">Définir le titre d’un axe</span><span class="sxs-lookup"><span data-stu-id="d5117-124">Set axis title</span></span>

<span data-ttu-id="d5117-125">L’exemple de code suivant définit le titre de l’axe des abscisses pour le premier graphique de la feuille de calcul sur **Product**.</span><span class="sxs-lookup"><span data-stu-id="d5117-125">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-126">**Graphique après la définition de l’axe des abscisses**</span><span class="sxs-lookup"><span data-stu-id="d5117-126">**Chart after title of category axis is set**</span></span>

![Graphique avec un titre d’axe dans Excel](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="d5117-128">Définir l’unité d’affichage de l’axe</span><span class="sxs-lookup"><span data-stu-id="d5117-128">Set axis display unit</span></span>

<span data-ttu-id="d5117-129">L’exemple de code suivant définit l’unité d’affichage de l’axe des ordonnées pour le premier graphique de la feuille de calcul sur **Hundreds**.</span><span class="sxs-lookup"><span data-stu-id="d5117-129">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-130">**Graphique après la définition de l’unité d’affichage de l’axe des ordonnées**</span><span class="sxs-lookup"><span data-stu-id="d5117-130">**Chart after display unit of value axis is set**</span></span>

![Graphique avec l’unité d’affichage de l’axe dans Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="d5117-132">Définir la visibilité du quadrillage dans un graphique</span><span class="sxs-lookup"><span data-stu-id="d5117-132">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="d5117-133">L’exemple de code suivant masque le quadrillage principal de l’axe des ordonnées du premier graphique de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="d5117-133">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet.</span></span> <span data-ttu-id="d5117-134">Vous pouvez afficher le quadrillage principal de l’axe des ordonnées du graphique en définissant `chart.axes.valueAxis.majorGridlines.visible` sur **true**.</span><span class="sxs-lookup"><span data-stu-id="d5117-134">You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to **true**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-135">**Graphique avec du quadrillage masqué**</span><span class="sxs-lookup"><span data-stu-id="d5117-135">**Chart with gridlines hidden**</span></span>

![Graphique avec du quadrillage masqué dans Excel](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="d5117-137">Courbes de tendance de graphiques</span><span class="sxs-lookup"><span data-stu-id="d5117-137">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="d5117-138">Ajouter une courbe de tendance</span><span class="sxs-lookup"><span data-stu-id="d5117-138">Add a trendline</span></span>

<span data-ttu-id="d5117-p106">L’exemple de code suivant ajoute une courbe de tendance de moyenne mobile à la première série du premier graphique de la feuille de calcul nommée **Sample**. La courbe de tendance affiche une moyenne mobile sur 5 périodes.</span><span class="sxs-lookup"><span data-stu-id="d5117-p106">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-141">**Graphique avec courbe de tendance de moyenne mobile**</span><span class="sxs-lookup"><span data-stu-id="d5117-141">**Chart with moving average trendline**</span></span>

![Graphique avec courbe de tendance de moyenne mobile dans Excel](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="d5117-143">Mettre à jour une courbe de tendance</span><span class="sxs-lookup"><span data-stu-id="d5117-143">Update a trendline</span></span>

<span data-ttu-id="d5117-144">L’exemple de code suivant définit la courbe de tendance sur le type **Linear** pour la première série du premier graphique de la feuille de calcul nommée **Sample**.</span><span class="sxs-lookup"><span data-stu-id="d5117-144">The following code sample sets the trendline to type **Linear** for the first series in the first chart in the worksheet named **Sample**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    var series = seriesCollection.getItemAt(0);
    series.trendlines.getItem(0).type = "Linear";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="d5117-145">**Graphique avec une courbe de tendance linéaire**</span><span class="sxs-lookup"><span data-stu-id="d5117-145">**Chart with linear trendline**</span></span>

![Graphique avec une courbe de tendance linéaire dans Excel](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a><span data-ttu-id="d5117-147">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="d5117-147">See also</span></span>

- [<span data-ttu-id="d5117-148">Concepts fondamentaux de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="d5117-148">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
