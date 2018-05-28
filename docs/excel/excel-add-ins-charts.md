---
title: Utiliser des graphiques ? l?aide de l?API JavaScript pour Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: c0f45892cb937a565a6855390344855f75e7473e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="cc3cc-102">Utiliser des graphiques ? l?aide de l?API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="cc3cc-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="cc3cc-103">Cet article fournit des exemples de code qui montrent comment effectuer des t?ches courantes ? l?aide de graphiques utilisant l?API JavaScript pour Excel.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-103">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API.</span></span> <span data-ttu-id="cc3cc-104">Pour une liste compl?te des propri?t?s et des m?thodes prises en charge par les objets **Chart** et **ChartCollection**, reportez-vous ? la rubrique [Objet Chart (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/chart) et [Objet ChartCollection (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection).</span><span class="sxs-lookup"><span data-stu-id="cc3cc-104">For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/chart) and [Chart Collection Object (JavaScript API for Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="cc3cc-105">Cr?er un graphique</span><span class="sxs-lookup"><span data-stu-id="cc3cc-105">Create a chart</span></span>

<span data-ttu-id="cc3cc-106">L?exemple de code suivant cr?e un graphique dans la feuille de calcul nomm?e **Sample**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-106">The following code sample creates a chart in the worksheet named **Sample**.</span></span> <span data-ttu-id="cc3cc-107">Il s?agit d?un graphique en **courbes** qui est fond? sur les donn?es de la plage **A1:B13**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-107">The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

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

<span data-ttu-id="cc3cc-108">**Nouveau graphique en courbes**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-108">**New line chart**</span></span>

![Nouveau graphique en courbes dans Excel](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="cc3cc-110">Ajouter une s?rie de donn?es ? un graphique</span><span class="sxs-lookup"><span data-stu-id="cc3cc-110">Add a data series to a chart</span></span>

<span data-ttu-id="cc3cc-111">L?exemple de code suivant ajoute une s?rie de donn?es au premier graphique de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-111">The following code sample adds a data series to the first chart in the worksheet.</span></span> <span data-ttu-id="cc3cc-112">La nouvelle s?rie de donn?es correspond ? la colonne nomm?e **2016** et repose sur les donn?es de la plage **D2:D5**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-112">The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

> [!NOTE]
> <span data-ttu-id="cc3cc-113">Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta).</span><span class="sxs-lookup"><span data-stu-id="cc3cc-113">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="cc3cc-114">Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-114">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

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

<span data-ttu-id="cc3cc-115">**Graphique avant l?ajout de la s?rie de donn?es 2016**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-115">**Chart before the 2016 data series is added**</span></span>

![Graphique dans Excel avant l?ajout de la s?rie de donn?es 2016](../images/excel-charts-data-series-before.png)

<span data-ttu-id="cc3cc-117">**Graphique apr?s l?ajout de la s?rie de donn?es 2016**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-117">**Chart after the 2016 data series is added**</span></span>

![Graphique dans Excel apr?s l?ajout de la s?rie de donn?es 2016](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="cc3cc-119">D?finir le titre du graphique</span><span class="sxs-lookup"><span data-stu-id="cc3cc-119">Set chart title</span></span>

<span data-ttu-id="cc3cc-120">L?exemple de code suivant d?finit le titre du premier graphique dans la feuille de calcul sur **Sales Data by Year**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-120">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="cc3cc-121">**Graphique apr?s la d?finition du titre**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-121">**Chart after title is set**</span></span>

![Graphique avec un titre dans Excel](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="cc3cc-123">D?finir les propri?t?s d?un axe d?un graphique</span><span class="sxs-lookup"><span data-stu-id="cc3cc-123">Set properties of an axis in a chart</span></span>

<span data-ttu-id="cc3cc-124">Les graphiques qui utilisent le [syst?me de coordonn?es cart?siennes](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), comme les histogrammes, les graphiques ? barres et les nuages de points, ont un axe des abscisses et un axe des ordonn?es.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-124">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis.</span></span> <span data-ttu-id="cc3cc-125">Ces exemples montrent comment d?finir le titre et afficher les unit?s d?un axe dans un graphique.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-125">These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="cc3cc-126">D?finir le titre d?un axe</span><span class="sxs-lookup"><span data-stu-id="cc3cc-126">Set axis title</span></span>

<span data-ttu-id="cc3cc-127">L?exemple de code suivant d?finit le titre de l?axe des abscisses pour le premier graphique de la feuille de calcul sur **Product**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-127">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="cc3cc-128">**Graphique apr?s la d?finition de l?axe des abscisses**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-128">**Chart after title of category axis is set**</span></span>

![Graphique avec un titre d?axe dans Excel](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="cc3cc-130">D?finir l?unit? d?affichage de l?axe</span><span class="sxs-lookup"><span data-stu-id="cc3cc-130">Set axis display unit</span></span>

<span data-ttu-id="cc3cc-131">L?exemple de code suivant d?finit l?unit? d?affichage de l?axe des ordonn?es pour le premier graphique de la feuille de calcul sur **Hundreds**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-131">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

> [!NOTE]
> <span data-ttu-id="cc3cc-132">Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta).</span><span class="sxs-lookup"><span data-stu-id="cc3cc-132">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="cc3cc-133">Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-133">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="cc3cc-134">**Graphique apr?s la d?finition de l?unit? d?affichage de l?axe des ordonn?es**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-134">**Chart after display unit of value axis is set**</span></span>

![Graphique avec l?unit? d?affichage de l?axe dans Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="cc3cc-136">D?finir la visibilit? du quadrillage dans un graphique</span><span class="sxs-lookup"><span data-stu-id="cc3cc-136">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="cc3cc-137">L?exemple de code suivant masque le quadrillage principal de l?axe des ordonn?es du premier graphique de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-137">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet.</span></span> <span data-ttu-id="cc3cc-138">Vous pouvez afficher le quadrillage principal de l?axe des ordonn?es du graphique en d?finissant `chart.axes.valueAxis.majorGridlines.visible` sur **true**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-138">You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to **true**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="cc3cc-139">**Graphique avec du quadrillage masqu?**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-139">**Chart with gridlines hidden**</span></span>

![Graphique avec du quadrillage masqu? dans Excel](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="cc3cc-141">Courbes de tendance de graphiques</span><span class="sxs-lookup"><span data-stu-id="cc3cc-141">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="cc3cc-142">Ajouter une courbe de tendance</span><span class="sxs-lookup"><span data-stu-id="cc3cc-142">Add a trendline</span></span>

<span data-ttu-id="cc3cc-p108">L?exemple de code suivant ajoute une courbe de tendance de moyenne mobile ? la premi?re s?rie du premier graphique de la feuille de calcul nomm?e **Sample**. La courbe de tendance affiche une moyenne mobile sur 5 p?riodes.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-p108">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

> [!NOTE]
> <span data-ttu-id="cc3cc-145">Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta).</span><span class="sxs-lookup"><span data-stu-id="cc3cc-145">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="cc3cc-146">Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-146">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="cc3cc-147">**Graphique avec courbe de tendance de moyenne mobile**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-147">**Chart with moving average trendline**</span></span>

![Graphique avec courbe de tendance de moyenne mobile dans Excel](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="cc3cc-149">Mettre ? jour une courbe de tendance</span><span class="sxs-lookup"><span data-stu-id="cc3cc-149">Update a trendline</span></span>

<span data-ttu-id="cc3cc-150">L?exemple de code suivant d?finit la courbe de tendance sur le type **Linear** pour la premi?re s?rie du premier graphique de la feuille de calcul nomm?e **Sample**.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-150">The following code sample sets the trendline to type **Linear** for the first series in the first chart in the worksheet named **Sample**.</span></span>

> [!NOTE]
> <span data-ttu-id="cc3cc-151">Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta).</span><span class="sxs-lookup"><span data-stu-id="cc3cc-151">This sample uses APIs that are currently available only in public preview (beta).</span></span> <span data-ttu-id="cc3cc-152">Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="cc3cc-152">To run this sample, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>

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

<span data-ttu-id="cc3cc-153">**Graphique avec une courbe de tendance lin?aire**</span><span class="sxs-lookup"><span data-stu-id="cc3cc-153">**Chart with linear trendline**</span></span>

![Graphique avec une courbe de tendance lin?aire dans Excel](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a><span data-ttu-id="cc3cc-155">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="cc3cc-155">See also</span></span>

- [<span data-ttu-id="cc3cc-156">Concepts de base de l?API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="cc3cc-156">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="cc3cc-157">Objet Chart (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="cc3cc-157">Chart Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/chart) 
- [<span data-ttu-id="cc3cc-158">Objet ChartCollection (API JavaScript pour Excel)</span><span class="sxs-lookup"><span data-stu-id="cc3cc-158">Chart Collection Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/chartcollection)