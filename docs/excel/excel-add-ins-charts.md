---
title: Utiliser des graphiques à l’aide de l’API JavaScript pour Excel
description: ''
ms.date: 03/11/2019
localization_priority: Priority
ms.openlocfilehash: f058110c7c150a75c847a07df83aa2795c891025
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600262"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a>Utiliser des graphiques à l’aide de l’API JavaScript pour Excel

This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API. For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart) and [Chart Collection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection).

## <a name="create-a-chart"></a>Création d’un graphique

The following code sample creates a chart in the worksheet named **Sample**. The chart is a **Line** chart that is based upon data in the range **A1:B13**.

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

**Nouveau graphique en courbes**

![Nouveau graphique en courbes dans Excel](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a>Ajouter une série de données à un graphique

The following code sample adds a data series to the first chart in the worksheet. The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.

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

**Graphique avant l’ajout de la série de données 2016**

![Graphique dans Excel avant l’ajout de la série de données 2016](../images/excel-charts-data-series-before.png)

**Graphique après l’ajout de la série de données 2016**

![Graphique dans Excel après l’ajout de la série de données 2016](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a>Définir le titre du graphique

L’exemple de code suivant définit le titre du premier graphique dans la feuille de calcul sur **Sales Data by Year**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique après la définition du titre**

![Graphique avec un titre dans Excel](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a>Définir les propriétés d’un axe d’un graphique

Les graphiques qui utilisent le [système de coordonnées cartésiennes](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), comme les histogrammes, les graphiques à barres et les nuages de points, ont un axe des abscisses et un axe des ordonnées. Ces exemples montrent comment définir le titre et afficher les unités d’un axe dans un graphique.

### <a name="set-axis-title"></a>Définir le titre d’un axe

L’exemple de code suivant définit le titre de l’axe des abscisses pour le premier graphique de la feuille de calcul sur **Product**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique après la définition de l’axe des abscisses**

![Graphique avec un titre d’axe dans Excel](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a>Définir l’unité d’affichage de l’axe

L’exemple de code suivant définit l’unité d’affichage de l’axe des ordonnées pour le premier graphique de la feuille de calcul sur **Hundreds**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique après la définition de l’unité d’affichage de l’axe des ordonnées**

![Graphique avec l’unité d’affichage de l’axe dans Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>Définir la visibilité du quadrillage dans un graphique

L’exemple de code suivant masque le quadrillage principal de l’axe des ordonnées du premier graphique de la feuille de calcul. Vous pouvez afficher le quadrillage principal de l’axe des ordonnées du graphique en définissant `chart.axes.valueAxis.majorGridlines.visible` sur **true**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique avec du quadrillage masqué**

![Graphique avec du quadrillage masqué dans Excel](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a>Courbes de tendance de graphiques

### <a name="add-a-trendline"></a>Ajouter une courbe de tendance

L’exemple de code suivant ajoute une courbe de tendance de moyenne mobile à la première série du premier graphique de la feuille de calcul nommée **Sample**. La courbe de tendance affiche une moyenne mobile sur 5 périodes.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique avec courbe de tendance de moyenne mobile**

![Graphique avec courbe de tendance de moyenne mobile dans Excel](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>Mettre à jour une courbe de tendance

L’exemple de code suivant définit la courbe de tendance sur le type **Linear** pour la première série du premier graphique de la feuille de calcul nommée **Sample**.

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

**Graphique avec une courbe de tendance linéaire**

![Graphique avec une courbe de tendance linéaire dans Excel](../images/excel-charts-trendline-linear.png)

## <a name="export-a-chart-as-an-image"></a>Exporter un graphique sous forme d’image

Vous pouvez générer des graphiques sous forme d’images en dehors d’Excel. `Chart.getImage` renvoie le graphique en tant que chaîne codée en Base64 représentant le graphique sous forme d’image JPEG. Le code suivant montre comment obtenir la chaîne de l’image et l’enregistrer dans la console.

```js
Excel.run(function (ctx) {
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    var imageAsString = chart.getImage();
    return context.sync().then(function () {
        console.log(imageAsString.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

`Chart.getImage` utilise trois paramètres facultatifs : largeur, hauteur et mode d’ajustement.

```typescript
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

Ces paramètres déterminent la taille de l’image. Les images sont toujours mises à l’échelle proportionnellement. Les paramètres de largeur et de hauteur appliquent des limites supérieures ou inférieures à l’image mise à l’échelle. `ImageFittingMode` contient trois valeurs avec les comportements suivants :

- `Fill` : la hauteur ou la largeur minimale de l’image est la hauteur ou la largeur spécifiée (dès que l’une ou l’autre est atteinte lorsque l’image est mise à l’échelle). Il s’agit du comportement par défaut lorsqu’aucun mode d’ajustement n’est spécifié.
- `Fit` : la hauteur ou la largeur maximale de l’image est la hauteur ou la largeur spécifiée (dès que l’une ou l’autre est atteinte lorsque l’image est mise à l’échelle).
- `FitAndCenter` : la hauteur ou la largeur maximale de l’image est la hauteur ou la largeur spécifiée (dès que l’une ou l’autre est atteinte lorsque l’image est mise à l’échelle). L’image générée est centrée par rapport à l’autre dimension.

## <a name="see-also"></a>Voir aussi

- [Concepts fondamentaux de programmation avec l’API JavaScript pour Excel](excel-add-ins-core-concepts.md)
