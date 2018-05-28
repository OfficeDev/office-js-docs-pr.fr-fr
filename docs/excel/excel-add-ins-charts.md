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
# <a name="work-with-charts-using-the-excel-javascript-api"></a>Utiliser des graphiques ? l?aide de l?API JavaScript pour Excel

Cet article fournit des exemples de code qui montrent comment effectuer des t?ches courantes ? l?aide de graphiques utilisant l?API JavaScript pour Excel. Pour une liste compl?te des propri?t?s et des m?thodes prises en charge par les objets **Chart** et **ChartCollection**, reportez-vous ? la rubrique [Objet Chart (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/chart) et [Objet ChartCollection (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection).

## <a name="create-a-chart"></a>Cr?er un graphique

L?exemple de code suivant cr?e un graphique dans la feuille de calcul nomm?e **Sample**. Il s?agit d?un graphique en **courbes** qui est fond? sur les donn?es de la plage **A1:B13**.

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


## <a name="add-a-data-series-to-a-chart"></a>Ajouter une s?rie de donn?es ? un graphique

L?exemple de code suivant ajoute une s?rie de donn?es au premier graphique de la feuille de calcul. La nouvelle s?rie de donn?es correspond ? la colonne nomm?e **2016** et repose sur les donn?es de la plage **D2:D5**.

> [!NOTE]
> Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta). Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

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

**Graphique avant l?ajout de la s?rie de donn?es 2016**

![Graphique dans Excel avant l?ajout de la s?rie de donn?es 2016](../images/excel-charts-data-series-before.png)

**Graphique apr?s l?ajout de la s?rie de donn?es 2016**

![Graphique dans Excel apr?s l?ajout de la s?rie de donn?es 2016](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a>D?finir le titre du graphique

L?exemple de code suivant d?finit le titre du premier graphique dans la feuille de calcul sur **Sales Data by Year**. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique apr?s la d?finition du titre**

![Graphique avec un titre dans Excel](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a>D?finir les propri?t?s d?un axe d?un graphique

Les graphiques qui utilisent le [syst?me de coordonn?es cart?siennes](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), comme les histogrammes, les graphiques ? barres et les nuages de points, ont un axe des abscisses et un axe des ordonn?es. Ces exemples montrent comment d?finir le titre et afficher les unit?s d?un axe dans un graphique.

### <a name="set-axis-title"></a>D?finir le titre d?un axe

L?exemple de code suivant d?finit le titre de l?axe des abscisses pour le premier graphique de la feuille de calcul sur **Product**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique apr?s la d?finition de l?axe des abscisses**

![Graphique avec un titre d?axe dans Excel](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a>D?finir l?unit? d?affichage de l?axe

L?exemple de code suivant d?finit l?unit? d?affichage de l?axe des ordonn?es pour le premier graphique de la feuille de calcul sur **Hundreds**.

> [!NOTE]
> Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta). Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique apr?s la d?finition de l?unit? d?affichage de l?axe des ordonn?es**

![Graphique avec l?unit? d?affichage de l?axe dans Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>D?finir la visibilit? du quadrillage dans un graphique

L?exemple de code suivant masque le quadrillage principal de l?axe des ordonn?es du premier graphique de la feuille de calcul. Vous pouvez afficher le quadrillage principal de l?axe des ordonn?es du graphique en d?finissant `chart.axes.valueAxis.majorGridlines.visible` sur **true**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Graphique avec du quadrillage masqu?**

![Graphique avec du quadrillage masqu? dans Excel](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a>Courbes de tendance de graphiques

### <a name="add-a-trendline"></a>Ajouter une courbe de tendance

L?exemple de code suivant ajoute une courbe de tendance de moyenne mobile ? la premi?re s?rie du premier graphique de la feuille de calcul nomm?e **Sample**. La courbe de tendance affiche une moyenne mobile sur 5 p?riodes.

> [!NOTE]
> Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta). Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

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

### <a name="update-a-trendline"></a>Mettre ? jour une courbe de tendance

L?exemple de code suivant d?finit la courbe de tendance sur le type **Linear** pour la premi?re s?rie du premier graphique de la feuille de calcul nomm?e **Sample**.

> [!NOTE]
> Cet exemple utilise des API qui ne sont actuellement disponibles qu'en pr?version publique (b?ta). Pour ex?cuter cet exemple de code, vous devez utiliser la biblioth?que b?ta du CDN Office.js?: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

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

**Graphique avec une courbe de tendance lin?aire**

![Graphique avec une courbe de tendance lin?aire dans Excel](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a>Voir aussi

- [Concepts de base de l?API JavaScript pour Excel](excel-add-ins-core-concepts.md)
- [Objet Chart (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/chart) 
- [Objet ChartCollection (API JavaScript pour Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection)