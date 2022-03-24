---
title: Excel l’ensemble de conditions requises de l’API JavaScript 1.1
description: Détails sur l’ensemble de conditions requises ExcelApi 1.1.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 45061afc7e401e18a67377bf88fa1670bb7a8ece
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745955"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel l’ensemble de conditions requises de l’API JavaScript 1.1

L’API JavaScript 1.1 pour Excel est la première version de l’API. Il s’agit du seul ensemble Excel de conditions requises spécifiques pris en charge par Excel 2016.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.1. Pour afficher la documentation de référence de l’API pour toutes les API Excel l’ensemble de conditions requises de l’API JavaScript 1.1, voir Excel API dans l’ensemble de conditions requises [1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true).

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#excel-excel-application-calculate-member(1))|Recalcule tous les classeurs actuellement ouverts dans Excel.|
||[calculationMode](/javascript/api/excel/excel.application#excel-excel-application-calculationmode-member)|Renvoie le mode de calcul utilisé dans le workbook, tel que défini par les constantes dans `Excel.CalculationMode`.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#excel-excel-binding-getrange-member(1))|Renvoie la plage représentée par la liaison.|
||[getTable()](/javascript/api/excel/excel.binding#excel-excel-binding-gettable-member(1))|Renvoie le tableau représenté par la liaison.|
||[getText()](/javascript/api/excel/excel.binding#excel-excel-binding-gettext-member(1))|Renvoie le texte représenté par la liaison.|
||[id](/javascript/api/excel/excel.binding#excel-excel-binding-id-member)|Représente l’identificateur de liaison.|
||[type](/javascript/api/excel/excel.binding#excel-excel-binding-type-member)|Renvoie le type de la liaison.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[count](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-count-member)|Renvoie le nombre de liaisons de la collection.|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitem-member(1))|Obtient un objet de liaison par ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemat-member(1))|Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.|
||[items](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#excel-excel-chart-axes-member)|Représente les axes du graphique.|
||[dataLabels](/javascript/api/excel/excel.chart#excel-excel-chart-datalabels-member)|Représente les étiquettes des données sur le graphique.|
||[delete()](/javascript/api/excel/excel.chart#excel-excel-chart-delete-member(1))|Supprime l’objet de graphique.|
||[format](/javascript/api/excel/excel.chart#excel-excel-chart-format-member)|Regroupe les propriétés de format de la zone de graphique.|
||[height](/javascript/api/excel/excel.chart#excel-excel-chart-height-member)|Spécifie la hauteur, en points, de l’objet graphique.|
||[left](/javascript/api/excel/excel.chart#excel-excel-chart-left-member)|La distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[legend](/javascript/api/excel/excel.chart#excel-excel-chart-legend-member)|Représente la légende du graphique.|
||[name](/javascript/api/excel/excel.chart#excel-excel-chart-name-member)|Spécifie le nom d’un objet graphique.|
||[series](/javascript/api/excel/excel.chart#excel-excel-chart-series-member)|Représente une série ou une collection de séries dans le graphique.|
||[setData(sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#excel-excel-chart-setdata-member(1))|Redéfinit les données sources du graphique.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#excel-excel-chart-setposition-member(1))|Positionne le graphique par rapport aux cellules dans la feuille de calcul.|
||[title](/javascript/api/excel/excel.chart#excel-excel-chart-title-member)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre.|
||[top](/javascript/api/excel/excel.chart#excel-excel-chart-top-member)|Spécifie la distance, en points, entre le bord supérieur de l’objet et le haut de la ligne 1 (dans une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|
||[width](/javascript/api/excel/excel.chart#excel-excel-chart-width-member)|Spécifie la largeur, en points, de l’objet graphique.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-fill-member)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-font-member)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) de l’objet.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-categoryaxis-member)|Représente l’axe des abscisses d’un graphique.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-seriesaxis-member)|Représente l’axe des séries d’un graphique 3D.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-valueaxis-member)|Représente l’axe des ordonnées.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-format-member)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorgridlines-member)|Renvoie un objet qui représente le quadrillage principal de l’axe spécifié.|
||[majorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorunit-member)|Représente l’intervalle entre deux graduations principales.|
||[maximum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-maximum-member)|Représente la valeur maximale sur l’axe des ordonnées.|
||[minimum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minimum-member)|Représente la valeur minimale sur l’axe des ordonnées.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorgridlines-member)|Renvoie un objet qui représente le quadrillage mineur de l’axe spécifié.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorunit-member)|Représente l’intervalle entre deux graduations secondaires.|
||[title](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-title-member)|Représente le titre de l’axe.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[police](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-font-member)|Spécifie les attributs de police (nom de police, taille de police, couleur, etc.) d’un élément d’axe de graphique.|
||[line](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-line-member)|Spécifie la mise en forme des lignes de graphique.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-format-member)|Spécifie la mise en forme du titre de l’axe du graphique.|
||[text](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-text-member)|Spécifie le titre de l’axe.|
||[visible](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-visible-member)|Spécifie si le titre de l’axe est visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[police](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-font-member)|Spécifie les attributs de police du titre de l’axe du graphique, tels que le nom de la police, la taille de police ou la couleur, de l’objet de titre de l’axe du graphique.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel. ChartType, sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-add-member(1))|Crée un graphique.|
||[count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|Renvoie le nombre de graphiques dans la feuille de calcul.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitem-member(1))|Extrait un graphique à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemat-member(1))|Extrait un graphique en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-fill-member)|Représente le format de remplissage de l’étiquette de données.|
||[police](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-font-member)|Représente les attributs de police (tels que le nom de la police, la taille de police et la couleur) d’une étiquette de données de graphique.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-format-member)|Spécifie le format des étiquettes de données de graphique, qui inclut la mise en forme de remplissage et de police.|
||[position](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-position-member)|Valeur qui représente la position de l’étiquette de données.|
||[séparateur](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-separator-member)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showbubblesize-member)|Spécifie si la taille des bulles des étiquettes de données est visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showcategoryname-member)|Spécifie si le nom de catégorie d’étiquette de données est visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showlegendkey-member)|Spécifie si le clé de légende d’étiquette de données est visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showpercentage-member)|Spécifie si le pourcentage d’étiquette de données est visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showseriesname-member)|Spécifie si le nom de la série d’étiquettes de données est visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showvalue-member)|Spécifie si la valeur de l’étiquette de données est visible.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-clear-member(1))|Permet d’effacer la couleur de remplissage d’un élément de graphique.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-setsolidcolor-member(1))|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-bold-member)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-color-member)|Représentation de code couleur HTML de la couleur du texte (par exemple, #FF0000 représente le rouge).|
||[italic](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-italic-member)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-name-member)|Nom de la police (par exemple, « Calibri »)|
||[taille](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-size-member)|Taille de la police (par exemple, 11)|
||[underline](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-underline-member)|Type de soulignement appliqué à la police.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-format-member)|Représente le format du quadrillage de graphique.|
||[visible](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-visible-member)|Spécifie si le quadrillage de l’axe est visible.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#excel-excel-chartgridlinesformat-line-member)|Représente le format des lignes du graphique.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-format-member)|Représente le format d’une légende de graphique, à savoir le format du remplissage et de la police.|
||[overlay](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-overlay-member)|Spécifie si la légende du graphique doit chevaucher le corps principal du graphique.|
||[position](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-position-member)|Spécifie la position de la légende sur le graphique.|
||[visible](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-visible-member)|Spécifie si la légende du graphique est visible.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-fill-member)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-font-member)|Représente les attributs de police tels que le nom de police, la taille de police et la couleur d’une légende de graphique.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-clear-member(1))|Permet d’effacer le format de trait d’un élément de graphique.|
||[color](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-color-member)|Code couleur HTML qui représente la couleur des lignes dans le graphique.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-format-member)|Regroupe les propriétés de format d’un point d’un graphique.|
||[value](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-value-member)|Renvoie la valeur d’un point du graphique.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-fill-member)|Représente le format de remplissage d’un graphique, qui inclut des informations de mise en forme d’arrière-plan.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[count](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-count-member)|Renvoie le nombre de points de graphique dans la série.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getitemat-member(1))|Extrait un point en fonction de sa position dans la série.|
||[items](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[format](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-format-member)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes.|
||[name](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-name-member)|Spécifie le nom d’une série dans un graphique.|
||[points](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-points-member)|Renvoie une collection de tous les points de la série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[count](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-count-member)|Renvoie le nombre de séries de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getitemat-member(1))|Extrait une série en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-fill-member)|Représente le format de remplissage d’une série du graphique, qui comprend les informations de mise en forme d’arrière-plan.|
||[line](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-line-member)|Représente le format des lignes.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-format-member)|Représente le format du titre d’un graphique, à savoir le format de remplissage et de la police.|
||[overlay](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-overlay-member)|Spécifie si le titre du graphique se superpose au graphique.|
||[text](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-text-member)|Spécifie le texte du titre du graphique.|
||[visible](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-visible-member)|Spécifie si le titre du graphique est visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-fill-member)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-font-member)|Représente les attributs de police (tels que le nom de police, la taille et la couleur de police) d’un objet.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrange-member(1))|Renvoie l’objet de plage qui est associé au nom.|
||[name](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-name-member)|Nom de l’objet.|
||[type](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-type-member)|Spécifie le type de la valeur renvoyée par la formule du nom.|
||[value](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-value-member)|Représente la valeur calculée par la formule du nom.|
||[visible](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-visible-member)|Spécifie si l’objet est visible.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitem-member(1))|Obtient un `NamedItem` objet à l’aide de son nom.|
||[items](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[adresse](/javascript/api/excel/excel.range#excel-excel-range-address-member)|Spécifie la référence de plage dans le style A1.|
||[addressLocal](/javascript/api/excel/excel.range#excel-excel-range-addresslocal-member)|Représente la référence de plage pour la plage spécifiée dans la langue de l’utilisateur.|
||[cellCount](/javascript/api/excel/excel.range#excel-excel-range-cellcount-member)|Spécifie le nombre de cellules dans la plage.|
||[Effacer (applyTo ? : Excel.ClearApplyTo)](/javascript/api/excel/excel.range#excel-excel-range-clear-member(1))|Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.|
||[columnCount](/javascript/api/excel/excel.range#excel-excel-range-columncount-member)|Spécifie le nombre total de colonnes dans la plage.|
||[columnIndex](/javascript/api/excel/excel.range#excel-excel-range-columnindex-member)|Spécifie le numéro de colonne de la première cellule de la plage.|
||[delete(shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-delete-member(1))|Supprime les cellules associées à la plage.|
||[format](/javascript/api/excel/excel.range#excel-excel-range-format-member)|Renvoie un objet format, qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de la plage.|
||[formulas](/javascript/api/excel/excel.range#excel-excel-range-formulas-member)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.range#excel-excel-range-formulaslocal-member)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getboundingrect-member(1))|Renvoie le plus petit objet de plage qui englobe les plages données.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcell-member(1))|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcolumn-member(1))|Obtient une colonne contenue dans la plage.|
||[getEntireColumn()](/javascript/api/excel/excel.range#excel-excel-range-getentirecolumn-member(1))|Obtient un objet qui représente la colonne entière de la plage (par exemple, si la plage actuelle représente les cellules « B4:E11 `getEntireColumn` », il s’agit d’une plage qui représente les colonnes « B:E »).|
||[getEntireRow()](/javascript/api/excel/excel.range#excel-excel-range-getentirerow-member(1))|Obtient un objet qui représente la ligne entière de la plage (par exemple, si la plage actuelle représente les cellules « B4:E11 `GetEntireRow` », il s’agit d’une plage qui représente les lignes « 4:11 »).|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersection-member(1))|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.|
||[getLastCell()](/javascript/api/excel/excel.range#excel-excel-range-getlastcell-member(1))|Obtient la dernière cellule de la plage.|
||[getLastColumn()](/javascript/api/excel/excel.range#excel-excel-range-getlastcolumn-member(1))|Obtient la dernière colonne de la plage.|
||[getLastRow()](/javascript/api/excel/excel.range#excel-excel-range-getlastrow-member(1))|Obtient la dernière ligne de la plage.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1))|Obtient un objet qui représente une plage décalée par rapport à la plage spécifiée.|
||[getRow(row: number)](/javascript/api/excel/excel.range#excel-excel-range-getrow-member(1))|Obtient une ligne contenue dans la plage.|
||[insert(shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1))|Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace.|
||[numberFormat](/javascript/api/excel/excel.range#excel-excel-range-numberformat-member)|Représente le Excel de format numérique de la plage donnée.|
||[rowCount](/javascript/api/excel/excel.range#excel-excel-range-rowcount-member)|Renvoie le nombre total de lignes de la plage.|
||[rowIndex](/javascript/api/excel/excel.range#excel-excel-range-rowindex-member)|Renvoie le numéro de ligne de la première cellule de la plage.|
||[select()](/javascript/api/excel/excel.range#excel-excel-range-select-member(1))|Sélectionne la plage spécifiée dans l’interface utilisateur d’Excel.|
||[text](/javascript/api/excel/excel.range#excel-excel-range-text-member)|Valeurs de texte de la plage spécifiée.|
||[valueTypes](/javascript/api/excel/excel.range#excel-excel-range-valuetypes-member)|Spécifie le type de données dans chaque cellule.|
||[values](/javascript/api/excel/excel.range#excel-excel-range-values-member)|Représente les valeurs brutes de la plage spécifiée.|
||[worksheet](/javascript/api/excel/excel.range#excel-excel-range-worksheet-member)|Feuille de calcul contenant la plage.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-color-member)|Code couleur HTML représentant la couleur de la bordure, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-sideindex-member)|Valeur constante qui indique un côté spécifique de la bordure.|
||[style](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-style-member)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure.|
||[weight](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-weight-member)|Spécifie l'épaisseur de la bordure autour d'une plage.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[count](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-count-member)|Nombre d’objets de bordure de la collection.|
||[getItem(index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitem-member(1))|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitemat-member(1))|Obtient un objet de bordure à l’aide de son indice.|
||[items](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-clear-member(1))|Réinitialise l’arrière-plan de la plage.|
||[color](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-color-member)|Code couleur HTML représentant la couleur d’arrière-plan, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-bold-member)|Représente l’état gras de la police.|
||[color](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-color-member)|Représentation de code couleur HTML de la couleur du texte (par exemple, #FF0000 représente le rouge).|
||[italic](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-italic-member)|Spécifie l’état italique de la police.|
||[name](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-name-member)|Nom de la police (par exemple, « Calibri »).|
||[taille](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-size-member)|Taille de police|
||[underline](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-underline-member)|Type de soulignement appliqué à la police.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Borders](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-borders-member)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage.|
||[fill](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-fill-member)|Renvoie l’objet de remplissage défini sur la plage globale.|
||[police](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-font-member)|Renvoie l’objet de police défini sur l’ensemble de la plage.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-horizontalalignment-member)|Représente l’alignement horizontal de l’objet spécifié.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-verticalalignment-member)|Représente l’alignement vertical de l’objet spécifié.|
||[wrapText](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-wraptext-member)|Spécifie si Excel le texte dans l’objet.|
|[Table](/javascript/api/excel/excel.table)|[colonnes](/javascript/api/excel/excel.table#excel-excel-table-columns-member)|Représente une collection de toutes les colonnes du tableau.|
||[delete()](/javascript/api/excel/excel.table#excel-excel-table-delete-member(1))|Supprime le tableau.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1))|Obtient l’objet de plage associé au corps de données du tableau.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#excel-excel-table-getheaderrowrange-member(1))|Obtient l’objet de plage associé à la ligne d’en-tête du tableau.|
||[getRange()](/javascript/api/excel/excel.table#excel-excel-table-getrange-member(1))|Renvoie l’objet de plage associé à l’intégralité du tableau.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#excel-excel-table-gettotalrowrange-member(1))|Obtient l’objet de plage associé à la ligne de total du tableau.|
||[id](/javascript/api/excel/excel.table#excel-excel-table-id-member)|Renvoie une valeur qui permet d’identifier le tableau dans un classeur donné.|
||[name](/javascript/api/excel/excel.table#excel-excel-table-name-member)|Nom du tableau.|
||[rows](/javascript/api/excel/excel.table#excel-excel-table-rows-member)|Représente une collection de toutes les lignes du tableau.|
||[showHeaders](/javascript/api/excel/excel.table#excel-excel-table-showheaders-member)|Spécifie si la ligne d’en-tête est visible.|
||[showTotals](/javascript/api/excel/excel.table#excel-excel-table-showtotals-member)|Spécifie si la ligne de total est visible.|
||[style](/javascript/api/excel/excel.table#excel-excel-table-style-member)|Valeur constante qui représente le style de tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|Crée une table.|
||[count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|Renvoie le nombre de tableaux dans le classeur.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|Obtient un tableau à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|Obtient un tableau en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-delete-member(1))|Supprime la colonne du tableau.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getdatabodyrange-member(1))|Obtient l’objet de plage associé au corps de données de la colonne.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getheaderrowrange-member(1))|Obtient l’objet de plage associé à la ligne d’en-tête de la colonne.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getrange-member(1))|Renvoie l’objet de plage associé à l’intégralité de la colonne.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-gettotalrowrange-member(1))|Obtient l’objet de plage associé à la ligne de total de la colonne.|
||[id](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-id-member)|Renvoie une clé unique qui identifie la colonne du tableau.|
||[index](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-index-member)|Renvoie le numéro d’indice de la colonne dans la collection de colonnes du tableau.|
||[name](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-name-member)|Spécifie le nom de la colonne de tableau.|
||[values](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-values-member)|Représente les valeurs brutes de la plage spécifiée.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-add-member(1))|Ajoute une nouvelle colonne au tableau.|
||[count](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-count-member)|Renvoie le nombre de colonnes du tableau.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitem-member(1))|Obtient un objet de colonne par son nom ou son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemat-member(1))|Obtient une colonne en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-delete-member(1))|Supprime la ligne du tableau.|
||[getRange()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-getrange-member(1))|Renvoie l’objet de plage associé à la ligne entière.|
||[index](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-index-member)|Renvoie le numéro d’indice de la ligne dans la collection de lignes du tableau.|
||[values](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-values-member)|Représente les valeurs brutes de la plage spécifiée.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean \| string \| number>> \| boolean \| string \| number, alwaysInsert?: boolean)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))|Ajoute une ou plusieurs lignes dans le tableau.|
||[count](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-count-member)|Renvoie le nombre de lignes du tableau.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getitemat-member(1))|Obtient une ligne en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Workbook](/javascript/api/excel/excel.workbook)|[application](/javascript/api/excel/excel.workbook#excel-excel-workbook-application-member)|Représente l’instance Excel’application qui contient ce manuel.|
||[bindings](/javascript/api/excel/excel.workbook#excel-excel-workbook-bindings-member)|Représente une collection de liaisons appartenant au classeur.|
||[getSelectedRange()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1))|Obtient la plage unique actuellement sélectionnée à partir du manuel.|
||[names](/javascript/api/excel/excel.workbook#excel-excel-workbook-names-member)|Représente une collection d’éléments nommés d’étendue de workbook (plages et constantes nommées).|
||[tables](/javascript/api/excel/excel.workbook#excel-excel-workbook-tables-member)|Représente une collection de tableaux associés au classeur.|
||[feuilles de calcul](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member)|Représente une collection de feuilles de calcul associées au classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-activate-member(1))|Active la feuille de calcul dans l’interface utilisateur Excel.|
||[charts](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-charts-member)|Renvoie une collection de graphiques qui font partie de la feuille de calcul.|
||[delete()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-delete-member(1))|Supprime la feuille de calcul du classeur.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcell-member(1))|Obtient l’objet `Range` contenant la cellule unique en fonction des numéros de ligne et de colonne.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1))|Obtient l’objet `Range` , qui représente un seul bloc rectangulaire de cellules, spécifié par l’adresse ou le nom.|
||[id](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-id-member)|Renvoie une valeur qui permet d’identifier la feuille de calcul de façon unique dans un classeur donné.|
||[name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)|Nom complet de la feuille de calcul.|
||[position](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-position-member)|Position de la feuille de calcul au sein du classeur (sur une base zéro).|
||[tables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tables-member)|Collection de tableaux qui font partie de la feuille de calcul.|
||[visibility](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-visibility-member)|Visibilité de la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-add-member(1))|Ajoute une nouvelle feuille de calcul au classeur.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getactiveworksheet-member(1))|Obtient la feuille de calcul active du classeur.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitem-member(1))|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
