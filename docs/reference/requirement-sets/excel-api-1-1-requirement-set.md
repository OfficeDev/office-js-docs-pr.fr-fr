---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,1
description: Détails sur l’ensemble de conditions requises ExcelApi 1,1
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 90d7ee7cef2e8c48e458b2e14893ba9c13c68a30
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940786"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Ensemble de conditions requises de l’API JavaScript pour Excel 1,1

L’API JavaScript 1.1 pour Excel est la première version de l’API. Il s’agit du seul ensemble de conditions requises spécifiques à Excel pris en charge par Excel 2016.

## <a name="api-list"></a>Liste des API

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[Calculate (calculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalcule tous les classeurs actuellement ouverts dans Excel.|
||[calculationMode](/javascript/api/excel/excel.application#calculationmode)|Renvoie le mode de calcul utilisé dans le classeur, tel que défini par les constantes dans Excel. CalculationMode. Les valeurs possibles sont `Automatic`les suivantes:, où Excel contrôle le recalcul; `AutomaticExceptTables`, où Excel contrôle le recalcul, mais ignore les modifications apportées aux tableaux; `Manual`, où le calcul est effectué lorsque l’utilisateur le demande.|
|[Binding](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Renvoie la plage représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Renvoie le tableau représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Renvoie le texte représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
||[id](/javascript/api/excel/excel.binding#id)|Représente l’identificateur de liaison. En lecture seule.|
||[type](/javascript/api/excel/excel.binding#type)|Renvoie le type de la liaison. Pour plus d’informations, voir Excel. BindingType. En lecture seule.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Obtient un objet de liaison par ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.|
||[count](/javascript/api/excel/excel.bindingcollection#count)|Renvoie le nombre de liaisons de la collection. En lecture seule.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Supprime l’objet de graphique.|
||[height](/javascript/api/excel/excel.chart#height)|Représente la hauteur, exprimée en points, de l’objet de graphique.|
||[left](/javascript/api/excel/excel.chart#left)|Distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.chart#name)|Représente le nom d’un objet de graphique.|
||[ordonné](/javascript/api/excel/excel.chart#axes)|Représente les axes du graphique. En lecture seule.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Représente les étiquettes des données sur le graphique. En lecture seule.|
||[format](/javascript/api/excel/excel.chart#format)|Regroupe les propriétés de format de la zone de graphique. En lecture seule.|
||[Legend](/javascript/api/excel/excel.chart#legend)|Représente la légende du graphique. En lecture seule.|
||[série](/javascript/api/excel/excel.chart#series)|Représente une série ou une collection de séries dans le graphique. En lecture seule.|
||[title](/javascript/api/excel/excel.chart#title)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre. En lecture seule.|
||[setData (sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Redéfinit les données sources du graphique.|
||[setPosition (startCell: chaîne \| de plage, endCell?: \| chaîne de plage)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Positionne le graphique par rapport aux cellules dans la feuille de calcul.|
||[top](/javascript/api/excel/excel.chart#top)|Représente la distance, en points, entre le bord supérieur de l’objet et la partie supérieure de la ligne 1 (sur une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|
||[width](/javascript/api/excel/excel.chart#width)|Représente la largeur, en points, de l’objet de graphique.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|
||[police](/javascript/api/excel/excel.chartareaformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) de l’objet. En lecture seule.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Représente l’axe des abscisses d’un graphique. En lecture seule.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Représente l’axe de séries d’un graphique 3D. En lecture seule.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Représente l’axe des ordonnées. En lecture seule.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Représente l’intervalle entre deux graduations principales. Peut être défini sur une valeur numérique ou une chaîne vide.  La valeur renvoyée est toujours un nombre.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Représente la valeur maximale sur l’axe des ordonnées.  Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique).  La valeur renvoyée est toujours un nombre.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Représente la valeur minimale sur l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Représente l’intervalle entre deux graduations secondaires. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police. En lecture seule.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage principal de l’axe spécifié. En lecture seule.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage secondaire de l’axe spécifié. En lecture seule.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Représente le titre de l’axe. En lecture seule.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[police](/javascript/api/excel/excel.chartaxisformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’un élément d’axe de graphique. En lecture seule.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Représente le format des lignes du graphique. En lecture seule.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Représente le format du titre d’un axe de graphique. En lecture seule.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Représente le titre de l’axe.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Valeur booléenne qui spécifie la visibilité d’un titre d’axe.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[police](/javascript/api/excel/excel.chartaxistitleformat#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de titre d’axe de graphique. En lecture seule.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[Add (type: Excel. ChartType, sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Crée un graphique.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Extrait un graphique en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.chartcollection#count)|Renvoie le nombre de graphiques dans la feuille de calcul. En lecture seule.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Représente le format de remplissage de l’étiquette de données. En lecture seule.|
||[police](/javascript/api/excel/excel.chartdatalabelformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de données de graphique. En lecture seule.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Représente le format des étiquettes de données du graphique, à savoir le format de remplissage et de la police. En lecture seule.|
||[para](/javascript/api/excel/excel.chartdatalabels#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Supprime la couleur de remplissage d’un élément de graphique.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.chartfont#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.chartfont#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.chartfont#size)|Taille de la police (par exemple, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. ChartUnderlineStyle.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Représente le format du quadrillage de graphique. En lecture seule.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Valeur booléenne indiquant si les lignes de quadrillage de l’axe sont visibles ou non.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Représente le format des lignes du graphique. En lecture seule.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Valeur booléenne indiquant si la légende du graphique doit chevaucher le corps principal du graphique.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Représente la position de la légende sur le graphique. Pour plus d’informations, voir Excel. ChartLegendPosition.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Représente le format d’une légende de graphique, à savoir le format du remplissage et de la police. En lecture seule.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Valeur booléenne qui représente la visibilité d’un objet ChartLegend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|
||[police](/javascript/api/excel/excel.chartlegendformat#font)|Représente les attributs de police, tels que le nom de police, la taille de police, la couleur, etc., d’une légende de graphique. En lecture seule.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Désactiver le format de ligne d’un élément de graphique.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|Code couleur HTML qui représente la couleur des lignes dans le graphique.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Regroupe les propriétés de format d’un point d’un graphique. En lecture seule.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Renvoie la valeur d’un point du graphique. En lecture seule.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Représente le format de remplissage d’un graphique, qui inclut des informations de mise en forme de l’arrière-plan. En lecture seule.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Extrait un point en fonction de sa position dans la série.|
||[count](/javascript/api/excel/excel.chartpointscollection#count)|Renvoie le nombre de points de graphique dans la série. En lecture seule.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Représente le nom d’une série dans un graphique.|
||[format](/javascript/api/excel/excel.chartseries#format)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes. En lecture seule.|
||[pointe](/javascript/api/excel/excel.chartseries#points)|Représente la collection de tous les points de la série. En lecture seule.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Extrait une série en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.chartseriescollection#count)|Renvoie le nombre de séries de la collection. En lecture seule.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Représente le format de remplissage d’une série du graphique, qui comprend les informations de mise en forme d’arrière-plan. En lecture seule.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Représente le format des lignes. En lecture seule.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Valeur booléenne indiquant si le titre du graphique recouvre le graphique ou non.|
||[format](/javascript/api/excel/excel.charttitle#format)|Représente le format du titre d’un graphique, à savoir le format de remplissage et de la police. En lecture seule.|
||[text](/javascript/api/excel/excel.charttitle#text)|Représente le texte du titre d’un graphique.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Valeur booléenne qui représente la visibilité d’un objet de titre de graphique.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|
||[police](/javascript/api/excel/excel.charttitleformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) pour un objet. En lecture seule.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Renvoie l’objet de plage qui est associé au nom. Renvoie une erreur si le type de l’élément nommé n’est pas une plage.|
||[name](/javascript/api/excel/excel.nameditem#name)|Nom de l’objet. En lecture seule.|
||[type](/javascript/api/excel/excel.nameditem#type)|Indique le type de la valeur renvoyée par la formule du nom. Pour plus d’informations, voir Excel. NamedItemType. En lecture seule.|
||[value](/javascript/api/excel/excel.nameditem#value)|Représente la valeur calculée par la formule du nom. Pour une plage nommée, renvoie l’adresse de la plage. En lecture seule.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Indique si l’objet est visible ou non.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Obtient un objet NamedItem à l’aide de son nom.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[Effacer (applyTo ? : Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.|
||[supprimer (Maj: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Supprime les cellules associées à la plage.|
||[formulas](/javascript/api/excel/excel.range#formulas)|Représente la formule dans le style de notation A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
||[getBoundingRect (anotherRange: chaîne \| de plage)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Renvoie le plus petit objet de plage qui englobe les plages données. Par exemple, la valeur GetBoundingRect pour « B2:C5 » et « D10:E15 » est « B2:E15 ».|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut être située en dehors des limites de sa plage parente, tant qu’elle reste dans la grille de la feuille de calcul. L’emplacement de la cellule renvoyée est déterminé à partir de la cellule supérieure gauche de la plage.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Obtient une colonne contenue dans la plage.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Obtient un objet qui représente la colonne entière de la plage (par exemple, si la plage actuelle représente les cellules «B4: E11», `getEntireColumn` qu’il s’agit d’une plage qui représente les colonnes «B:E»).|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Obtient un objet qui représente la ligne entière de la plage (par exemple, si la plage actuelle représente les cellules «B4: E11», `GetEntireRow` qu’il s’agit d’une plage qui représente les lignes «4:11»).|
||[getIntersection (anotherRange: chaîne \| de plage)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Obtient la dernière cellule de la plage. Par exemple, la dernière cellule de la plage « B2:D5 » est « D5 ».|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|Obtient la dernière colonne de la plage. Par exemple, la dernière colonne de la plage « B2:D5 » est « D2:D5 ».|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|Obtient la dernière ligne de la plage. Par exemple, la dernière ligne de la plage « B2:D5 » est « B5:D5 ».|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Obtient un objet qui représente une plage décalée par rapport à la plage spécifiée. Les dimensions de la plage renvoyée correspondent à cette plage. Si la plage obtenue se retrouve en dehors des limites de grille de la feuille de calcul, une erreur est déclenchée.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Obtient une ligne contenue dans la plage.|
||[Insérer (Maj: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace. Renvoie un nouvel objet Range dans l’espace vide qui s’est créé.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Représente le code de format de nombre d’Excel pour la plage donnée.|
||[adresse](/javascript/api/excel/excel.range#address)|Représente la référence de plage dans le style a1. La valeur de l’adresse contiendra la référence de la feuille (par exemple, «Sheet1! A1: B4 "). En lecture seule.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Représente la référence de la plage spécifiée dans le langage de l’utilisateur. En lecture seule.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Nombre de cellules dans la plage. Cette API renvoie -1 si le nombre de cellules est supérieur à 2^31-1 (2 147 483 647). En lecture seule.|
||[NbreColonnes](/javascript/api/excel/excel.range#columncount)|Représente le nombre total de colonnes dans la plage. En lecture seule.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Représente le numéro de colonne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[format](/javascript/api/excel/excel.range#format)|Renvoie un objet format qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de la plage. En lecture seule.|
||[Stopp](/javascript/api/excel/excel.range#rowcount)|Renvoie le nombre total de lignes de la plage. En lecture seule.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Renvoie le numéro de ligne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
||[text](/javascript/api/excel/excel.range#text)|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Représente le type de données de chaque cellule. En lecture seule.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|Feuille de calcul contenant la plage. En lecture seule.|
||[select()](/javascript/api/excel/excel.range#select--)|Sélectionne la plage spécifiée dans l’interface utilisateur d’Excel.|
||[values](/javascript/api/excel/excel.range#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Valeur constante qui indique un côté spécifique de la bordure. Pour plus d’informations, voir Excel. BorderIndex. En lecture seule.|
||[style](/javascript/api/excel/excel.rangeborder#style)|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Pour plus d’informations, voir Excel. BorderLineStyle.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Spécifie l'épaisseur de la bordure autour d'une plage. Pour plus d’informations, voir Excel. BorderWeight.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Obtient un objet de bordure à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Obtient un objet de bordure à l’aide de son indice.|
||[count](/javascript/api/excel/excel.rangebordercollection#count)|Nombre d’objets de bordure de la collection. En lecture seule.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Réinitialise l’arrière-plan de la plage.|
||[color](/javascript/api/excel/excel.rangefill#color)|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.rangefont#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.rangefont#name)|Nom de la police (par exemple « Calibri »)|
||[size](/javascript/api/excel/excel.rangefont#size)|Taille de police|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Type de soulignement appliqué à la police. Pour plus d’informations, voir Excel. RangeUnderlineStyle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Représente l’alignement horizontal de l’objet spécifié. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[Borders](/javascript/api/excel/excel.rangeformat#borders)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage. En lecture seule.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Renvoie l’objet de remplissage défini sur la plage globale. En lecture seule.|
||[police](/javascript/api/excel/excel.rangeformat#font)|Renvoie l’objet de police défini sur l’ensemble de la plage. En lecture seule.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Représente l’alignement vertical de l’objet spécifié. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Indique si Excel renvoie le texte à la ligne dans l’objet. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Supprime le tableau.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Obtient l’objet de plage associé au corps de données du tableau.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Obtient l’objet de plage associé à la ligne d’en-tête du tableau.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Renvoie l’objet de plage associé à l’intégralité du tableau.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Renvoie l’objet de plage associé à la ligne de total du tableau.|
||[name](/javascript/api/excel/excel.table#name)|Nom du tableau.|
||[colonnes](/javascript/api/excel/excel.table#columns)|Représente une collection de toutes les colonnes du tableau. En lecture seule.|
||[id](/javascript/api/excel/excel.table#id)|Renvoie une valeur qui identifie le tableau dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque le tableau est renommé. En lecture seule.|
||[rows](/javascript/api/excel/excel.table#rows)|Représente une collection de toutes les lignes du tableau. En lecture seule.|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|Indique si la ligne d’en-tête est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne d’en-tête.|
||[ShowTotals,](/javascript/api/excel/excel.table#showtotals)|Indique si la ligne de total est visible ou non. Cette valeur peut être définie de manière à afficher ou à masquer la ligne de total.|
||[style](/javascript/api/excel/excel.table#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont les suivantes : TableStyleLight1 à TableStyleLight21, TableStyleMedium1 à TableStyleMedium28, TableStyleStyleDark1 à TableStyleStyleDark11. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[Add (Address: Range \| String, hasHeaders: Boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Crée un tableau L’adresse de la source ou de l’objet de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Obtient un tableau à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Obtient un tableau en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.tablecollection#count)|Renvoie le nombre de tableaux dans le classeur. En lecture seule.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Supprime la colonne du tableau.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Obtient l’objet de plage associé au corps de données de la colonne.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Obtient l’objet de plage associé à la ligne d’en-tête de la colonne.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Renvoie l’objet de plage associé à l’intégralité de la colonne.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Obtient l’objet de plage associé à la ligne de total de la colonne.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Représente le nom de la colonne du tableau.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Renvoie une clé unique qui identifie la colonne du tableau. En lecture seule.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Renvoie le numéro d’indice de la colonne dans la collection de colonnes du tableau. Avec indice zéro. En lecture seule.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[Add (index?: Number, Values?: Array<Array<\| Boolean \| String Number \|>> \| Boolean \| String Number, Name?: String)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Ajoute une nouvelle colonne au tableau.|
||[getItem (Key: valeur \| numérique)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Obtient un objet de colonne par son nom ou son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Obtient une colonne en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.tablecolumncollection#count)|Renvoie le nombre de colonnes du tableau. En lecture seule.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Supprime la ligne du tableau.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Renvoie l’objet de plage associé à la ligne entière.|
||[index](/javascript/api/excel/excel.tablerow#index)|Renvoie le numéro d’indice de la ligne dans la collection de lignes du tableau. Avec indice zéro. En lecture seule.|
||[values](/javascript/api/excel/excel.tablerow#values)|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[Add (index?: Number, Values?: Array<Array<\| Boolean \| String Number \|>> \| Boolean \| String Number)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Ajoute une ou plusieurs lignes dans le tableau. L’objet renvoyé sera placé en premier dans les lignes récemment ajoutées.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Obtient une ligne en fonction de sa position dans la collection.|
||[count](/javascript/api/excel/excel.tablerowcollection#count)|Renvoie le nombre de lignes du tableau. En lecture seule.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange ()](/javascript/api/excel/excel.workbook#getselectedrange--)|Obtient la plage unique actuellement sélectionnée du classeur. Si plusieurs plages sont sélectionnées, cette méthode génère une erreur.|
||[application](/javascript/api/excel/excel.workbook#application)|Représente l’instance de l’application Excel qui contient ce classeur. En lecture seule.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Représente une collection de liaisons appartenant au classeur. En lecture seule.|
||[noms](/javascript/api/excel/excel.workbook#names)|Représente une collection d’éléments nommés portant sur le classeur (appelés plages et constantes). En lecture seule.|
||[emplois](/javascript/api/excel/excel.workbook#tables)|Représente une collection de tableaux associés au classeur. En lecture seule.|
||[feuilles](/javascript/api/excel/excel.workbook#worksheets)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Active la feuille de calcul dans l’interface utilisateur Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Supprime la feuille de calcul du classeur. Notez que si la visibilité de la feuille de calcul est définie sur «VeryHidden», l’opération de suppression échouera avec un GeneralException.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut être située en dehors des limites de sa plage parente, tant qu’elle reste dans la grille de la feuille de calcul.|
||[getRange (Address?: String)](/javascript/api/excel/excel.worksheet#getrange-address-)|Obtient l’objet de plage, représentant un seul bloc de cellules rectangulaires, spécifié par l’adresse ou le nom.|
||[name](/javascript/api/excel/excel.worksheet#name)|Nom complet de la feuille de calcul.|
||[position](/javascript/api/excel/excel.worksheet#position)|Position de la feuille de calcul au sein du classeur (sur une base zéro).|
||[bulles](/javascript/api/excel/excel.worksheet#charts)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
||[id](/javascript/api/excel/excel.worksheet#id)|Renvoie une valeur qui permet d’identifier la feuille de calcul de façon unique dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque la feuille de calcul est renommée ou déplacée. En lecture seule.|
||[emplois](/javascript/api/excel/excel.worksheet#tables)|Collection de tableaux qui font partie de la feuille de calcul. En lecture seule.|
||[excellente](/javascript/api/excel/excel.worksheet#visibility)|Visibilité de la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[Add (Name?: String)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Ajoute une nouvelle feuille de calcul au classeur. La feuille de calcul est ajoutée à la fin des feuilles de calcul existantes. Si vous souhaitez activer la feuille de calcul nouvellement ajoutée, appelez la méthode .activate() pour cette feuille.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Obtient la feuille de calcul active du classeur.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
