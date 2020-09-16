---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,7
description: Détails sur l’ensemble de conditions requises ExcelApi 1,7.
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad0b1a205191ae5fd2b68b933cdf3bb757ecbd2b
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819650"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Nouveautés de l’API JavaScript 1.7 pour Excel

Les fonctionnalités Excel JavaScript API ensemble de conditions 1.7 incluent des API pour les graphiques, événements, feuilles de calcul, plages, propriétés de document, éléments nommés, options de protection et styles.

## <a name="customize-charts"></a>Personnaliser des graphiques

Avec le nouvel API graphique, vous pouvez créer des types de graphiques supplémentaires, ajouter une série de données à un graphique, définir le titre du graphique, ajouter un titre d’axe, ajouter une unité d’affichage, ajouter une courbe de tendance avec moyenne mobile, modifier une courbe de tendance en ligne, et bien plus encore. Voici quelques exemples :

* Axe du graphique - obtenir, définir, mettre en forme et supprimer une unité d’axe, une étiquette et un titre dans un graphique.
* Série de graphique - ajouter, configurer et supprimer une série dans un graphique.  Modifier les marqueurs de série, les commandes traçage et le redimensionnement.
* Courbes de tendance de graphique - ajouter, obtenir et mettre en forme des courbes de tendance dans un graphique.
* Légende de graphique - mettre en forme la police de légende dans un graphique.
* Point de graphique - définir la couleur du point de graphique.
* Sous-chaîne de titre du graphique - obtenir et définir une sous-chaîne de titre d’un graphique.
* Type de graphique - option pour créer plusieurs types de graphiques.

## <a name="events"></a>Événements

Les API Événements pour Excel fournissent un grand nombre de gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. Pour une liste des événements qui sont actuellement disponibles, voir [Manipuler des Événements à l’aide de l’API JavaScript Excel](../../excel/excel-add-ins-events.md).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personnaliser l’apparence de feuilles de calcul et des plages

À l’aide des nouveaux API, vous pouvez personnaliser l’apparence de feuilles de calcul de plusieurs façons :

* Figer les volets pour conserver certaines lignes ou colonnes visibles lorsque vous faites défiler la feuille de calcul. Par exemple, si la première ligne dans votre feuille de calcul contient des en-têtes, vous pouvez figer cette ligne de sorte que les en-têtes de colonne restent visibles pendant le défilement vers le bas de la feuille de calcul.
* Modifier la couleur d’onglet de la feuille de calcul.
* Ajouter des en-têtes de feuille de calcul.

Vous pouvez personnaliser l’apparence des plages de plusieurs façons :

* Définir le style de cellule pour une plage pour vous assurer que toutes les cellules dans la plage ont une mise en forme cohérente. Un style de cellule est un ensemble défini de caractéristiques de mise en forme, comme les polices et les tailles de police, formats des nombres, bordures de cellule et ombrage de cellule. Utilisez un des styles de cellule intégrés d’Excel ou créer votre propre style de cellule personnalisé.
* Définit l’orientation du texte pour une plage.
* Ajouter ou modifier un lien hypertexte sur une plage qui permet d’accéder à un autre emplacement dans le classeur ou à un emplacement externe.

## <a name="manage-document-properties"></a>Gérer les propriétés du document

À l’aide des API de propriétés du document, vous pouvez accéder aux propriétés de document intégrées et également créer et gérer les propriétés de document personnalisées pour stocker l’état du classeur et lire le flux de travail et la logique d’entreprise.

## <a name="copy-worksheets"></a>Obtenir des feuilles de calcul

À l’aide des API de copie de feuille de calcul , vous pouvez copier les données et le format à partir d’une feuille de calcul dans une nouvelle feuille de calcul au sein du même classeur et réduire la quantité de transfert de données nécessaire.

## <a name="handle-ranges-with-ease"></a>Gérer les plages en toute simplicité

À l’aide des API de plage différente, vous pouvez effectuer des actions telles qu’obtenir la région environnante, obtenir une plage redimensionnée et bien plus encore. Ces API doivent rendre des tâches telles que la manipulation de plage et l’adressage beaucoup plus efficaces.

De plus :

* Options de protection de classeur et feuille de calcul : utilisez ces API pour protéger les données dans une feuille de calcul et la structure du classeur.
* Mettre à jour un élément nommé : utilisez cet API pour mettre à jour un élément nommé.
* Obtenir la cellule active : utilisez cet API pour obtenir la cellule active d’un classeur.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,7. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,7 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,7 ou version antérieure](/javascript/api/excel?view=excel-js-1.7&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[Graphique](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Représente le type d’un graphique. Pour plus d’informations, voir Excel. ChartType.|
||[id](/javascript/api/excel/excel.chart#id)|ID unique du graphique. En lecture seule.|
||[ShowAllFieldButtons,](/javascript/api/excel/excel.chart#showallfieldbuttons)|Représente l’affichage de tous les boutons de champ dans un graphique croisé dynamique.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[route](/javascript/api/excel/excel.chartareaformat#border)|Représente le format de bordure de la zone de graphique, qui inclut couleur, style de style et épaisseur. En lecture seule.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (type : Excel. ChartAxisType, Group ?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Renvoie l’axe spécifique identifié par type et par groupe.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Renvoie ou définit l’unité de base pour l’axe des abscisses spécifié.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Renvoie ou définit le type d’axe de catégorie.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Représente l’unité d’affichage de l’axe. Pour plus d’informations, voir Excel. ChartAxisDisplayUnit.|
||[LogBase,](/javascript/api/excel/excel.chartaxis#logbase)|Représente la base du logarithme lorsque vous utilisez des échelles logarithmiques.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Représente le type de graduation principale pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité principale pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Représente le type de graduation secondaire pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité secondaire pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Représente le groupe pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisGroup. En lecture seule.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Représente la valeur unité d’affichage personnalisé d’axe. En lecture seule. Pour définir cette propriété, utilisez la méthode SetCustomDisplayUnit(double).|
||[height](/javascript/api/excel/excel.chartaxis#height)|Représente la hauteur, exprimée en points, de l’axe de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Représente la distance en points, du bord gauche de l’axe au bord gauche de la zone de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Représente la distance en points, du bord supérieur de l’axe au bord supérieur de la zone de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Représente le type d’axe. Pour plus d’informations, voir Excel. ChartAxisType.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Représente la largeur, en points, de l’axe de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Représente si Microsoft Excel trace des points de données à du dernier au premier.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Représente le type d’échelle de l’axe des ordonnées. Pour plus d’informations, voir Excel. ChartAxisScaleType.|
||[setCategoryNames (sourceData : Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Définit tous les noms de catégorie pour l’axe spécifié.|
||[setCustomDisplayUnit (valeur : nombre)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Définit l’unité d’affichage axe sur une valeur personnalisée.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Représente si l’étiquette d’unité de l’affichage d’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Représentant la position des étiquettes de graduation sur l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickLabelPosition.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Représente le nombre de catégories ou séries entre les étiquettes de graduation. Peut être une valeur de 1 à 31999 ou une chaîne vide pour le paramètre automatique. La valeur renvoyée est toujours un nombre.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Représente le nombre de catégories ou séries entre les graduations.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Valeur booléenne qui représente la visibilité d’un axe.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Désactiver le format de bordure d’un élément de graphique.|
||[color](/javascript/api/excel/excel.chartborder#color)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Représente le style de trait de la bordure. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabel#autotext)|Valeur booléenne représentant si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|
||[position](/javascript/api/excel/excel.chartdatalabel#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Représente le format d’étiquette de données graphique.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Représente la hauteur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Représente la largeur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
||[para](/javascript/api/excel/excel.chartdatalabel#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Chaîne représentant le texte d’étiquette de données dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Représente l’orientation du texte de l’étiquette de données du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[police](/javascript/api/excel/excel.chartformatstring#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de graphique.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Représente la hauteur, en points, de la légende du graphique. NULL si la légende n’est pas visible.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Représente la gauche, en points, d’une légende de graphique. NULL si la légende n’est pas visible.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Représente une collection de legendEntries dans la légende. En lecture seule.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Représente si la légende est ombrée sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Représente la partie supérieure d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Représente la largeur, exprimée en points, de la légende du graphique. NULL si la légende n’est pas visible.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Représente la hauteur de legendEntry sur la légende du graphique.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Représente l’index de legendEntry sur la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Représente la partie gauche d’un graphique legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Représente la partie supérieure d’un graphique legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Représente la largeur de legendEntry sur la légende d’un graphique.|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Représente la partie visible d’une entrée de légende de graphique.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Renvoie le nombre de legendEntry de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Renvoie un legendEntry à l’index donné.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Représente le style de trait. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Indique si un point de données a une étiquette de données. Non applicable pour les graphiques en surface.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Représentation de code de couleur HTML de la couleur de premier plan du point de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerSize,](/javascript/api/excel/excel.chartpoint#markersize)|Représente la taille du marqueur du point de données.|
||[MarkerStyle,](/javascript/api/excel/excel.chartpoint#markerstyle)|Représente le style du marqueur du point de données de graphique. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Renvoie l’étiquette de données d’un point du graphique. En lecture seule.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[route](/javascript/api/excel/excel.chartpointformat#border)|Représente le format de bordure d’un point de données de graphique, qui inclut les informations de couleur, de style et de poids. En lecture seule.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Représente le type de graphique d’une série. Pour plus d’informations, voir Excel. ChartType.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Supprime la série graphique.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Représente la taille du centre d’une série de graphiques en anneaux.  Valide uniquement dans les graphiques en anneaux et doughnutExploded.|
||[filtré](/javascript/api/excel/excel.chartseries#filtered)|Valeur booléenne représentant si la série est filtrée ou non. Non applicable pour les graphiques en surface.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Représente la largeur de l’intervalle d’une série de graphique.  Valide uniquement sur les graphiques en barres et colonnes, ainsi que|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Valeur booléenne représentant si la série a des étiquettes de données ou non.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Représente la couleur d’arrière-plan de marqueurs d’une série de graphiques.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Représente la couleur de premier plan de marqueurs d’une série de graphiques.|
||[MarkerSize,](/javascript/api/excel/excel.chartseries#markersize)|Représente la taille du marqueur d’une série de graphiques.|
||[MarkerStyle,](/javascript/api/excel/excel.chartseries#markerstyle)|Représente le style du marqueur d’une série de graphiques. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
||[PlotOrder,](/javascript/api/excel/excel.chartseries#plotorder)|Représente l’ordre de traçage d’une série de graphiques au sein du groupe de graphiques.|
||[Trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Représente la collection de courbes de tendance de la série. En lecture seule.|
||[setBubbleSizes (sourceData : Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Définit des tailles de bulles pour une série de graphiques. Fonctionne uniquement pour les graphiques en bulles.|
||[SetValues (sourceData : Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Définit des valeurs pour une série de graphiques. Pour un graphique en nuages de points, cela signifie des valeurs de l’axe Y.|
||[setXAxisValues (sourceData : Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Définit des valeurs d’axe Y pour une série de graphiques. Fonctionne uniquement pour les graphiques en nuages de points.|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|Valeur booléenne représentant si la série a une ombre ou non.|
||[Unie](/javascript/api/excel/excel.chartseries#smooth)|Valeur booléenne représentant si la série est fluide ou non. Applicable uniquement aux graphiques en courbes et à nuages de points.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Add (Name ?: String, index ?: Number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Ajouter une nouvelle série à la collection. La nouvelle série ajoutée n’est pas visible jusqu’à ce que les valeurs de l’axe des ordonnées/valeurs de l’axe x/tailles des bulles lui soient attribuées (en fonction du type de graphique).|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (début : nombre, longueur : nombre)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Obtenir la sous-chaîne d’un titre de graphique. Le saut de ligne « \n » compte également un seul caractère.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Représente l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitle#left)|Représente la distance en points, du bord gauche du titre du graphique au bord gauche de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[position](/javascript/api/excel/excel.charttitle#position)|Représente la position du titre du graphique. Pour plus d’informations, voir Excel. ChartTitlePosition.|
||[height](/javascript/api/excel/excel.charttitle#height)|Représente la hauteur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
||[width](/javascript/api/excel/excel.charttitle#width)|Représente la largeur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
||[setFormula (Formula : String)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Définit une valeur de chaîne qui représente la formule de titre de graphique à l’aide de la notation de style A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Représente l’orientation du texte du titre du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttitle#top)|Représente la distance en points, du bord supérieur du titre du graphique au bord supérieur de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Représente l’alignement vertical du titre du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[route](/javascript/api/excel/excel.charttitleformat#border)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de style et le poids. En lecture seule.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Représente le nombre de points que la courbe de tendance étend en arrière.|
||[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Supprime l’objet courbe de tendance.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Représente le nombre de points que la courbe de tendance étend en avant.|
||[ordonn](/javascript/api/excel/excel.charttrendline#intercept)|Représente la valeur intercept de la courbe de tendance. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs automatiques). La valeur renvoyée est toujours un nombre.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Représente la période d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Représente le nom de la courbe de tendance. Peut être configurée pour une valeur de chaîne, ou peut être configurée pour que la valeur null représente les valeurs automatiques. La valeur renvoyée est toujours une chaîne.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Représente l’ordre d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type polynomial.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Représente la mise en forme de courbe de tendance de graphique.|
||[étiquette](/javascript/api/excel/excel.charttrendline#label)|Représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Représente le type de courbe de tendance de graphique.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[Add (type ?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Ajoute une nouvelle courbe de tendance à la collection de courbes de tendance.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Renvoie le nombre de courbes de tendance de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Obtient un objet courbe de tendance par index, c'est-à-dire par ordre d’insertion dans le tableau des éléments.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Représente le format des lignes du graphique. En lecture seule.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.customproperty#key)|Obtient la clé de la propriété personnalisée. En lecture seule.|
||[type](/javascript/api/excel/excel.customproperty#type)|Obtient le type de valeur de la propriété personnalisée. En lecture seule.|
||[value](/javascript/api/excel/excel.customproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[Add (Key : chaîne, value : any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Lève une exception si la propriété personnalisée n’existe pas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[RefreshAll, ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Actualise toutes les dataConnections de la collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[créés](/javascript/api/excel/excel.documentproperties#author)|Obtient ou définit l’auteur du classeur.|
||[catégories](/javascript/api/excel/excel.documentproperties#category)|Obtient ou définit la catégorie du classeur.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Obtient ou définit les commentaires du classeur.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Obtient ou définit la compagnie du classeur.|
||[Mots clés](/javascript/api/excel/excel.documentproperties#keywords)|Obtient ou définit les mots clés du classeur.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|Obtient ou définit le responsable du classeur.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Obtient la date de création du classeur. En lecture seule.|
||[personnalisé](/javascript/api/excel/excel.documentproperties#custom)|Obtient la collection de propriétés personnalisées du classeur. En lecture seule.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Obtient ou définit le dernier auteur du classeur. En lecture seule.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Obtient le numéro de révision du classeur. En lecture seule.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Obtient ou définit le sujet du classeur.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Obtient ou définit le titre du classeur.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Obtient ou définit la formule de l’élément nommé.  La formule commence toujours par un signe '='.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Renvoie un objet contenant les valeurs et les types de l’élément nommé. En lecture seule.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Représente les types de chaque élément dans le tableau d’éléments nommés|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Représente les valeurs de chaque élément dans le tableau élément nommé.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows : nombre, numColumns : nombre)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Obtient un objet Plage avec la même cellule supérieure gauche que l’objet de Plage en cours, mais avec un nombre spécifié de lignes et colonnes.|
||[getImage ()](/javascript/api/excel/excel.range#getimage--)|Affiche la plage en tant qu’image png encodée au format Base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Renvoie un objet PLage qui représente la région environnante pour la cellule en haut à gauche de cette plage. Une région environnante est une plage délimitée par une combinaison de lignes et de colonnes vides par rapport à cette plage.|
||[lien hypertexte](/javascript/api/excel/excel.range#hyperlink)|Représente le lien hypertexte de la plage active.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Représente le code de format numérique d’Excel pour la plage donnée en tant que chaîne dans la langue de l’utilisateur.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Représente si la plage active est une colonne entière. En lecture seule.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Représente si la plage active est une ligne entière. En lecture seule.|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|Affiche la carte pour une cellule active si son contenu est riche en valeur.|
||[style](/javascript/api/excel/excel.range#style)|Représente le style de la plage actuelle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Obtient ou définit l’orientation du texte de toutes les cellules dans la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Indique si la largeur de la colonne de l’objet Range est égale à la largeur standard de la feuille.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[adresse](/javascript/api/excel/excel.rangehyperlink#address)|Représente l’url cible du lien hypertexte.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Représente la cible de référence de document pour le lien hypertexte.|
||[Info](/javascript/api/excel/excel.rangehyperlink#screentip)|Représente la chaîne affichée lorsque vous pointez sur le lien hypertexte.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Représente la chaîne qui s’affiche dans la cellule en haut à gauche de la plage.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|
||[delete()](/javascript/api/excel/excel.style#delete--)|Supprime ce style.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Indique si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Représente l’alignement horizontal pour le style. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[IncludeAlignment,](/javascript/api/excel/excel.style#includealignment)|Indique si le style inclut les propriétés AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, et TextOrientation.|
||[IncludeBorder,](/javascript/api/excel/excel.style#includeborder)|Indique si le style inclut les propriétés dColor, ColorIndex, LineStyle, et Weight border.|
||[IncludeFont,](/javascript/api/excel/excel.style#includefont)|Indique si le style inclut les propriétés dBackground, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, et Underline font.|
||[IncludeNumber,](/javascript/api/excel/excel.style#includenumber)|Indique si le style inclut la propriété NumberFormat.|
||[IncludePatterns,](/javascript/api/excel/excel.style#includepatterns)|Indique si le style inclut les propriétés Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, et PatternColorIndex interior.|
||[IncludeProtection,](/javascript/api/excel/excel.style#includeprotection)|Indique si le style inclut les propriétés FormulaHidden et Locked protection.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[locked](/javascript/api/excel/excel.style#locked)|Indique si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|Le code de format du nombre format pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|Le code de format localisé du nombre format pour le style.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|L’ordre de lecture du style.|
||[Borders](/javascript/api/excel/excel.style#borders)|Collection de bordures de quatre objets qui représente le style des quatre bordures.|
||[builtIn](/javascript/api/excel/excel.style#builtin)|Indique si le style est un style intégré.|
||[fill](/javascript/api/excel/excel.style#fill)|Remplissage du style.|
||[police](/javascript/api/excel/excel.style#font)|Renvoie un objet Police qui représente la police du style.|
||[name](/javascript/api/excel/excel.style#name)|Nom du style.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Indique si le texte s’ajuste automatiquement pour tenir dans la largeur de colonne disponible.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|L’orientation du texte pour le style.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Représente l’alignement vertical du style. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Indique si Microsoft Excel renvoie le texte à la ligne dans l’objet.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Ajoute un nouveau style à la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Obtient un style par nom.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Tableau](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Survient lorsque les données des cellules changent sur une table spécifique.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Se produit lorsque la sélection change sur une table spécifique.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[adresse](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtient l’adresse qui représente la zone modifiée d’un tableau figurant dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié. Pour plus d’informations, voir Excel. DataChangeType.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Obtient la plage qui représente la zone modifiée d’un tableau dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Obtient la plage qui représente la zone modifiée d’un tableau dans une feuille de calcul spécifique. Il peut renvoyer un objet null.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Obtient l’id du tableau dans lequel les données sont modifiées.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Survient lors de la modification des données d’une table dans un classeur ou d’une feuille de calcul.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone sélectionnée d’un tableau dans une feuille de calcul spécifique.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Indique si la sélection est dans un tableau, l’adresse sera superflue si IsInsideTable est faux.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Obtient l’id du tableau dans lequel la sélection est modifiée.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType. En lecture seule.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|
|[Classeur](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Obtient la cellule active du classeur.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Représente toutes les connexions de données du classeur. En lecture seule.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtient le nom du classeur. En lecture seule.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtient les propriétés du classeur. En lecture seule.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Renvoie un objet de protection de classeur pour un classeur. En lecture seule.|
||[proposés](/javascript/api/excel/excel.workbook#styles)|Représente une collection de styles associés au classeur. En lecture seule.|
|[Objetworkbookprotection](/javascript/api/excel/excel.workbookprotection)|[Protect (Password ?: String)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protège un classeur. Échoue si le classeur est protégé.|
||[sécurisé](/javascript/api/excel/excel.workbookprotection#protected)|Indique si le classeur est protégé. En lecture seule.|
||[Unprotect (Password ?: String)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Annule la protection un classeur.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[Copy (positionType ?: Excel. WorksheetPositionType, relativeTo ?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copie une feuille de calcul et la place à la position spécifiée. Renvoie la feuille de calcul copiée.|
||[getRangeByIndexes (startRow : nombre, ColonneDébut : nombre, rowCount : nombre, columnCount : nombre)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Obtient l’objet plage commençant à un index de ligne et de colonne particulier et couvrant un certain nombre de lignes et de colonnes.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Obtient un objet qui peut être utilisé pour manipuler les volets figés de la feuille de calcul. En lecture seule.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Se produit lorsque la feuille de calcul est activée.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Se produit lorsque des données sont modifiées dans une feuille de calcul spécifique.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Se produit lorsque la feuille de calcul est désactivée.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Se produit lorsque la sélection change dans une feuille de calcul spécifique.|
||[StandardHeight,](/javascript/api/excel/excel.worksheet#standardheight)|Renvoie la hauteur standard (par défaut) de toutes les lignes dans la feuille de calcul, en points. En lecture seule.|
||[StandardWidth,](/javascript/api/excel/excel.worksheet#standardwidth)|Renvoie ou définit la largeur standard (par défaut) de toutes les colonnes dans la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|Obtient ou modifie la couleur d’onglet de la feuille de calcul.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est activée.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est ajoutée au classeur.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié. Pour plus d’informations, voir Excel. DataChangeType.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique. Il peut renvoyer un objet null.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est activée.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Survient lors de l’ajout d’une nouvelle feuille de calcul au classeur.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est désactivée.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Survient lors de la suppression d’une feuille de calcul du classeur.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est desactivée.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est supprimée du classeur.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange : chaîne de plage \| )](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[freezeColumns (Count ?: nombre)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Figer la/les première(s) colonne(s) de la feuille de calcul en place.|
||[freezeRows (Count ?: nombre)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Figer la/les première(s) ligne(s) de la feuille de calcul en place.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[Unfreeze ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Supprime tous les volets figés dans la feuille de calcul.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[Unprotect (Password ?: String)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Annule la protection d’une feuille de calcul.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Représente l’option de protection de feuille de calcul qui autorise la modification d’objets.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Représente l’option de protection de feuille de calcul qui autorise la modification de scénarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Représente l’option de protection de feuille de calcul qui autorise le mode sélection.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone sélectionnée dans une feuille de calcul spécifique.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
