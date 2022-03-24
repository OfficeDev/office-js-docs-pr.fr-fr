---
title: Excel conditions requises de l’API JavaScript 1.7
description: Détails sur l’ensemble de conditions requises ExcelApi 1.7.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: cd8f0f333b76306a6feecff95b9ba8831428606a
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744527"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Nouveautés de l’API JavaScript 1.7 pour Excel

Les fonctionnalités Excel JavaScript API ensemble de conditions 1.7 incluent des API pour les graphiques, événements, feuilles de calcul, plages, propriétés de document, éléments nommés, options de protection et styles.

## <a name="customize-charts"></a>Personnaliser des graphiques

Avec le nouvel API graphique, vous pouvez créer des types de graphiques supplémentaires, ajouter une série de données à un graphique, définir le titre du graphique, ajouter un titre d’axe, ajouter une unité d’affichage, ajouter une courbe de tendance avec moyenne mobile, modifier une courbe de tendance en ligne, et bien plus encore. Voici quelques exemples.

- Axe du graphique - obtenir, définir, mettre en forme et supprimer une unité d’axe, une étiquette et un titre dans un graphique.
- Série de graphique - ajouter, configurer et supprimer une série dans un graphique.  Modifier les marqueurs de série, les commandes traçage et le redimensionnement.
- Courbes de tendance de graphique - ajouter, obtenir et mettre en forme des courbes de tendance dans un graphique.
- Légende de graphique - mettre en forme la police de légende dans un graphique.
- Point de graphique - définir la couleur du point de graphique.
- Sous-chaîne de titre du graphique - obtenir et définir une sous-chaîne de titre d’un graphique.
- Type de graphique - option pour créer plusieurs types de graphiques.

## <a name="events"></a>Événements

Les API Événements pour Excel fournissent un grand nombre de gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. Pour une liste des événements qui sont actuellement disponibles, voir [Manipuler des Événements à l’aide de l’API JavaScript Excel](../../excel/excel-add-ins-events.md).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personnaliser l’apparence de feuilles de calcul et des plages

À l’aide des nouveaux API, vous pouvez personnaliser l’apparence de feuilles de calcul de plusieurs façons :

- Figer les volets pour conserver certaines lignes ou colonnes visibles lorsque vous faites défiler la feuille de calcul. Par exemple, si la première ligne dans votre feuille de calcul contient des en-têtes, vous pouvez figer cette ligne de sorte que les en-têtes de colonne restent visibles pendant le défilement vers le bas de la feuille de calcul.
- Modifier la couleur d’onglet de la feuille de calcul.
- Ajouter des en-têtes de feuille de calcul.

Vous pouvez personnaliser l’apparence des plages de plusieurs façons :

- Définir le style de cellule pour une plage pour vous assurer que toutes les cellules dans la plage ont une mise en forme cohérente. Un style de cellule est un ensemble défini de caractéristiques de mise en forme, comme les polices et les tailles de police, formats des nombres, bordures de cellule et ombrage de cellule. Utilisez un des styles de cellule intégrés d’Excel ou créer votre propre style de cellule personnalisé.
- Définit l’orientation du texte pour une plage.
- Ajouter ou modifier un lien hypertexte sur une plage qui permet d’accéder à un autre emplacement dans le classeur ou à un emplacement externe.

## <a name="manage-document-properties"></a>Gérer les propriétés du document

À l’aide des API de propriétés du document, vous pouvez accéder aux propriétés de document intégrées et également créer et gérer les propriétés de document personnalisées pour stocker l’état du classeur et lire le flux de travail et la logique d’entreprise.

## <a name="copy-worksheets"></a>Obtenir des feuilles de calcul

À l’aide des API de copie de feuille de calcul , vous pouvez copier les données et le format à partir d’une feuille de calcul dans une nouvelle feuille de calcul au sein du même classeur et réduire la quantité de transfert de données nécessaire.

## <a name="handle-ranges-with-ease"></a>Gérer les plages en toute simplicité

À l’aide des API de plage différente, vous pouvez effectuer des actions telles qu’obtenir la région environnante, obtenir une plage redimensionnée et bien plus encore. Ces API doivent rendre des tâches telles que la manipulation de plage et l’adressage beaucoup plus efficaces.

De plus :

- Options de protection de classeur et feuille de calcul : utilisez ces API pour protéger les données dans une feuille de calcul et la structure du classeur.
- Mettre à jour un élément nommé : utilisez cet API pour mettre à jour un élément nommé.
- Obtenir la cellule active : utilisez cet API pour obtenir la cellule active d’un classeur.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API Excel l’ensemble de conditions requises de l’API JavaScript 1.7. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.7 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#excel-excel-chart-charttype-member)|Spécifie le type du graphique.|
||[id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member)|L’ID unique du graphique.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#excel-excel-chart-showallfieldbuttons-member)|Spécifie s’il faut afficher tous les boutons de champ sur une PivotChart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[bordure](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-border-member)|Représente le format de bordure de la zone de graphique, qui inclut la couleur, le style de trait et l’pondération.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type : Excel. ChartAxisType, group? : Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-getitem-member(1))|Renvoie l’axe spécifique identifié par type et par groupe.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[axisGroup](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-axisgroup-member)|Spécifie le groupe de l’axe spécifié.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-basetimeunit-member)|Spécifie l’unité de base de l’axe des catégories spécifié.|
||[categoryType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-categorytype-member)|Spécifie le type d’axe des catégories.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-customdisplayunit-member)|Spécifie la valeur d’unité d’affichage de l’axe personnalisé.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-displayunit-member)|Représente l’unité d’affichage de l’axe.|
||[height](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-height-member)|Spécifie la hauteur, en points, de l’axe du graphique.|
||[left](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-left-member)|Spécifie la distance, en points, entre le bord gauche de l’axe et la gauche de la zone de graphique.|
||[logBase](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-logbase-member)|Spécifie la base du logarithme lors de l’utilisation des échelles logarithmiques.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortickmark-member)|Spécifie le type de la coche principale de l’axe spécifié.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortimeunitscale-member)|Spécifie la valeur d’échelle d’unité principale pour l’axe des catégories lorsque la `categoryType` propriété est définie `dateAxis`sur .|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortickmark-member)|Spécifie le type de marque de cocher mineure pour l’axe spécifié.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortimeunitscale-member)|Spécifie la valeur d’échelle d’unité mineure pour l’axe des catégories lorsque la `categoryType` propriété est définie sur `dateAxis`.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-reverseplotorder-member)|Spécifie si Excel points de données du dernier au premier.|
||[scaleType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-scaletype-member)|Spécifie le type d’échelle de l’axe des valeurs.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcategorynames-member(1))|Définit tous les noms de catégorie pour l’axe spécifié.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcustomdisplayunit-member(1))|Définit l’unité d’affichage axe sur une valeur personnalisée.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-showdisplayunitlabel-member)|Spécifie si l’étiquette d’unité d’affichage de l’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelposition-member)|Spécifie la position des étiquettes de graduation sur l'axe spécifié.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelspacing-member)|Spécifie le nombre d’catégories ou de séries entre les étiquettes de coche.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-tickmarkspacing-member)|Spécifie le nombre de catégories ou de séries entre les marques de cocher.|
||[top](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-top-member)|Spécifie la distance, en points, entre le bord supérieur de l’axe et le haut de la zone de graphique.|
||[type](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-type-member)|Spécifie le type d’axe.|
||[visible](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-visible-member)|Spécifie si l’axe est visible.|
||[width](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-width-member)|Spécifie la largeur, en points, de l’axe du graphique.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-color-member)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-linestyle-member)|Représente le style de trait de la bordure.|
||[weight](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-weight-member)|Représente l’épaisseur de bordure, en points.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-position-member)|Valeur qui représente la position de l’étiquette de données.|
||[séparateur](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-separator-member)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showbubblesize-member)|Spécifie si la taille des bulles des étiquettes de données est visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showcategoryname-member)|Spécifie si le nom de catégorie d’étiquette de données est visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showlegendkey-member)|Spécifie si le clé de légende d’étiquette de données est visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showpercentage-member)|Spécifie si le pourcentage d’étiquette de données est visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showseriesname-member)|Spécifie si le nom de la série d’étiquettes de données est visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showvalue-member)|Spécifie si la valeur de l’étiquette de données est visible.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[police](/javascript/api/excel/excel.chartformatstring#excel-excel-chartformatstring-font-member)|Représente les attributs de police, tels que le nom de police, la taille de police et la couleur d’un objet de caractères de graphique.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-height-member)|Spécifie la hauteur, en points, de la légende sur le graphique.|
||[left](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-left-member)|Spécifie la valeur gauche, en points, de la légende du graphique.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-legendentries-member)|Représente une collection de legendEntries dans la légende.|
||[showShadow](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-showshadow-member)|Spécifie si la légende possède une ombre sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-top-member)|Spécifie le haut d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-width-member)|Spécifie la largeur, en points, de la légende sur le graphique.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-visible-member)|Représente la visibilité d’une entrée de légende de graphique.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getcount-member(1))|Renvoie le nombre d’entrées de légende dans la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getitemat-member(1))|Renvoie une entrée de légende à l’index donné.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-linestyle-member)|Représente le style de trait.|
||[weight](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-weight-member)|Représente l’épaisseur de bordure, en points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[dataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-datalabel-member)|Renvoie l’étiquette de données d’un point du graphique.|
||[hasDataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-hasdatalabel-member)|Indique si un point de données possède une étiquette de données.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerbackgroundcolor-member)|Représentation de code couleur HTML de la couleur d’arrière-plan de marque d’un point de données (par exemple, #FF0000 représente le rouge).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerforegroundcolor-member)|Représentation de code couleur HTML de la couleur de premier plan du marqueur d’un point de données (par exemple, #FF0000 représente le rouge).|
||[markerSize](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markersize-member)|Représente la taille du marqueur d’un point de données.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerstyle-member)|Représente le style du marqueur du point de données de graphique.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[bordure](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-border-member)|Représente le format de bordure d’un point de données de graphique, qui inclut des informations sur la couleur, le style et l’poids.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-charttype-member)|Représente le type de graphique d’une série.|
||[delete()](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-delete-member(1))|Supprime la série graphique.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-doughnutholesize-member)|Représente la taille du centre d’une série de graphiques en anneaux.|
||[filtered](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-filtered-member)|Spécifie si la série est filtrée.|
||[gapWidth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gapwidth-member)|Représente la largeur de l’intervalle d’une série de graphique.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-hasdatalabels-member)|Spécifie si la série possède des étiquettes de données.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerbackgroundcolor-member)|Spécifie la couleur d’arrière-plan du marqueur d’une série de graphiques.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerforegroundcolor-member)|Spécifie la couleur de premier plan du marqueur d’une série de graphiques.|
||[markerSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markersize-member)|Spécifie la taille de marqueur d’une série de graphiques.|
||[markerStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerstyle-member)|Spécifie le style de marqueur d’une série de graphiques.|
||[plotOrder](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-plotorder-member)|Spécifie l’ordre de traçage d’une série de graphiques dans le groupe de graphiques.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setbubblesizes-member(1))|Définit les tailles des bulles pour une série de graphiques.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setvalues-member(1))|Définit les valeurs d’une série de graphiques.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setxaxisvalues-member(1))|Définit les valeurs de l’axe des x pour une série de graphiques.|
||[showShadow](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showshadow-member)|Spécifie si la série possède une ombre.|
||[smooth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-smooth-member)|Spécifie si la série est lisse.|
||[trendlines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-trendlines-member)|Collection des tendances de la série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-add-member(1))|Ajouter une nouvelle série à la collection.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-getsubstring-member(1))|Obtenir la sous-stration d’un titre de graphique.|
||[height](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-height-member)|Représente la hauteur, exprimée en points, du titre du graphique.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-horizontalalignment-member)|Spécifie l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-left-member)|Spécifie la distance, en points, entre le bord gauche du titre du graphique et le bord gauche de la zone de graphique.|
||[position](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-position-member)|Représente la position du titre du graphique.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-setformula-member(1))|Définit une valeur de chaîne qui représente la formule de titre de graphique à l’aide de la notation de style A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-showshadow-member)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-textorientation-member)|Spécifie l’angle vers lequel le texte est orienté pour le titre du graphique.|
||[top](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-top-member)|Spécifie la distance, en points, entre le bord supérieur du titre du graphique et le haut de la zone de graphique.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-verticalalignment-member)|Spécifie l’alignement vertical du titre du graphique.|
||[width](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-width-member)|Spécifie la largeur, en points, du titre du graphique.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[bordure](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-border-member)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de trait et l’pondération.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-delete-member(1))|Supprime l’objet courbe de tendance.|
||[format](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-format-member)|Représente la mise en forme de courbe de tendance de graphique.|
||[intercept](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-intercept-member)|Représente la valeur intercept de la courbe de tendance.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-movingaverageperiod-member)|Représente la période d’une courbe de tendance de graphique.|
||[name](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-name-member)|Représente le nom de la courbe de tendance.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-polynomialorder-member)|Représente l’ordre d’une courbe de tendance de graphique.|
||[type](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-type-member)|Représente le type de courbe de tendance de graphique.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-add-member(1))|Ajoute une nouvelle courbe de tendance à la collection de courbes de tendance.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getcount-member(1))|Renvoie le nombre de courbes de tendance de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getitem-member(1))|Obtient un objet de courbe de tendance par index, qui est l’ordre d’insertion dans le tableau d’éléments.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#excel-excel-charttrendlineformat-line-member)|Représente le format des lignes du graphique.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-delete-member(1))|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-key-member)|Clé de la propriété personnalisée.|
||[type](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-type-member)|Type de la valeur utilisée pour la propriété personnalisée.|
||[value](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-value-member)|Valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-add-member(1))|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-deleteall-member(1))|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getcount-member(1))|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitem-member(1))|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitemornullobject-member(1))|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#excel-excel-dataconnectioncollection-refreshall-member(1))|Actualise toutes les connexions de données dans la collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-author-member)|Auteur du livre.|
||[category](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-category-member)|Catégorie du classez.|
||[comments](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-comments-member)|Commentaires du workbook.|
||[company](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-company-member)|Société du workbook.|
||[creationDate](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-creationdate-member)|Obtient la date de création du classeur.|
||[custom](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member)|Obtient la collection de propriétés personnalisées du classeur.|
||[keywords](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-keywords-member)|Mots clés du manuel.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-lastauthor-member)|Obtient ou définit le dernier auteur du classeur.|
||[manager](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-manager-member)|Responsable du manuel.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-revisionnumber-member)|Obtient le numéro de révision du classeur.|
||[subject](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-subject-member)|Objet du manuel.|
||[title](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-title-member)|Titre du manuel.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[arrayValues](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-arrayvalues-member)|Renvoie un objet contenant les valeurs et les types de l’élément nommé.|
||[formula](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-formula-member)|Formule de l’élément nommé.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-types-member)|Représente les types de chaque élément dans le tableau d’éléments nommés|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-values-member)|Représente les valeurs de chaque élément dans le tableau élément nommé.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#excel-excel-range-getabsoluteresizedrange-member(1))|Obtient un `Range` objet avec la même cellule supérieure `Range` gauche que l’objet actuel, mais avec le nombre spécifié de lignes et de colonnes.|
||[getImage()](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1))|Restituer la plage en tant qu’image png codée en base 64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#excel-excel-range-getsurroundingregion-member(1))|Renvoie un `Range` objet qui représente la région environnante pour la cellule supérieure gauche de cette plage.|
||[lien hypertexte](/javascript/api/excel/excel.range#excel-excel-range-hyperlink-member)|Représente le lien hypertexte de la plage actuelle.|
||[isEntireColumn](/javascript/api/excel/excel.range#excel-excel-range-isentirecolumn-member)|Représente si la plage active est une colonne entière.|
||[isEntireRow](/javascript/api/excel/excel.range#excel-excel-range-isentirerow-member)|Représente si la plage active est une ligne entière.|
||[numberFormatLocal](/javascript/api/excel/excel.range#excel-excel-range-numberformatlocal-member)|Représente le code Excel format numérique de la plage donnée, en fonction des paramètres de langue de l’utilisateur.|
||[showCard()](/javascript/api/excel/excel.range#excel-excel-range-showcard-member(1))|Affiche la carte pour une cellule active si son contenu est riche en valeur.|
||[style](/javascript/api/excel/excel.range#excel-excel-range-style-member)|Représente le style de la plage actuelle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-textorientation-member)|Orientation du texte de toutes les cellules de la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardheight-member)|Détermine si la hauteur de ligne de l’objet `Range` est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardwidth-member)|Spécifie si la largeur de colonne de l’objet `Range` est égale à la largeur standard de la feuille.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[adresse](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-address-member)|Représente la cible d’URL pour le lien hypertexte.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-documentreference-member)|Représente la cible de référence du document pour le lien hypertexte.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-screentip-member)|Représente la chaîne affichée lorsque vous pointez sur le lien hypertexte.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-texttodisplay-member)|Représente la chaîne qui s’affiche dans la cellule en haut à gauche de la plage.|
|[Style](/javascript/api/excel/excel.style)|[Borders](/javascript/api/excel/excel.style#excel-excel-style-borders-member)|Collection de quatre objets de bordure qui représentent le style des quatre bordures.|
||[builtIn](/javascript/api/excel/excel.style#excel-excel-style-builtin-member)|Spécifie si le style est un style intégré.|
||[delete()](/javascript/api/excel/excel.style#excel-excel-style-delete-member(1))|Supprime ce style.|
||[fill](/javascript/api/excel/excel.style#excel-excel-style-fill-member)|Remplissage du style.|
||[police](/javascript/api/excel/excel.style#excel-excel-style-font-member)|Objet `Font` qui représente la police du style.|
||[formulaHidden](/javascript/api/excel/excel.style#excel-excel-style-formulahidden-member)|Spécifie si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.style#excel-excel-style-horizontalalignment-member)|Représente l’alignement horizontal pour le style.|
||[includeAlignment](/javascript/api/excel/excel.style#excel-excel-style-includealignment-member)|Spécifie si le style inclut le retrait automatique, l’alignement horizontal, l’alignement vertical, le texte de wrap, le niveau de retrait et les propriétés d’orientation du texte.|
||[includeBorder](/javascript/api/excel/excel.style#excel-excel-style-includeborder-member)|Indique si le style inclut les propriétés de couleur, d’index de couleur, de style de trait et de bordure de poids.|
||[includeFont](/javascript/api/excel/excel.style#excel-excel-style-includefont-member)|Spécifie si le style inclut les propriétés d’arrière-plan, de gras, de couleur, d’index de couleur, de style de police, d’italique, de nom, de taille, de strikethrough, d’indice, d’exposant et de soulignement de police.|
||[includeNumber](/javascript/api/excel/excel.style#excel-excel-style-includenumber-member)|Spécifie si le style inclut la propriété de format numérique.|
||[includePatterns](/javascript/api/excel/excel.style#excel-excel-style-includepatterns-member)|Spécifie si le style inclut la couleur, l’index de couleur, l’inversion si négatif, le motif, la couleur de motif et les propriétés de l’intérieur de l’index de couleur de motif.|
||[includeProtection](/javascript/api/excel/excel.style#excel-excel-style-includeprotection-member)|Spécifie si le style inclut les propriétés de protection masquées et verrouillées de la formule.|
||[indentLevel](/javascript/api/excel/excel.style#excel-excel-style-indentlevel-member)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[locked](/javascript/api/excel/excel.style#excel-excel-style-locked-member)|Spécifie si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[name](/javascript/api/excel/excel.style#excel-excel-style-name-member)|Nom du style.|
||[numberFormat](/javascript/api/excel/excel.style#excel-excel-style-numberformat-member)|Le code de format du nombre format pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#excel-excel-style-numberformatlocal-member)|Le code de format localisé du nombre format pour le style.|
||[readingOrder](/javascript/api/excel/excel.style#excel-excel-style-readingorder-member)|L’ordre de lecture du style.|
||[shrinkToFit](/javascript/api/excel/excel.style#excel-excel-style-shrinktofit-member)|Spécifie si le texte est automatiquement réduit pour tenir dans la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.style#excel-excel-style-verticalalignment-member)|Spécifie l’alignement vertical du style.|
||[wrapText](/javascript/api/excel/excel.style#excel-excel-style-wraptext-member)|Spécifie si Excel le texte dans l’objet.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|Ajoute un nouveau style à la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitem-member(1))|Obtient une `Style` par nom.|
||[items](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)|Se produit lorsque les données des cellules changent dans un tableau spécifique.|
||[onSelectionChanged](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)|Se produit lorsque la sélection change dans un tableau spécifique.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[adresse](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-address-member)|Obtient l’adresse qui représente la zone modifiée d’un tableau figurant dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-changetype-member)|Obtient le type de modification qui représente la façon dont l’événement modifié est déclenché.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-source-member)|Obtient la source de l’événement.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-tableid-member)|Obtient l’ID de la table dans laquelle les données ont été modifiées.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle les données ont été modifiées.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)|Se produit lorsque des données changent dans une table d’un workbook ou d’une feuille de calcul.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-address-member)|Obtient l’adresse de plage qui représente la zone sélectionnée d’un tableau dans une feuille de calcul spécifique.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-isinsidetable-member)|Spécifie si la sélection se trouve à l’intérieur d’un tableau.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-tableid-member)|Obtient l’ID du tableau dans lequel la sélection a été modifiée.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle la sélection a été modifiée.|
|[Workbook](/javascript/api/excel/excel.workbook)|[dataConnections](/javascript/api/excel/excel.workbook#excel-excel-workbook-dataconnections-member)|Représente toutes les connexions de données dans le workbook.|
||[getActiveCell()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivecell-member(1))|Obtient la cellule active du classeur.|
||[name](/javascript/api/excel/excel.workbook#excel-excel-workbook-name-member)|Obtient le nom du classeur.|
||[properties](/javascript/api/excel/excel.workbook#excel-excel-workbook-properties-member)|Obtient les propriétés du classeur.|
||[protection](/javascript/api/excel/excel.workbook#excel-excel-workbook-protection-member)|Renvoie l’objet de protection d’un workbook.|
||[styles](/javascript/api/excel/excel.workbook#excel-excel-workbook-styles-member)|Représente une collection de styles associés au classeur.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protect-member(1))|Protège un classeur.|
||[protected](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protected-member)|Spécifie si le workbook est protégé.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-unprotect-member(1))|Annule la protection un classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel. WorksheetPositionType, relativeTo? : Excel. Feuille de calcul)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-copy-member(1))|Copie une feuille de calcul et la place à la position spécifiée.|
||[freezePanes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezepanes-member)|Obtient un objet qui peut être utilisé pour manipuler des volets figés dans la feuille de calcul.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrangebyindexes-member(1))|Obtient l’objet `Range` qui commence à un index de ligne et un index de colonne particuliers et s’étend sur un certain nombre de lignes et de colonnes.|
||[onActivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)|Se produit lorsque la feuille de calcul est activée.|
||[onChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member)|Se produit lorsque des données changent dans une feuille de calcul spécifique.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|Se produit lorsque la feuille de calcul est désactivée.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)|Se produit lorsque la sélection change dans une feuille de calcul spécifique.|
||[standardHeight](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardheight-member)|Renvoie la hauteur standard (par défaut) de toutes les lignes dans la feuille de calcul, en points.|
||[standardWidth](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardwidth-member)|Spécifie la largeur standard (par défaut) de toutes les colonnes de la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabcolor-member)|Couleur de l’onglet de la feuille de calcul.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul qui est activée.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-source-member)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul qui est ajoutée au manuel.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-address-member)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|Obtient le type de modification qui représente la façon dont l’événement modifié est déclenché.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-source-member)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle les données ont été modifiées.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member)|Se produit lorsqu’une feuille de calcul du manuel est activée.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member)|Se produit lorsqu’une nouvelle feuille de calcul est ajoutée au manuel.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member)|Se produit lorsqu’une feuille de calcul du manuel est désactivée.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member)|Se produit lorsqu’une feuille de calcul est supprimée du manuel.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul qui est désactivée.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-source-member)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul qui est supprimée du manuel.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezeat-member(1))|Définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezecolumns-member(1))|Figer la première ou les premières colonnes de la feuille de calcul en place.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezerows-member(1))|Figer la ou les lignes du haut de la feuille de calcul en place.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocation-member(1))|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocationornullobject-member(1))|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-unfreeze-member(1))|Supprime tous les volets figés dans la feuille de calcul.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-unprotect-member(1))|Annule la protection d’une feuille de calcul.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditobjects-member)|Représente l’option de protection de feuille de calcul permettant la modification d’objets.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditscenarios-member)|Représente l’option de protection de feuille de calcul qui permet la modification des scénarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-selectionmode-member)|Représente l’option de protection de feuille de calcul qui autorise le mode sélection.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-address-member)|Obtient l’adresse de plage qui représente la zone sélectionnée dans une feuille de calcul spécifique.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle la sélection a été modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
