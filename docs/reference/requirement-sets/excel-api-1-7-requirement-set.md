---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.7
description: Détails sur l’ensemble de conditions requises ExcelApi 1.7.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 67f30fd61e3065f8d7d193668c6f79fd09debf2f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350210"
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

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.7. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.7 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Spécifie le type du graphique.|
||[id](/javascript/api/excel/excel.chart#id)|ID unique du graphique.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Spécifie s’il faut afficher tous les boutons de champ sur une PivotChart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|Représente le format de bordure de la zone de graphique, qui inclut la couleur, le style de trait et l’pondération.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type : Excel. ChartAxisType, group? : Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Renvoie l’axe spécifique identifié par type et par groupe.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Spécifie l’unité de base de l’axe des catégories spécifié.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Spécifie le type d’axe des catégories.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Représente l’unité d’affichage de l’axe.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Spécifie la base du logarithme lors de l’utilisation des échelles logarithmiques.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Spécifie le type de la coche principale de l’axe spécifié.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Spécifie la valeur d’échelle d’unité principale pour l’axe des catégories lorsque la propriété CategoryType est définie sur Échelle de temps.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Spécifie le type de marque de cocher mineure pour l’axe spécifié.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Spécifie la valeur d’échelle d’unité mineure pour l’axe des catégories lorsque la propriété CategoryType est définie sur Échelle de temps.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Spécifie le groupe de l’axe spécifié.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Spécifie la valeur d’unité d’affichage de l’axe personnalisé.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Spécifie la hauteur, en points, de l’axe du graphique.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Spécifie la distance, en points, entre le bord gauche de l’axe et la gauche de la zone de graphique.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Spécifie la distance, en points, entre le bord supérieur de l’axe et le haut de la zone de graphique.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Spécifie le type d’axe.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Spécifie la largeur, en points, de l’axe du graphique.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Spécifie si Excel points de données du dernier au premier.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Spécifie le type d’échelle de l’axe des valeurs.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Définit tous les noms de catégorie pour l’axe spécifié.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Définit l’unité d’affichage axe sur une valeur personnalisée.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Spécifie si l’étiquette d’unité d’affichage de l’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Spécifie la position des étiquettes de graduation sur l'axe spécifié.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Spécifie le nombre d’catégories ou de séries entre les étiquettes de coche.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Spécifie le nombre de catégories ou de séries entre les marques de cocher.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Spécifie si l’axe est visible.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Représente le style de trait de la bordure.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données.|
||[séparateur](/javascript/api/excel/excel.chartdatalabel#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Spécifie si la taille des bulles des étiquettes de données est visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Spécifie si le nom de catégorie d’étiquette de données est visible.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Spécifie si le clé de légende d’étiquette de données est visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Spécifie si le pourcentage d’étiquette de données est visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Spécifie si le nom de la série d’étiquettes de données est visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Spécifie si la valeur de l’étiquette de données est visible.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[police](/javascript/api/excel/excel.chartformatstring#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Spécifie la hauteur, en points, de la légende sur le graphique.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Spécifie la gauche, en points, de la légende du graphique.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Représente une collection de legendEntries dans la légende.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Spécifie si la légende possède une ombre sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Spécifie le haut d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Spécifie la largeur, en points, de la légende du graphique.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Représente la partie visible d’une entrée de légende de graphique.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Renvoie le nombre de legendEntry de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Renvoie un legendEntry à l’index donné.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Représente le style de trait.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Indique si un point de données possède une étiquette de données.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Représentation de code couleur HTML de la couleur d’arrière-plan du marqueur du point de données (par exemple, #FF0000 représente le rouge).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Représentation de code couleur HTML de la couleur de premier plan du marqueur du point de données (par exemple, #FF0000 représente le rouge).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Représente la taille du marqueur du point de données.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Représente le style du marqueur du point de données de graphique.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Renvoie l’étiquette de données d’un point du graphique.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|Représente le format de bordure d’un point de données de graphique, qui inclut des informations sur la couleur, le style et l’poids.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Représente le type de graphique d’une série.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Supprime la série graphique.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Représente la taille du centre d’une série de graphiques en anneaux.|
||[filtered](/javascript/api/excel/excel.chartseries#filtered)|Spécifie si la série est filtrée.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Représente la largeur de l’intervalle d’une série de graphique.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Spécifie si la série possède des étiquettes de données.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Spécifie la couleur d’arrière-plan des marqueurs d’une série de graphiques.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Spécifie la couleur de premier plan des marqueurs d’une série de graphiques.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Spécifie la taille de marqueur d’une série de graphiques.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|Spécifie le style de marqueur d’une série de graphiques.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Spécifie l’ordre de traçage d’une série de graphiques dans le groupe de graphiques.|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Collection des tendances de la série.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Définit les tailles des bulles pour une série de graphiques.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Définit les valeurs d’une série de graphiques.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Définit les valeurs de l’axe X d’une série de graphiques.|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|Spécifie si la série possède une ombre.|
||[smooth](/javascript/api/excel/excel.chartseries#smooth)|Spécifie si la série est lisse.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Ajouter une nouvelle série à la collection.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Obtenir la sous-stration d’un titre de graphique.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Spécifie l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitle#left)|Spécifie la distance, en points, entre le bord gauche du titre du graphique et le bord gauche de la zone de graphique.|
||[position](/javascript/api/excel/excel.charttitle#position)|Représente la position du titre du graphique.|
||[height](/javascript/api/excel/excel.charttitle#height)|Représente la hauteur, exprimée en points, du titre du graphique.|
||[width](/javascript/api/excel/excel.charttitle#width)|Spécifie la largeur, en points, du titre du graphique.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Définit une valeur de chaîne qui représente la formule de titre de graphique à l’aide de la notation de style A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Spécifie l’angle vers lequel le texte est orienté pour le titre du graphique.|
||[top](/javascript/api/excel/excel.charttitle#top)|Spécifie la distance, en points, entre le bord supérieur du titre du graphique et le haut de la zone de graphique.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Spécifie l’alignement vertical du titre du graphique.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de trait et l’pondération.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Supprime l’objet courbe de tendance.|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|Représente la valeur intercept de la courbe de tendance.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Représente la période d’une courbe de tendance de graphique.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Représente le nom de la courbe de tendance.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Représente l’ordre d’une courbe de tendance de graphique.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Représente la mise en forme de courbe de tendance de graphique.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Représente le type de courbe de tendance de graphique.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Ajoute une nouvelle courbe de tendance à la collection de courbes de tendance.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Renvoie le nombre de courbes de tendance de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Obtient un objet courbe de tendance par index, c'est-à-dire par ordre d’insertion dans le tableau des éléments.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Représente le format des lignes du graphique.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.customproperty#key)|Clé de la propriété personnalisée.|
||[type](/javascript/api/excel/excel.customproperty#type)|Type de la valeur utilisée pour la propriété personnalisée.|
||[value](/javascript/api/excel/excel.customproperty#value)|Valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Actualise toutes les dataConnections de la collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#author)|Auteur du livre.|
||[category](/javascript/api/excel/excel.documentproperties#category)|Catégorie du classez.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Commentaires du workbook.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Société du workbook.|
||[keywords](/javascript/api/excel/excel.documentproperties#keywords)|Mots clés du workbook.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|Responsable du manuel.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Obtient la date de création du classeur.|
||[custom](/javascript/api/excel/excel.documentproperties#custom)|Obtient la collection de propriétés personnalisées du classeur.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Obtient ou définit le dernier auteur du classeur.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Obtient le numéro de révision du classeur.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Objet du manuel.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Titre du workbook.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Formule de l’élément nommé.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Renvoie un objet contenant les valeurs et les types de l’élément nommé.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Représente les types de chaque élément dans le tableau d’éléments nommés|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Représente les valeurs de chaque élément dans le tableau élément nommé.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Obtient un objet Plage avec la même cellule supérieure gauche que l’objet de Plage en cours, mais avec un nombre spécifié de lignes et colonnes.|
||[getImage()](/javascript/api/excel/excel.range#getimage--)|Restituer la plage en tant qu’image png codée en base 64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Renvoie un objet PLage qui représente la région environnante pour la cellule en haut à gauche de cette plage.|
||[lien hypertexte](/javascript/api/excel/excel.range#hyperlink)|Représente le lien hypertexte de la plage actuelle.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Représente le code Excel format numérique de la plage donnée, en fonction des paramètres de langue de l’utilisateur.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Représente si la plage active est une colonne entière.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Représente si la plage active est une ligne entière.|
||[showCard()](/javascript/api/excel/excel.range#showcard--)|Affiche la carte pour une cellule active si son contenu est riche en valeur.|
||[style](/javascript/api/excel/excel.range#style)|Représente le style de la plage actuelle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Orientation du texte de toutes les cellules de la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Spécifie si la largeur de colonne de l’objet Range est égale à la largeur standard de la feuille.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[adresse](/javascript/api/excel/excel.rangehyperlink#address)|Représente l’url cible du lien hypertexte.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Représente la cible de référence du document pour le lien hypertexte.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screentip)|Représente la chaîne affichée lorsque vous pointez sur le lien hypertexte.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Représente la chaîne qui s’affiche dans la cellule en haut à gauche de la plage.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Supprime ce style.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Spécifie si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Représente l’alignement horizontal pour le style.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Spécifie si le style inclut les propriétés AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel et TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Spécifie si le style inclut les propriétés de bordure Color, ColorIndex, LineStyle et Weight.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Spécifie si le style inclut les propriétés Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript et Underline font.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Spécifie si le style inclut la propriété NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Spécifie si le style inclut les propriétés d’intérieur Color, ColorIndex, InvertIfNegative, Pattern, PatternColor et PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Spécifie si le style inclut les propriétés de protection FormulaHidden et Locked.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[locked](/javascript/api/excel/excel.style#locked)|Spécifie si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|Le code de format du nombre format pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|Le code de format localisé du nombre format pour le style.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|L’ordre de lecture du style.|
||[Borders](/javascript/api/excel/excel.style#borders)|Collection de bordures de quatre objets qui représente le style des quatre bordures.|
||[builtIn](/javascript/api/excel/excel.style#builtin)|Spécifie si le style est un style intégré.|
||[fill](/javascript/api/excel/excel.style#fill)|Remplissage du style.|
||[police](/javascript/api/excel/excel.style#font)|Renvoie un objet Police qui représente la police du style.|
||[name](/javascript/api/excel/excel.style#name)|Nom du style.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Spécifie si le texte est automatiquement réduit pour tenir dans la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Spécifie l’alignement vertical du style.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Spécifie si Excel le texte dans l’objet.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Ajoute un nouveau style à la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Obtient un style par nom.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Se produit lorsque les données des cellules changent dans un tableau spécifique.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Se produit lorsque la sélection change dans un tableau spécifique.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[adresse](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtient l’adresse qui représente la zone modifiée d’un tableau figurant dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Obtient la source de l’événement.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Obtient l’id du tableau dans lequel les données sont modifiées.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Se produit lorsque des données changent dans une table d’un workbook ou d’une feuille de calcul.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone sélectionnée d’un tableau dans une feuille de calcul spécifique.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Spécifie si la sélection se trouve à l’intérieur d’un tableau, l’adresse sera inutile si IsInsideTable est false.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Obtient l’id du tableau dans lequel la sélection est modifiée.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Obtient la cellule active du classeur.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Représente toutes les connexions de données dans le workbook.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtient le nom du classeur.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtient les propriétés du classeur.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Renvoie l’objet de protection d’un workbook.|
||[styles](/javascript/api/excel/excel.workbook#styles)|Représente une collection de styles associés au classeur.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protège un classeur.|
||[protected](/javascript/api/excel/excel.workbookprotection#protected)|Spécifie si le workbook est protégé.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Annule la protection un classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Feuille de calcul)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copie une feuille de calcul et la place à la position spécifiée.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Obtient l’objet plage commençant à un index de ligne et de colonne particulier et couvrant un certain nombre de lignes et de colonnes.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Obtient un objet qui peut être utilisé pour manipuler des volets figés sur la feuille de calcul.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Se produit lorsque la feuille de calcul est activée.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Se produit lorsque des données sont modifiées dans une feuille de calcul spécifique.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Se produit lorsque la feuille de calcul est désactivée.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Se produit lorsque la sélection change dans une feuille de calcul spécifique.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Renvoie la hauteur standard (par défaut) de toutes les lignes dans la feuille de calcul, en points.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Spécifie la largeur standard (par défaut) de toutes les colonnes de la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|Couleur de l’onglet de la feuille de calcul.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est activée.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est ajoutée au classeur.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Se produit lorsqu’une feuille de calcul du manuel est activée.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Se produit lorsqu’une nouvelle feuille de calcul est ajoutée au manuel.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Se produit lorsqu’une feuille de calcul du manuel est désactivée.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Se produit lorsqu’une feuille de calcul est supprimée du manuel.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est desactivée.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est supprimée du classeur.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Figer la/les première(s) colonne(s) de la feuille de calcul en place.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Figer la/les première(s) ligne(s) de la feuille de calcul en place.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Supprime tous les volets figés dans la feuille de calcul.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Annule la protection d’une feuille de calcul.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Représente l’option de protection de feuille de calcul qui autorise la modification d’objets.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Représente l’option de protection de feuille de calcul qui autorise la modification de scénarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Représente l’option de protection de feuille de calcul qui autorise le mode sélection.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone sélectionnée dans une feuille de calcul spécifique.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
