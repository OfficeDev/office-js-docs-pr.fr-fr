---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,7
description: Détails sur l’ensemble de conditions requises ExcelApi 1,7.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ea1fe7a3d28acce2d1f4e9ff33f7b2bd31758fbd
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996234"
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
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Cette énumération spécifie le type de graphique.|
||[id](/javascript/api/excel/excel.chart#id)|ID unique du graphique.|
||[ShowAllFieldButtons,](/javascript/api/excel/excel.chart#showallfieldbuttons)|Indique si tous les boutons de champ d’un graphique croisé dynamique doivent être affichés.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[bordure](/javascript/api/excel/excel.chartareaformat#border)|Représente le format de bordure de la zone de graphique, qui inclut couleur, style de style et épaisseur.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (type : Excel. ChartAxisType, Group ?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Renvoie l’axe spécifique identifié par type et par groupe.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Spécifie l’unité de base pour l’axe des abscisses spécifié.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Spécifie le type d’axe de catégorie.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Représente l’unité d’affichage de l’axe.|
||[LogBase,](/javascript/api/excel/excel.chartaxis#logbase)|Indique la base du logarithme lors de l’utilisation d’échelles logarithmiques.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Cette énumération spécifie le type de graduation principale pour l’axe spécifié.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Spécifie la valeur d’échelle d’unité principale pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Cette énumération spécifie le type de graduation secondaire pour l’axe spécifié.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Indique la valeur d’échelle d’unité secondaire pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Cette énumération spécifie le groupe pour l’axe spécifié.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Spécifie la valeur de l’unité d’affichage de l’axe personnalisé.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Indique la hauteur, exprimée en points, de l’axe du graphique.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Indique la distance, en points, entre le bord gauche de l’axe et la partie gauche de la zone de graphique.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Indique la distance, en points, entre le bord supérieur de l’axe et le haut de la zone de graphique.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Spécifie le type d’axe.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Indique la largeur, exprimée en points, de l’axe du graphique.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Indique si Excel trace les points de données du dernier au premier.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Indique le type d’étendue de l’axe des ordonnées.|
||[setCategoryNames (sourceData : Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Définit tous les noms de catégorie pour l’axe spécifié.|
||[setCustomDisplayUnit (valeur : nombre)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Définit l’unité d’affichage axe sur une valeur personnalisée.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Indique si l’étiquette de l’unité d’affichage de l’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Spécifie la position des étiquettes de graduation sur l'axe spécifié.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Indique le nombre d’abscisses ou de séries entre les étiquettes de graduation.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Indique le nombre d’abscisses ou de séries entre les marques de graduation.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Indique si l’axe est visible.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Représente le style de trait de la bordure.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données.|
||[para](/javascript/api/excel/excel.chartdatalabel#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Indique si la taille de la bulle des étiquettes de données est visible.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Indique si le nom de catégorie de l’étiquette de données est visible.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Indique si la légende de l’étiquette de données est visible.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Indique si le pourcentage de l’étiquette de données est visible.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Indique si le nom de série des étiquettes de données est visible.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Indique si la valeur de l’étiquette de données est visible.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[police](/javascript/api/excel/excel.chartformatstring#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Indique la hauteur, exprimée en points, de la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Indique la gauche, en points, de la légende du graphique.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Représente une collection de legendEntries dans la légende.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Indique si la légende est ombrée sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Cette énumération spécifie le bord supérieur d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Indique la largeur, exprimée en points, de la légende du graphique.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Représente la partie visible d’une entrée de légende de graphique.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Renvoie le nombre de legendEntry de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Renvoie un legendEntry à l’index donné.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Représente le style de trait.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Indique si un point de données a une étiquette de données.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur du point de données (par exemple, #FF0000 représente le rouge).|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Représentation de code de couleur HTML de la couleur de premier plan du point de données (par exemple, #FF0000 représente le rouge).|
||[MarkerSize,](/javascript/api/excel/excel.chartpoint#markersize)|Représente la taille du marqueur du point de données.|
||[MarkerStyle,](/javascript/api/excel/excel.chartpoint#markerstyle)|Représente le style du marqueur du point de données de graphique.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Renvoie l’étiquette de données d’un point du graphique.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[bordure](/javascript/api/excel/excel.chartpointformat#border)|Représente le format de bordure d’un point de données de graphique, qui inclut les informations de couleur, de style et de poids.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Représente le type de graphique d’une série.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Supprime la série graphique.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Représente la taille du centre d’une série de graphiques en anneaux.|
||[filtré](/javascript/api/excel/excel.chartseries#filtered)|Indique si la série est filtrée.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Représente la largeur de l’intervalle d’une série de graphique.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Indique si la série possède des étiquettes de données.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Cette énumération spécifie la couleur d’arrière-plan des marqueurs d’une série de graphique.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Cette énumération spécifie la couleur de premier plan des marqueurs d’une série de graphiques.|
||[MarkerSize,](/javascript/api/excel/excel.chartseries#markersize)|Cette énumération spécifie la taille du marqueur d’une série de graphiques.|
||[MarkerStyle,](/javascript/api/excel/excel.chartseries#markerstyle)|Cette énumération spécifie le style de marqueur d’une série de graphiques.|
||[PlotOrder,](/javascript/api/excel/excel.chartseries#plotorder)|Cette énumération spécifie l’ordre de traçage d’une série de graphiques au sein du groupe de graphiques.|
||[Trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Collection de courbes de tendance de la série.|
||[setBubbleSizes (sourceData : Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Définit la taille des bulles pour une série de graphiques.|
||[SetValues (sourceData : Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Définit les valeurs d’une série de graphique.|
||[setXAxisValues (sourceData : Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Définit les valeurs de l’axe X pour une série de graphiques.|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|Indique si la série a une ombre.|
||[Unie](/javascript/api/excel/excel.chartseries#smooth)|Indique si la série est lisse.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Add (Name ?: String, index ?: Number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Ajouter une nouvelle série à la collection.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (début : nombre, longueur : nombre)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Obtenir la sous-chaîne d’un titre de graphique.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Spécifie l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitle#left)|Indique la distance, en points, entre le bord gauche du titre du graphique et le bord gauche de la zone de graphique.|
||[position](/javascript/api/excel/excel.charttitle#position)|Représente la position du titre du graphique.|
||[height](/javascript/api/excel/excel.charttitle#height)|Représente la hauteur, exprimée en points, du titre du graphique.|
||[width](/javascript/api/excel/excel.charttitle#width)|Indique la largeur, exprimée en points, du titre du graphique.|
||[setFormula (Formula : String)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Définit une valeur de chaîne qui représente la formule de titre de graphique à l’aide de la notation de style A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Spécifie l’angle auquel le texte est orienté pour le titre du graphique.|
||[top](/javascript/api/excel/excel.charttitle#top)|Indique la distance, en points, entre le bord supérieur du titre du graphique et le haut de la zone de graphique.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Indique l’alignement vertical du titre du graphique.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[bordure](/javascript/api/excel/excel.charttitleformat#border)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de style et le poids.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Supprime l’objet courbe de tendance.|
||[ordonn](/javascript/api/excel/excel.charttrendline#intercept)|Représente la valeur intercept de la courbe de tendance.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Représente la période d’une courbe de tendance de graphique.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Représente le nom de la courbe de tendance.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Représente l’ordre d’une courbe de tendance de graphique.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Représente la mise en forme de courbe de tendance de graphique.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Représente le type de courbe de tendance de graphique.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[Add (type ?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Ajoute une nouvelle courbe de tendance à la collection de courbes de tendance.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Renvoie le nombre de courbes de tendance de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Obtient un objet courbe de tendance par index, c'est-à-dire par ordre d’insertion dans le tableau des éléments.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Représente le format des lignes du graphique.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.customproperty#key)|Clé de la propriété personnalisée.|
||[type](/javascript/api/excel/excel.customproperty#type)|Type de la valeur utilisée pour la propriété personnalisée.|
||[value](/javascript/api/excel/excel.customproperty#value)|Valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[Add (Key : chaîne, value : any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[RefreshAll, ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Actualise toutes les dataConnections de la collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[créés](/javascript/api/excel/excel.documentproperties#author)|Auteur du classeur.|
||[catégories](/javascript/api/excel/excel.documentproperties#category)|Catégorie du classeur.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Commentaires du classeur.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Société du classeur.|
||[Mots clés](/javascript/api/excel/excel.documentproperties#keywords)|Mots clés du classeur.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|Gestionnaire du classeur.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Obtient la date de création du classeur.|
||[personnalisé](/javascript/api/excel/excel.documentproperties#custom)|Obtient la collection de propriétés personnalisées du classeur.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Obtient ou définit le dernier auteur du classeur.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Obtient le numéro de révision du classeur.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Objet du classeur.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Titre du classeur.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Formule de l’élément nommé.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Renvoie un objet contenant les valeurs et les types de l’élément nommé.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Représente les types de chaque élément dans le tableau d’éléments nommés|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Représente les valeurs de chaque élément dans le tableau élément nommé.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows : nombre, numColumns : nombre)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Obtient un objet Plage avec la même cellule supérieure gauche que l’objet de Plage en cours, mais avec un nombre spécifié de lignes et colonnes.|
||[getImage ()](/javascript/api/excel/excel.range#getimage--)|Affiche la plage en tant qu’image png encodée au format Base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Renvoie un objet PLage qui représente la région environnante pour la cellule en haut à gauche de cette plage.|
||[lien hypertexte](/javascript/api/excel/excel.range#hyperlink)|Représente le lien hypertexte de la plage active.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Représente le code de format de nombre d’Excel pour la plage donnée, en fonction des paramètres de langue de l’utilisateur.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Représente si la plage active est une colonne entière.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Représente si la plage active est une ligne entière.|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|Affiche la carte pour une cellule active si son contenu est riche en valeur.|
||[style](/javascript/api/excel/excel.range#style)|Représente le style de la plage actuelle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Orientation du texte de toutes les cellules dans la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Indique si la largeur de la colonne de l’objet Range est égale à la largeur standard de la feuille.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[adresse](/javascript/api/excel/excel.rangehyperlink#address)|Représente l’url cible du lien hypertexte.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Représente la cible de référence de document pour le lien hypertexte.|
||[Info](/javascript/api/excel/excel.rangehyperlink#screentip)|Représente la chaîne affichée lorsque vous pointez sur le lien hypertexte.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Représente la chaîne qui s’affiche dans la cellule en haut à gauche de la plage.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Supprime ce style.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Indique si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Représente l’alignement horizontal pour le style.|
||[IncludeAlignment,](/javascript/api/excel/excel.style#includealignment)|Indique si le style inclut les propriétés autoindent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel et TextOrientation.|
||[IncludeBorder,](/javascript/api/excel/excel.style#includeborder)|Indique si le style inclut les propriétés Color, ColorIndex, LineStyle et Weight bordure.|
||[IncludeFont,](/javascript/api/excel/excel.style#includefont)|Indique si le style inclut les propriétés de police arrière-plan, gras, couleur, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript et Underline.|
||[IncludeNumber,](/javascript/api/excel/excel.style#includenumber)|Indique si le style inclut la propriété NumberFormat.|
||[IncludePatterns,](/javascript/api/excel/excel.style#includepatterns)|Indique si le style inclut les propriétés Color, ColorIndex, InvertIfNegative, pattern, PatternColor et PatternColorIndex, Interior.|
||[IncludeProtection,](/javascript/api/excel/excel.style#includeprotection)|Indique si le style inclut les propriétés FormulaHidden et protection verrouillée.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[locked](/javascript/api/excel/excel.style#locked)|Indique si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|Le code de format du nombre format pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|Le code de format localisé du nombre format pour le style.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|L’ordre de lecture du style.|
||[Borders](/javascript/api/excel/excel.style#borders)|Collection de bordures de quatre objets qui représente le style des quatre bordures.|
||[builtIn](/javascript/api/excel/excel.style#builtin)|Indique si le style est un style prédéfini.|
||[fill](/javascript/api/excel/excel.style#fill)|Remplissage du style.|
||[police](/javascript/api/excel/excel.style#font)|Renvoie un objet Police qui représente la police du style.|
||[name](/javascript/api/excel/excel.style#name)|Nom du style.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Indique si le texte s’ajuste automatiquement à la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Spécifie l’alignement vertical du style.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Indique si Excel renvoie le texte à la ligne dans l’objet.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Ajoute un nouveau style à la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Obtient un style par nom.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Survient lorsque les données des cellules changent sur une table spécifique.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Se produit lorsque la sélection change sur une table spécifique.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[adresse](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtient l’adresse qui représente la zone modifiée d’un tableau figurant dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Obtient la source de l’événement.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Obtient l’id du tableau dans lequel les données sont modifiées.|
||[type](/javascript/api/excel/excel.tablechangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Survient lors de la modification des données d’une table dans un classeur ou d’une feuille de calcul.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone sélectionnée d’un tableau dans une feuille de calcul spécifique.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Indique si la sélection se trouve dans un tableau, l’adresse est inutile si IsInsideTable a la valeur false.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Obtient l’id du tableau dans lequel la sélection est modifiée.|
||[type](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Obtient la cellule active du classeur.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Représente toutes les connexions de données du classeur.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtient le nom du classeur.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtient les propriétés du classeur.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Renvoie l’objet protection pour un classeur.|
||[proposés](/javascript/api/excel/excel.workbook#styles)|Représente une collection de styles associés au classeur.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[Protect (Password ?: String)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protège un classeur.|
||[sécurisé](/javascript/api/excel/excel.workbookprotection#protected)|Indique si le classeur est protégé.|
||[Unprotect (Password ?: String)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Annule la protection un classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (positionType ?: Excel. WorksheetPositionType, relativeTo ?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copie une feuille de calcul et la place à la position spécifiée.|
||[getRangeByIndexes (startRow : nombre, ColonneDébut : nombre, rowCount : nombre, columnCount : nombre)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Obtient l’objet plage commençant à un index de ligne et de colonne particulier et couvrant un certain nombre de lignes et de colonnes.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Obtient un objet qui peut être utilisé pour manipuler les volets figés de la feuille de calcul.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Se produit lorsque la feuille de calcul est activée.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Se produit lorsque des données sont modifiées dans une feuille de calcul spécifique.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Se produit lorsque la feuille de calcul est désactivée.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Se produit lorsque la sélection change dans une feuille de calcul spécifique.|
||[StandardHeight,](/javascript/api/excel/excel.worksheet#standardheight)|Renvoie la hauteur standard (par défaut) de toutes les lignes dans la feuille de calcul, en points.|
||[StandardWidth,](/javascript/api/excel/excel.worksheet#standardwidth)|Spécifie la largeur standard (par défaut) de toutes les colonnes de la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|Couleur d’onglet de la feuille de calcul.|
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
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est activée.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Survient lors de l’ajout d’une nouvelle feuille de calcul au classeur.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est désactivée.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Survient lors de la suppression d’une feuille de calcul du classeur.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est desactivée.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtient le type de l’événement.|
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
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
