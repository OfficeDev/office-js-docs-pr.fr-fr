---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,7
description: Détails sur l’ensemble de conditions requises ExcelApi 1,7
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c84d099982225bae11cb3deba8a0503da0695aed
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771987"
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

Les API Événements pour Excel fournissent un grand nombre de gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. Pour une liste des événements qui sont actuellement disponibles, voir [Manipuler des Événements à l’aide de l’API JavaScript Excel](/office/dev/add-ins/excel/excel-add-ins-events).

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

| Class | Champs | Description |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Représente le type d’un graphique. Pour plus d’informations, voir Excel. ChartType.|
||[id](/javascript/api/excel/excel.chart#id)|ID unique du graphique. En lecture seule.|
||[ShowAllFieldButtons,](/javascript/api/excel/excel.chart#showallfieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[route](/javascript/api/excel/excel.chartareaformat#border)|Représente le format de bordure de la zone de graphique, qui inclut couleur, style de style et épaisseur. En lecture seule.|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[route](/javascript/api/excel/excel.chartareaformatdata#border)|Représente le format de bordure de la zone de graphique, qui inclut couleur, style de style et épaisseur. En lecture seule.|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[route](/javascript/api/excel/excel.chartareaformatloadoptions#border)|Représente le format de bordure de la zone de graphique, qui inclut couleur, style de style et épaisseur.|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[route](/javascript/api/excel/excel.chartareaformatupdatedata#border)|Représente le format de bordure de la zone de graphique, qui inclut couleur, style de style et épaisseur.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (type: "non valide \| " "Category \| " "value \| " "Series", Group?: "Primary \| " "Secondary")](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Renvoie l’axe spécifique identifié par type et par groupe.|
||[getItem (type: Excel. ChartAxisType, Group?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Renvoie l’axe spécifique identifié par type et par groupe.|
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
||[setCategoryNames (sourceData: Range)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Définit tous les noms de catégorie pour l’axe spécifié.|
||[setCustomDisplayUnit (valeur: nombre)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Définit l’unité d’affichage axe sur une valeur personnalisée.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Représente si l’étiquette d’unité de l’affichage d’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Représentant la position des étiquettes de graduation sur l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickLabelPosition.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Représente le nombre de catégories ou séries entre les étiquettes de graduation. Peut être une valeur de 1 à 31999 ou une chaîne vide pour le paramètre automatique. La valeur renvoyée est toujours un nombre.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Représente le nombre de catégories ou séries entre les graduations.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Valeur booléenne qui représente la visibilité d’un axe.|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[axisGroup](/javascript/api/excel/excel.chartaxisdata#axisgroup)|Représente le groupe pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisGroup. En lecture seule.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisdata#basetimeunit)|Renvoie ou définit l’unité de base pour l’axe des abscisses spécifié.|
||[categoryType](/javascript/api/excel/excel.chartaxisdata#categorytype)|Renvoie ou définit le type d’axe de catégorie.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisdata#customdisplayunit)|Représente la valeur unité d’affichage personnalisé d’axe. En lecture seule. Pour définir cette propriété, utilisez la méthode SetCustomDisplayUnit(double).|
||[displayUnit](/javascript/api/excel/excel.chartaxisdata#displayunit)|Représente l’unité d’affichage de l’axe. Pour plus d’informations, voir Excel. ChartAxisDisplayUnit.|
||[height](/javascript/api/excel/excel.chartaxisdata#height)|Représente la hauteur, exprimée en points, de l’axe de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[left](/javascript/api/excel/excel.chartaxisdata#left)|Représente la distance en points, du bord gauche de l’axe au bord gauche de la zone de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[LogBase,](/javascript/api/excel/excel.chartaxisdata#logbase)|Représente la base du logarithme lorsque vous utilisez des échelles logarithmiques.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisdata#majortickmark)|Représente le type de graduation principale pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#majortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité principale pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisdata#minortickmark)|Représente le type de graduation secondaire pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#minortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité secondaire pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisdata#reverseplotorder)|Représente si Microsoft Excel trace des points de données à du dernier au premier.|
||[scaleType](/javascript/api/excel/excel.chartaxisdata#scaletype)|Représente le type d’échelle de l’axe des ordonnées. Pour plus d’informations, voir Excel. ChartAxisScaleType.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisdata#showdisplayunitlabel)|Représente si l’étiquette d’unité de l’affichage d’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisdata#ticklabelposition)|Représentant la position des étiquettes de graduation sur l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickLabelPosition.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisdata#ticklabelspacing)|Représente le nombre de catégories ou séries entre les étiquettes de graduation. Peut être une valeur de 1 à 31999 ou une chaîne vide pour le paramètre automatique. La valeur renvoyée est toujours un nombre.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisdata#tickmarkspacing)|Représente le nombre de catégories ou séries entre les graduations.|
||[top](/javascript/api/excel/excel.chartaxisdata#top)|Représente la distance en points, du bord supérieur de l’axe au bord supérieur de la zone de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[type](/javascript/api/excel/excel.chartaxisdata#type)|Représente le type d’axe. Pour plus d’informations, voir Excel. ChartAxisType.|
||[visible](/javascript/api/excel/excel.chartaxisdata#visible)|Valeur booléenne qui représente la visibilité d’un axe.|
||[width](/javascript/api/excel/excel.chartaxisdata#width)|Représente la largeur, en points, de l’axe de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[axisGroup](/javascript/api/excel/excel.chartaxisloadoptions#axisgroup)|Représente le groupe pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisGroup. En lecture seule.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisloadoptions#basetimeunit)|Renvoie ou définit l’unité de base pour l’axe des abscisses spécifié.|
||[categoryType](/javascript/api/excel/excel.chartaxisloadoptions#categorytype)|Renvoie ou définit le type d’axe de catégorie.|
||[dépasse](/javascript/api/excel/excel.chartaxisloadoptions#crosses)|[Déconseillé; conservé pour la compatibilité descendante avec les solutions existantes de première partie]. Utilisez `Position` à la place.|
||[crossesAt](/javascript/api/excel/excel.chartaxisloadoptions#crossesat)|[Déconseillé; conservé pour la compatibilité descendante avec les solutions existantes de première partie]. Utilisez `PositionAt` à la place.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisloadoptions#customdisplayunit)|Représente la valeur unité d’affichage personnalisé d’axe. En lecture seule. Pour définir cette propriété, utilisez la méthode SetCustomDisplayUnit(double).|
||[displayUnit](/javascript/api/excel/excel.chartaxisloadoptions#displayunit)|Représente l’unité d’affichage de l’axe. Pour plus d’informations, voir Excel. ChartAxisDisplayUnit.|
||[height](/javascript/api/excel/excel.chartaxisloadoptions#height)|Représente la hauteur, exprimée en points, de l’axe de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[left](/javascript/api/excel/excel.chartaxisloadoptions#left)|Représente la distance en points, du bord gauche de l’axe au bord gauche de la zone de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[LogBase,](/javascript/api/excel/excel.chartaxisloadoptions#logbase)|Représente la base du logarithme lorsque vous utilisez des échelles logarithmiques.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#majortickmark)|Représente le type de graduation principale pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#majortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité principale pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#minortickmark)|Représente le type de graduation secondaire pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#minortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité secondaire pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisloadoptions#reverseplotorder)|Représente si Microsoft Excel trace des points de données à du dernier au premier.|
||[scaleType](/javascript/api/excel/excel.chartaxisloadoptions#scaletype)|Représente le type d’échelle de l’axe des ordonnées. Pour plus d’informations, voir Excel. ChartAxisScaleType.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisloadoptions#showdisplayunitlabel)|Représente si l’étiquette d’unité de l’affichage d’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelposition)|Représentant la position des étiquettes de graduation sur l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickLabelPosition.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelspacing)|Représente le nombre de catégories ou séries entre les étiquettes de graduation. Peut être une valeur de 1 à 31999 ou une chaîne vide pour le paramètre automatique. La valeur renvoyée est toujours un nombre.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisloadoptions#tickmarkspacing)|Représente le nombre de catégories ou séries entre les graduations.|
||[top](/javascript/api/excel/excel.chartaxisloadoptions#top)|Représente la distance en points, du bord supérieur de l’axe au bord supérieur de la zone de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
||[type](/javascript/api/excel/excel.chartaxisloadoptions#type)|Représente le type d’axe. Pour plus d’informations, voir Excel. ChartAxisType.|
||[visible](/javascript/api/excel/excel.chartaxisloadoptions#visible)|Valeur booléenne qui représente la visibilité d’un axe.|
||[width](/javascript/api/excel/excel.chartaxisloadoptions#width)|Représente la largeur, en points, de l’axe de graphique. NULL si l’axe n’est pas visible. En lecture seule.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[baseTimeUnit](/javascript/api/excel/excel.chartaxisupdatedata#basetimeunit)|Renvoie ou définit l’unité de base pour l’axe des abscisses spécifié.|
||[categoryType](/javascript/api/excel/excel.chartaxisupdatedata#categorytype)|Renvoie ou définit le type d’axe de catégorie.|
||[displayUnit](/javascript/api/excel/excel.chartaxisupdatedata#displayunit)|Représente l’unité d’affichage de l’axe. Pour plus d’informations, voir Excel. ChartAxisDisplayUnit.|
||[LogBase,](/javascript/api/excel/excel.chartaxisupdatedata#logbase)|Représente la base du logarithme lorsque vous utilisez des échelles logarithmiques.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#majortickmark)|Représente le type de graduation principale pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#majortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité principale pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#minortickmark)|Représente le type de graduation secondaire pour l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickMark.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#minortimeunitscale)|Renvoie ou définit la valeur d’échelle d’unité secondaire pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisupdatedata#reverseplotorder)|Représente si Microsoft Excel trace des points de données à du dernier au premier.|
||[scaleType](/javascript/api/excel/excel.chartaxisupdatedata#scaletype)|Représente le type d’échelle de l’axe des ordonnées. Pour plus d’informations, voir Excel. ChartAxisScaleType.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisupdatedata#showdisplayunitlabel)|Représente si l’étiquette d’unité de l’affichage d’axe est visible.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelposition)|Représentant la position des étiquettes de graduation sur l’axe spécifié. Pour plus d’informations, voir Excel. ChartAxisTickLabelPosition.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelspacing)|Représente le nombre de catégories ou séries entre les étiquettes de graduation. Peut être une valeur de 1 à 31999 ou une chaîne vide pour le paramètre automatique. La valeur renvoyée est toujours un nombre.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisupdatedata#tickmarkspacing)|Représente le nombre de catégories ou séries entre les graduations.|
||[visible](/javascript/api/excel/excel.chartaxisupdatedata#visible)|Valeur booléenne qui représente la visibilité d’un axe.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Représente le style de trait de la bordure. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[Set (propriétés: Excel. ChartBorder)](/javascript/api/excel/excel.chartborder#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartBorderUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartborder#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartBorderData](/javascript/api/excel/excel.chartborderdata)|[color](/javascript/api/excel/excel.chartborderdata#color)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborderdata#linestyle)|Représente le style de trait de la bordure. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartborderdata#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartBorderLoadOptions](/javascript/api/excel/excel.chartborderloadoptions)|[$all](/javascript/api/excel/excel.chartborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartborderloadoptions#color)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborderloadoptions#linestyle)|Représente le style de trait de la bordure. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartborderloadoptions#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartBorderUpdateData](/javascript/api/excel/excel.chartborderupdatedata)|[color](/javascript/api/excel/excel.chartborderupdatedata#color)|Code couleur HTML qui représente la couleur des bordures dans le graphique.|
||[lineStyle](/javascript/api/excel/excel.chartborderupdatedata#linestyle)|Représente le style de trait de la bordure. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartborderupdatedata#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartcollectionloadoptions#charttype)|Pour chaque élément de la collection: représente le type du graphique. Pour plus d’informations, voir Excel. ChartType.|
||[id](/javascript/api/excel/excel.chartcollectionloadoptions#id)|Pour chaque élément de la collection: ID unique du graphique. En lecture seule.|
||[ShowAllFieldButtons,](/javascript/api/excel/excel.chartcollectionloadoptions#showallfieldbuttons)|Pour chaque élément de la collection: indique si tous les boutons de champ d’un graphique croisé dynamique doivent être affichés.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[chartType](/javascript/api/excel/excel.chartdata#charttype)|Représente le type d’un graphique. Pour plus d’informations, voir Excel. ChartType.|
||[id](/javascript/api/excel/excel.chartdata#id)|ID unique du graphique. En lecture seule.|
||[ShowAllFieldButtons,](/javascript/api/excel/excel.chartdata#showallfieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[para](/javascript/api/excel/excel.chartdatalabel#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[Set (propriétés: Excel. ChartDataLabel)](/javascript/api/excel/excel.chartdatalabel#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartDataLabelUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartdatalabel#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[position](/javascript/api/excel/excel.chartdatalabeldata#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[para](/javascript/api/excel/excel.chartdatalabeldata#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabeldata#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabeldata#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabeldata#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabeldata#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabeldata#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabeldata#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelloadoptions#$all)||
||[position](/javascript/api/excel/excel.chartdatalabelloadoptions#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[para](/javascript/api/excel/excel.chartdatalabelloadoptions#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelloadoptions#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelloadoptions#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabelloadoptions#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelloadoptions#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelloadoptions#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabelloadoptions#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[position](/javascript/api/excel/excel.chartdatalabelupdatedata#position)|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Pour plus d’informations, voir Excel. ChartDataLabelPosition.|
||[para](/javascript/api/excel/excel.chartdatalabelupdatedata#separator)|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelupdatedata#showbubblesize)|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelupdatedata#showcategoryname)|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|
||[ShowLegendKey,](/javascript/api/excel/excel.chartdatalabelupdatedata#showlegendkey)|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelupdatedata#showpercentage)|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelupdatedata#showseriesname)|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|
||[showValue](/javascript/api/excel/excel.chartdatalabelupdatedata#showvalue)|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[police](/javascript/api/excel/excel.chartformatstring#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de graphique.|
||[Set (propriétés: Excel. ChartFormatString)](/javascript/api/excel/excel.chartformatstring#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartFormatStringUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartformatstring#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartFormatStringData](/javascript/api/excel/excel.chartformatstringdata)|[police](/javascript/api/excel/excel.chartformatstringdata#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de graphique.|
|[ChartFormatStringLoadOptions](/javascript/api/excel/excel.chartformatstringloadoptions)|[$all](/javascript/api/excel/excel.chartformatstringloadoptions#$all)||
||[police](/javascript/api/excel/excel.chartformatstringloadoptions#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de graphique.|
|[ChartFormatStringUpdateData](/javascript/api/excel/excel.chartformatstringupdatedata)|[police](/javascript/api/excel/excel.chartformatstringupdatedata#font)|Représente les attributs de police, tels que le nom de la police, la taille de police, la couleur, etc., de l’objet de graphique.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Représente la hauteur, en points, de la légende du graphique. NULL si la légende n’est pas visible.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Représente la gauche, en points, d’une légende de graphique. NULL si la légende n’est pas visible.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Représente une collection de legendEntries dans la légende. En lecture seule.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showshadow)|Représente si la légende est ombrée sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Représente la partie supérieure d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Représente la largeur, exprimée en points, de la légende du graphique. NULL si la légende n’est pas visible.|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[height](/javascript/api/excel/excel.chartlegenddata#height)|Représente la hauteur, en points, de la légende du graphique. NULL si la légende n’est pas visible.|
||[left](/javascript/api/excel/excel.chartlegenddata#left)|Représente la gauche, en points, d’une légende de graphique. NULL si la légende n’est pas visible.|
||[legendEntries](/javascript/api/excel/excel.chartlegenddata#legendentries)|Représente une collection de legendEntries dans la légende. En lecture seule.|
||[showShadow](/javascript/api/excel/excel.chartlegenddata#showshadow)|Représente si la légende est ombrée sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegenddata#top)|Représente la partie supérieure d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegenddata#width)|Représente la largeur, exprimée en points, de la légende du graphique. NULL si la légende n’est pas visible.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[Set (propriétés: Excel. ChartLegendEntry)](/javascript/api/excel/excel.chartlegendentry#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartLegendEntryUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartlegendentry#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Représente la partie visible d’une entrée de légende de graphique.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Renvoie le nombre de legendEntry de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Renvoie un legendEntry à l’index donné.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#visible)|Pour chaque élément de la collection: représente l’élément visible d’une entrée de légende de graphique.|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[visible](/javascript/api/excel/excel.chartlegendentrydata#visible)|Représente la partie visible d’une entrée de légende de graphique.|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentryloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentryloadoptions#visible)|Représente la partie visible d’une entrée de légende de graphique.|
|[ChartLegendEntryUpdateData](/javascript/api/excel/excel.chartlegendentryupdatedata)|[visible](/javascript/api/excel/excel.chartlegendentryupdatedata#visible)|Représente la partie visible d’une entrée de légende de graphique.|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[height](/javascript/api/excel/excel.chartlegendloadoptions#height)|Représente la hauteur, en points, de la légende du graphique. NULL si la légende n’est pas visible.|
||[left](/javascript/api/excel/excel.chartlegendloadoptions#left)|Représente la gauche, en points, d’une légende de graphique. NULL si la légende n’est pas visible.|
||[showShadow](/javascript/api/excel/excel.chartlegendloadoptions#showshadow)|Représente si la légende est ombrée sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegendloadoptions#top)|Représente la partie supérieure d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegendloadoptions#width)|Représente la largeur, exprimée en points, de la légende du graphique. NULL si la légende n’est pas visible.|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[height](/javascript/api/excel/excel.chartlegendupdatedata#height)|Représente la hauteur, en points, de la légende du graphique. NULL si la légende n’est pas visible.|
||[left](/javascript/api/excel/excel.chartlegendupdatedata#left)|Représente la gauche, en points, d’une légende de graphique. NULL si la légende n’est pas visible.|
||[showShadow](/javascript/api/excel/excel.chartlegendupdatedata#showshadow)|Représente si la légende est ombrée sur le graphique.|
||[top](/javascript/api/excel/excel.chartlegendupdatedata#top)|Représente la partie supérieure d’une légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegendupdatedata#width)|Représente la largeur, exprimée en points, de la légende du graphique. NULL si la légende n’est pas visible.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Représente le style de trait. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[lineStyle](/javascript/api/excel/excel.chartlineformatdata#linestyle)|Représente le style de trait. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartlineformatdata#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[lineStyle](/javascript/api/excel/excel.chartlineformatloadoptions#linestyle)|Représente le style de trait. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartlineformatloadoptions#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[lineStyle](/javascript/api/excel/excel.chartlineformatupdatedata#linestyle)|Représente le style de trait. Pour plus d’informations, voir Excel. ChartLineStyle.|
||[weight](/javascript/api/excel/excel.chartlineformatupdatedata#weight)|Représente l’épaisseur de bordure, en points.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[chartType](/javascript/api/excel/excel.chartloadoptions#charttype)|Représente le type d’un graphique. Pour plus d’informations, voir Excel. ChartType.|
||[id](/javascript/api/excel/excel.chartloadoptions#id)|ID unique du graphique. En lecture seule.|
||[ShowAllFieldButtons,](/javascript/api/excel/excel.chartloadoptions#showallfieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Indique si un point de données a une étiquette de données. Non applicable pour les graphiques en surface.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Représentation de code de couleur HTML de la couleur de premier plan du point de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerSize,](/javascript/api/excel/excel.chartpoint#markersize)|Représente la taille du marqueur du point de données.|
||[MarkerStyle,](/javascript/api/excel/excel.chartpoint#markerstyle)|Représente le style du marqueur du point de données de graphique. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Renvoie l’étiquette de données d’un point du graphique. En lecture seule.|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[dataLabel](/javascript/api/excel/excel.chartpointdata#datalabel)|Renvoie l’étiquette de données d’un point du graphique. En lecture seule.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointdata#hasdatalabel)|Indique si un point de données a une étiquette de données. Non applicable pour les graphiques en surface.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartpointdata#markerbackgroundcolor)|Représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartpointdata#markerforegroundcolor)|Représentation de code de couleur HTML de la couleur de premier plan du point de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerSize,](/javascript/api/excel/excel.chartpointdata#markersize)|Représente la taille du marqueur du point de données.|
||[MarkerStyle,](/javascript/api/excel/excel.chartpointdata#markerstyle)|Représente le style du marqueur du point de données de graphique. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[route](/javascript/api/excel/excel.chartpointformat#border)|Représente le format de bordure d’un point de données de graphique, qui inclut les informations de couleur, de style et de poids. En lecture seule.|
|[ChartPointFormatData](/javascript/api/excel/excel.chartpointformatdata)|[route](/javascript/api/excel/excel.chartpointformatdata#border)|Représente le format de bordure d’un point de données de graphique, qui inclut les informations de couleur, de style et de poids. En lecture seule.|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[route](/javascript/api/excel/excel.chartpointformatloadoptions#border)|Représente le format de bordure d’un point de données de graphique, qui inclut les informations de couleur, de style et de poids.|
|[ChartPointFormatUpdateData](/javascript/api/excel/excel.chartpointformatupdatedata)|[route](/javascript/api/excel/excel.chartpointformatupdatedata#border)|Représente le format de bordure d’un point de données de graphique, qui inclut les informations de couleur, de style et de poids.|
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointloadoptions#datalabel)|Renvoie l’étiquette de données d’un point du graphique.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointloadoptions#hasdatalabel)|Indique si un point de données a une étiquette de données. Non applicable pour les graphiques en surface.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartpointloadoptions#markerbackgroundcolor)|Représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartpointloadoptions#markerforegroundcolor)|Représentation de code de couleur HTML de la couleur de premier plan du point de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerSize,](/javascript/api/excel/excel.chartpointloadoptions#markersize)|Représente la taille du marqueur du point de données.|
||[MarkerStyle,](/javascript/api/excel/excel.chartpointloadoptions#markerstyle)|Représente le style du marqueur du point de données de graphique. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[dataLabel](/javascript/api/excel/excel.chartpointupdatedata#datalabel)|Renvoie l’étiquette de données d’un point du graphique.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointupdatedata#hasdatalabel)|Indique si un point de données a une étiquette de données. Non applicable pour les graphiques en surface.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartpointupdatedata#markerbackgroundcolor)|Représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartpointupdatedata#markerforegroundcolor)|Représentation de code de couleur HTML de la couleur de premier plan du point de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerSize,](/javascript/api/excel/excel.chartpointupdatedata#markersize)|Représente la taille du marqueur du point de données.|
||[MarkerStyle,](/javascript/api/excel/excel.chartpointupdatedata#markerstyle)|Représente le style du marqueur du point de données de graphique. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#datalabel)|Pour chaque élément de la collection: renvoie l’étiquette de données d’un point de graphique.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#hasdatalabel)|Pour chaque élément de la collection: indique si un point de données a une étiquette de données. Non applicable pour les graphiques en surface.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerbackgroundcolor)|Pour chaque élément de la collection: représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur du point de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerforegroundcolor)|Pour chaque élément de la collection: représentation de code de couleur HTML de la couleur de premier plan du point de données. Par exemple, #FF0000 représente le rouge.|
||[MarkerSize,](/javascript/api/excel/excel.chartpointscollectionloadoptions#markersize)|Pour chaque élément de la collection: représente la taille du marqueur du point de données.|
||[MarkerStyle,](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerstyle)|Pour chaque élément de la collection: représente le style de marqueur d’un point de données de graphique. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
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
||[setBubbleSizes (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Définit des tailles de bulles pour une série de graphiques. Fonctionne uniquement pour les graphiques en bulles.|
||[SetValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Définit des valeurs pour une série de graphiques. Pour un graphique en nuages de points, cela signifie des valeurs de l’axe Y.|
||[setXAxisValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Définit des valeurs d’axe Y pour une série de graphiques. Fonctionne uniquement pour les graphiques en nuages de points.|
||[showShadow](/javascript/api/excel/excel.chartseries#showshadow)|Valeur booléenne représentant si la série a une ombre ou non.|
||[Unie](/javascript/api/excel/excel.chartseries#smooth)|Valeur booléenne représentant si la série est fluide ou non. Applicable uniquement aux graphiques en courbes et à nuages de points.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Add (Name?: String, index?: Number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Ajouter une nouvelle série à la collection. La nouvelle série ajoutée n’est pas visible jusqu’à ce que les valeurs de l’axe des ordonnées/valeurs de l’axe x/tailles des bulles lui soient attribuées (en fonction du type de graphique).|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartseriescollectionloadoptions#charttype)|Pour chaque élément de la collection: représente le type de graphique d’une série. Pour plus d’informations, voir Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#doughnutholesize)|Pour chaque élément de la collection: représente la taille du centre d’une série de graphiques.  Valide uniquement dans les graphiques en anneaux et doughnutExploded.|
||[filtré](/javascript/api/excel/excel.chartseriescollectionloadoptions#filtered)|Pour chaque élément de la collection: valeur booléenne représentant si la série est filtrée ou non. Non applicable pour les graphiques en surface.|
||[gapWidth](/javascript/api/excel/excel.chartseriescollectionloadoptions#gapwidth)|Pour chaque élément de la collection: représente la largeur de l’intervalle d’une série de graphique.  Valide uniquement sur les graphiques en barres et colonnes, ainsi que|
||[hasDataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#hasdatalabels)|Pour chaque élément de la collection: valeur booléenne représentant si la série possède des étiquettes de données ou non.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerbackgroundcolor)|Pour chaque élément de la collection: représente la couleur d’arrière-plan des marqueurs d’une série de graphique.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerforegroundcolor)|Pour chaque élément de la collection: représente la couleur de premier plan des marqueurs d’une série de graphiques.|
||[MarkerSize,](/javascript/api/excel/excel.chartseriescollectionloadoptions#markersize)|Pour chaque élément de la collection: représente la taille du marqueur d’une série de graphique.|
||[MarkerStyle,](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerstyle)|Pour chaque élément de la collection: représente le style de marqueur d’une série de graphique. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
||[PlotOrder,](/javascript/api/excel/excel.chartseriescollectionloadoptions#plotorder)|Pour chaque élément de la collection: représente l’ordre de traçage d’une série de graphiques au sein du groupe de graphiques.|
||[showShadow](/javascript/api/excel/excel.chartseriescollectionloadoptions#showshadow)|Pour chaque élément de la collection: valeur booléenne représentant si la série a une ombre ou non.|
||[Unie](/javascript/api/excel/excel.chartseriescollectionloadoptions#smooth)|Pour chaque élément de la collection: valeur booléenne représentant si la série est lisse ou non. Applicable uniquement aux graphiques en courbes et à nuages de points.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[chartType](/javascript/api/excel/excel.chartseriesdata#charttype)|Représente le type de graphique d’une série. Pour plus d’informations, voir Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesdata#doughnutholesize)|Représente la taille du centre d’une série de graphiques en anneaux.  Valide uniquement dans les graphiques en anneaux et doughnutExploded.|
||[filtré](/javascript/api/excel/excel.chartseriesdata#filtered)|Valeur booléenne représentant si la série est filtrée ou non. Non applicable pour les graphiques en surface.|
||[gapWidth](/javascript/api/excel/excel.chartseriesdata#gapwidth)|Représente la largeur de l’intervalle d’une série de graphique.  Valide uniquement sur les graphiques en barres et colonnes, ainsi que|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesdata#hasdatalabels)|Valeur booléenne représentant si la série a des étiquettes de données ou non.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartseriesdata#markerbackgroundcolor)|Représente la couleur d’arrière-plan de marqueurs d’une série de graphiques.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartseriesdata#markerforegroundcolor)|Représente la couleur de premier plan de marqueurs d’une série de graphiques.|
||[MarkerSize,](/javascript/api/excel/excel.chartseriesdata#markersize)|Représente la taille du marqueur d’une série de graphiques.|
||[MarkerStyle,](/javascript/api/excel/excel.chartseriesdata#markerstyle)|Représente le style du marqueur d’une série de graphiques. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
||[PlotOrder,](/javascript/api/excel/excel.chartseriesdata#plotorder)|Représente l’ordre de traçage d’une série de graphiques au sein du groupe de graphiques.|
||[showShadow](/javascript/api/excel/excel.chartseriesdata#showshadow)|Valeur booléenne représentant si la série a une ombre ou non.|
||[Unie](/javascript/api/excel/excel.chartseriesdata#smooth)|Valeur booléenne représentant si la série est fluide ou non. Applicable uniquement aux graphiques en courbes et à nuages de points.|
||[Trendlines](/javascript/api/excel/excel.chartseriesdata#trendlines)|Représente la collection de courbes de tendance de la série. En lecture seule.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[chartType](/javascript/api/excel/excel.chartseriesloadoptions#charttype)|Représente le type de graphique d’une série. Pour plus d’informations, voir Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesloadoptions#doughnutholesize)|Représente la taille du centre d’une série de graphiques en anneaux.  Valide uniquement dans les graphiques en anneaux et doughnutExploded.|
||[filtré](/javascript/api/excel/excel.chartseriesloadoptions#filtered)|Valeur booléenne représentant si la série est filtrée ou non. Non applicable pour les graphiques en surface.|
||[gapWidth](/javascript/api/excel/excel.chartseriesloadoptions#gapwidth)|Représente la largeur de l’intervalle d’une série de graphique.  Valide uniquement sur les graphiques en barres et colonnes, ainsi que|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesloadoptions#hasdatalabels)|Valeur booléenne représentant si la série a des étiquettes de données ou non.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartseriesloadoptions#markerbackgroundcolor)|Représente la couleur d’arrière-plan de marqueurs d’une série de graphiques.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartseriesloadoptions#markerforegroundcolor)|Représente la couleur de premier plan de marqueurs d’une série de graphiques.|
||[MarkerSize,](/javascript/api/excel/excel.chartseriesloadoptions#markersize)|Représente la taille du marqueur d’une série de graphiques.|
||[MarkerStyle,](/javascript/api/excel/excel.chartseriesloadoptions#markerstyle)|Représente le style du marqueur d’une série de graphiques. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
||[PlotOrder,](/javascript/api/excel/excel.chartseriesloadoptions#plotorder)|Représente l’ordre de traçage d’une série de graphiques au sein du groupe de graphiques.|
||[showShadow](/javascript/api/excel/excel.chartseriesloadoptions#showshadow)|Valeur booléenne représentant si la série a une ombre ou non.|
||[Unie](/javascript/api/excel/excel.chartseriesloadoptions#smooth)|Valeur booléenne représentant si la série est fluide ou non. Applicable uniquement aux graphiques en courbes et à nuages de points.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[chartType](/javascript/api/excel/excel.chartseriesupdatedata#charttype)|Représente le type de graphique d’une série. Pour plus d’informations, voir Excel. ChartType.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesupdatedata#doughnutholesize)|Représente la taille du centre d’une série de graphiques en anneaux.  Valide uniquement dans les graphiques en anneaux et doughnutExploded.|
||[filtré](/javascript/api/excel/excel.chartseriesupdatedata#filtered)|Valeur booléenne représentant si la série est filtrée ou non. Non applicable pour les graphiques en surface.|
||[gapWidth](/javascript/api/excel/excel.chartseriesupdatedata#gapwidth)|Représente la largeur de l’intervalle d’une série de graphique.  Valide uniquement sur les graphiques en barres et colonnes, ainsi que|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesupdatedata#hasdatalabels)|Valeur booléenne représentant si la série a des étiquettes de données ou non.|
||[MarkerBackgroundColor,](/javascript/api/excel/excel.chartseriesupdatedata#markerbackgroundcolor)|Représente la couleur d’arrière-plan de marqueurs d’une série de graphiques.|
||[MarkerForegroundColor,](/javascript/api/excel/excel.chartseriesupdatedata#markerforegroundcolor)|Représente la couleur de premier plan de marqueurs d’une série de graphiques.|
||[MarkerSize,](/javascript/api/excel/excel.chartseriesupdatedata#markersize)|Représente la taille du marqueur d’une série de graphiques.|
||[MarkerStyle,](/javascript/api/excel/excel.chartseriesupdatedata#markerstyle)|Représente le style du marqueur d’une série de graphiques. Pour plus d’informations, voir Excel. ChartMarkerStyle.|
||[PlotOrder,](/javascript/api/excel/excel.chartseriesupdatedata#plotorder)|Représente l’ordre de traçage d’une série de graphiques au sein du groupe de graphiques.|
||[showShadow](/javascript/api/excel/excel.chartseriesupdatedata#showshadow)|Valeur booléenne représentant si la série a une ombre ou non.|
||[Unie](/javascript/api/excel/excel.chartseriesupdatedata#smooth)|Valeur booléenne représentant si la série est fluide ou non. Applicable uniquement aux graphiques en courbes et à nuages de points.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (début: nombre, longueur: nombre)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Obtenir la sous-chaîne d’un titre de graphique. Le saut de ligne «\n» compte également un seul caractère.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Représente l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitle#left)|Représente la distance en points, du bord gauche du titre du graphique au bord gauche de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[position](/javascript/api/excel/excel.charttitle#position)|Représente la position du titre du graphique. Pour plus d’informations, voir Excel. ChartTitlePosition.|
||[height](/javascript/api/excel/excel.charttitle#height)|Représente la hauteur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
||[width](/javascript/api/excel/excel.charttitle#width)|Représente la largeur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
||[setFormula (Formula: String)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Définit une valeur de chaîne qui représente la formule de titre de graphique à l’aide de la notation de style A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#showshadow)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Représente l’orientation du texte du titre du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttitle#top)|Représente la distance en points, du bord supérieur du titre du graphique au bord supérieur de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Représente l’alignement vertical du titre du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartTitleData](/javascript/api/excel/excel.charttitledata)|[height](/javascript/api/excel/excel.charttitledata#height)|Représente la hauteur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitledata#horizontalalignment)|Représente l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitledata#left)|Représente la distance en points, du bord gauche du titre du graphique au bord gauche de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[position](/javascript/api/excel/excel.charttitledata#position)|Représente la position du titre du graphique. Pour plus d’informations, voir Excel. ChartTitlePosition.|
||[showShadow](/javascript/api/excel/excel.charttitledata#showshadow)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitledata#textorientation)|Représente l’orientation du texte du titre du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttitledata#top)|Représente la distance en points, du bord supérieur du titre du graphique au bord supérieur de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitledata#verticalalignment)|Représente l’alignement vertical du titre du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
||[width](/javascript/api/excel/excel.charttitledata#width)|Représente la largeur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[route](/javascript/api/excel/excel.charttitleformat#border)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de style et le poids. En lecture seule.|
|[ChartTitleFormatData](/javascript/api/excel/excel.charttitleformatdata)|[route](/javascript/api/excel/excel.charttitleformatdata#border)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de style et le poids. En lecture seule.|
|[ChartTitleFormatLoadOptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[route](/javascript/api/excel/excel.charttitleformatloadoptions#border)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de style et le poids.|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[route](/javascript/api/excel/excel.charttitleformatupdatedata#border)|Représente le format de bordure du titre du graphique, qui inclut la couleur, le style de style et le poids.|
|[ChartTitleLoadOptions](/javascript/api/excel/excel.charttitleloadoptions)|[height](/javascript/api/excel/excel.charttitleloadoptions#height)|Représente la hauteur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitleloadoptions#horizontalalignment)|Représente l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitleloadoptions#left)|Représente la distance en points, du bord gauche du titre du graphique au bord gauche de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[position](/javascript/api/excel/excel.charttitleloadoptions#position)|Représente la position du titre du graphique. Pour plus d’informations, voir Excel. ChartTitlePosition.|
||[showShadow](/javascript/api/excel/excel.charttitleloadoptions#showshadow)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitleloadoptions#textorientation)|Représente l’orientation du texte du titre du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttitleloadoptions#top)|Représente la distance en points, du bord supérieur du titre du graphique au bord supérieur de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleloadoptions#verticalalignment)|Représente l’alignement vertical du titre du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
||[width](/javascript/api/excel/excel.charttitleloadoptions#width)|Représente la largeur, exprimée en points, du titre du graphique. NULL si le titre du graphique n’est pas visible. En lecture seule.|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[horizontalAlignment](/javascript/api/excel/excel.charttitleupdatedata#horizontalalignment)|Représente l’alignement horizontal du titre du graphique.|
||[left](/javascript/api/excel/excel.charttitleupdatedata#left)|Représente la distance en points, du bord gauche du titre du graphique au bord gauche de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[position](/javascript/api/excel/excel.charttitleupdatedata#position)|Représente la position du titre du graphique. Pour plus d’informations, voir Excel. ChartTitlePosition.|
||[showShadow](/javascript/api/excel/excel.charttitleupdatedata#showshadow)|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|
||[textOrientation](/javascript/api/excel/excel.charttitleupdatedata#textorientation)|Représente l’orientation du texte du titre du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttitleupdatedata#top)|Représente la distance en points, du bord supérieur du titre du graphique au bord supérieur de la zone de graphique. NULL si le titre du graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleupdatedata#verticalalignment)|Représente l’alignement vertical du titre du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Supprime l’objet courbe de tendance.|
||[ordonn](/javascript/api/excel/excel.charttrendline#intercept)|Représente la valeur intercept de la courbe de tendance. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs automatiques). La valeur renvoyée est toujours un nombre.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Représente la période d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Représente le nom de la courbe de tendance. Peut être configurée pour une valeur de chaîne, ou peut être configurée pour que la valeur null représente les valeurs automatiques. La valeur renvoyée est toujours une chaîne.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Représente l’ordre d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type polynomial.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Représente la mise en forme de courbe de tendance de graphique.|
||[Set (propriétés: Excel. ChartTrendline)](/javascript/api/excel/excel.charttrendline#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartTrendlineUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendline#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Représente le type de courbe de tendance de graphique.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[Add (type?: "linéaire" \| "exponentiel" \| "logarithmique" \| "MovingAverage" \| " \| puissance" "Power")](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Ajoute une nouvelle courbe de tendance à la collection de courbes de tendance.|
||[Add (type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Ajoute une nouvelle courbe de tendance à la collection de courbes de tendance.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Renvoie le nombre de courbes de tendance de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Obtient un objet courbe de tendance par index, c'est-à-dire par ordre d’insertion dans le tableau des éléments.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#format)|Pour chaque élément de la collection: représente la mise en forme d’une courbe de tendance de graphique.|
||[ordonn](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#intercept)|Pour chaque élément de la collection: représente la valeur d’interception de la courbe de tendance. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs automatiques). La valeur renvoyée est toujours un nombre.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#movingaverageperiod)|Pour chaque élément de la collection: représente la période d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#name)|Pour chaque élément de la collection: représente le nom de la courbe de tendance. Peut être configurée pour une valeur de chaîne, ou peut être configurée pour que la valeur null représente les valeurs automatiques. La valeur renvoyée est toujours une chaîne.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#polynomialorder)|Pour chaque élément de la collection: représente l’ordre d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type polynomial.|
||[type](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#type)|Pour chaque élément de la collection: représente le type d’une courbe de tendance de graphique.|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[format](/javascript/api/excel/excel.charttrendlinedata#format)|Représente la mise en forme de courbe de tendance de graphique.|
||[ordonn](/javascript/api/excel/excel.charttrendlinedata#intercept)|Représente la valeur intercept de la courbe de tendance. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs automatiques). La valeur renvoyée est toujours un nombre.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinedata#movingaverageperiod)|Représente la période d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlinedata#name)|Représente le nom de la courbe de tendance. Peut être configurée pour une valeur de chaîne, ou peut être configurée pour que la valeur null représente les valeurs automatiques. La valeur renvoyée est toujours une chaîne.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinedata#polynomialorder)|Représente l’ordre d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type polynomial.|
||[type](/javascript/api/excel/excel.charttrendlinedata#type)|Représente le type de courbe de tendance de graphique.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Représente le format des lignes du graphique. En lecture seule.|
||[Set (propriétés: Excel. ChartTrendlineFormat)](/javascript/api/excel/excel.charttrendlineformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartTrendlineFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendlineformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartTrendlineFormatData](/javascript/api/excel/excel.charttrendlineformatdata)|[line](/javascript/api/excel/excel.charttrendlineformatdata#line)|Représente le format des lignes du graphique. En lecture seule.|
|[ChartTrendlineFormatLoadOptions](/javascript/api/excel/excel.charttrendlineformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charttrendlineformatloadoptions#line)|Représente le format des lignes du graphique.|
|[ChartTrendlineFormatUpdateData](/javascript/api/excel/excel.charttrendlineformatupdatedata)|[line](/javascript/api/excel/excel.charttrendlineformatupdatedata#line)|Représente le format des lignes du graphique.|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlineloadoptions#format)|Représente la mise en forme de courbe de tendance de graphique.|
||[ordonn](/javascript/api/excel/excel.charttrendlineloadoptions#intercept)|Représente la valeur intercept de la courbe de tendance. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs automatiques). La valeur renvoyée est toujours un nombre.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineloadoptions#movingaverageperiod)|Représente la période d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlineloadoptions#name)|Représente le nom de la courbe de tendance. Peut être configurée pour une valeur de chaîne, ou peut être configurée pour que la valeur null représente les valeurs automatiques. La valeur renvoyée est toujours une chaîne.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineloadoptions#polynomialorder)|Représente l’ordre d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type polynomial.|
||[type](/javascript/api/excel/excel.charttrendlineloadoptions#type)|Représente le type de courbe de tendance de graphique.|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[format](/javascript/api/excel/excel.charttrendlineupdatedata#format)|Représente la mise en forme de courbe de tendance de graphique.|
||[ordonn](/javascript/api/excel/excel.charttrendlineupdatedata#intercept)|Représente la valeur intercept de la courbe de tendance. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs automatiques). La valeur renvoyée est toujours un nombre.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineupdatedata#movingaverageperiod)|Représente la période d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlineupdatedata#name)|Représente le nom de la courbe de tendance. Peut être configurée pour une valeur de chaîne, ou peut être configurée pour que la valeur null représente les valeurs automatiques. La valeur renvoyée est toujours une chaîne.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineupdatedata#polynomialorder)|Représente l’ordre d’une courbe de tendance de graphique. Applicable uniquement à la courbe de tendance avec le type polynomial.|
||[type](/javascript/api/excel/excel.charttrendlineupdatedata#type)|Représente le type de courbe de tendance de graphique.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[chartType](/javascript/api/excel/excel.chartupdatedata#charttype)|Représente le type d’un graphique. Pour plus d’informations, voir Excel. ChartType.|
||[ShowAllFieldButtons,](/javascript/api/excel/excel.chartupdatedata#showallfieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.customproperty#key)|Obtient la clé de la propriété personnalisée. En lecture seule.|
||[type](/javascript/api/excel/excel.customproperty#type)|Obtient le type de valeur de la propriété personnalisée. En lecture seule.|
||[Set (propriétés: Excel. CustomProperty)](/javascript/api/excel/excel.customproperty#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. CustomPropertyUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.customproperty#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[value](/javascript/api/excel/excel.customproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[Add (Key: chaîne, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Crée une nouvelle propriété personnalisée ou en définit une existante.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Supprime toutes les propriétés personnalisées de cette collection.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Obtient le nombre des propriétés personnalisées.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Lève une exception si la propriété personnalisée n’existe pas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CustomPropertyCollectionLoadOptions](/javascript/api/excel/excel.custompropertycollectionloadoptions)|[$all](/javascript/api/excel/excel.custompropertycollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertycollectionloadoptions#key)|Pour chaque élément de la collection: obtient la clé de la propriété personnalisée. En lecture seule.|
||[type](/javascript/api/excel/excel.custompropertycollectionloadoptions#type)|Pour chaque élément de la collection: obtient le type de valeur de la propriété personnalisée. En lecture seule.|
||[value](/javascript/api/excel/excel.custompropertycollectionloadoptions#value)|Pour chaque élément de la collection: Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyData](/javascript/api/excel/excel.custompropertydata)|[key](/javascript/api/excel/excel.custompropertydata#key)|Obtient la clé de la propriété personnalisée. En lecture seule.|
||[type](/javascript/api/excel/excel.custompropertydata#type)|Obtient le type de valeur de la propriété personnalisée. En lecture seule.|
||[value](/javascript/api/excel/excel.custompropertydata#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyLoadOptions](/javascript/api/excel/excel.custompropertyloadoptions)|[$all](/javascript/api/excel/excel.custompropertyloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertyloadoptions#key)|Obtient la clé de la propriété personnalisée. En lecture seule.|
||[type](/javascript/api/excel/excel.custompropertyloadoptions#type)|Obtient le type de valeur de la propriété personnalisée. En lecture seule.|
||[value](/javascript/api/excel/excel.custompropertyloadoptions#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[CustomPropertyUpdateData](/javascript/api/excel/excel.custompropertyupdatedata)|[value](/javascript/api/excel/excel.custompropertyupdatedata#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[RefreshAll, ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Actualise toutes les dataConnections de la collection.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[créés](/javascript/api/excel/excel.documentproperties#author)|Obtient ou définit l’auteur du classeur.|
||[catégories](/javascript/api/excel/excel.documentproperties#category)|Obtient ou définit la catégorie du classeur.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Obtient ou définit les commentaires du classeur.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Obtient ou définit la compagnie du classeur.|
||[Mots clés](/javascript/api/excel/excel.documentproperties#keywords)|Obtient ou définit les mots clés du classeur.|
||[dirigeant](/javascript/api/excel/excel.documentproperties#manager)|Obtient ou définit le responsable du classeur.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Obtient la date de création du classeur. En lecture seule.|
||[personnalisé](/javascript/api/excel/excel.documentproperties#custom)|Obtient la collection de propriétés personnalisées du classeur. En lecture seule.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Obtient ou définit le dernier auteur du classeur. En lecture seule.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Obtient le numéro de révision du classeur. En lecture seule.|
||[Set (propriétés: Excel. DocumentProperties)](/javascript/api/excel/excel.documentproperties#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. DocumentPropertiesUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.documentproperties#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Obtient ou définit le sujet du classeur.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Obtient ou définit le titre du classeur.|
|[DocumentPropertiesData](/javascript/api/excel/excel.documentpropertiesdata)|[créés](/javascript/api/excel/excel.documentpropertiesdata#author)|Obtient ou définit l’auteur du classeur.|
||[catégories](/javascript/api/excel/excel.documentpropertiesdata#category)|Obtient ou définit la catégorie du classeur.|
||[comments](/javascript/api/excel/excel.documentpropertiesdata#comments)|Obtient ou définit les commentaires du classeur.|
||[company](/javascript/api/excel/excel.documentpropertiesdata#company)|Obtient ou définit la compagnie du classeur.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesdata#creationdate)|Obtient la date de création du classeur. En lecture seule.|
||[personnalisé](/javascript/api/excel/excel.documentpropertiesdata#custom)|Obtient la collection de propriétés personnalisées du classeur. En lecture seule.|
||[Mots clés](/javascript/api/excel/excel.documentpropertiesdata#keywords)|Obtient ou définit les mots clés du classeur.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesdata#lastauthor)|Obtient ou définit le dernier auteur du classeur. En lecture seule.|
||[dirigeant](/javascript/api/excel/excel.documentpropertiesdata#manager)|Obtient ou définit le responsable du classeur.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesdata#revisionnumber)|Obtient le numéro de révision du classeur. En lecture seule.|
||[subject](/javascript/api/excel/excel.documentpropertiesdata#subject)|Obtient ou définit le sujet du classeur.|
||[title](/javascript/api/excel/excel.documentpropertiesdata#title)|Obtient ou définit le titre du classeur.|
|[DocumentPropertiesLoadOptions](/javascript/api/excel/excel.documentpropertiesloadoptions)|[$all](/javascript/api/excel/excel.documentpropertiesloadoptions#$all)||
||[créés](/javascript/api/excel/excel.documentpropertiesloadoptions#author)|Obtient ou définit l’auteur du classeur.|
||[catégories](/javascript/api/excel/excel.documentpropertiesloadoptions#category)|Obtient ou définit la catégorie du classeur.|
||[comments](/javascript/api/excel/excel.documentpropertiesloadoptions#comments)|Obtient ou définit les commentaires du classeur.|
||[company](/javascript/api/excel/excel.documentpropertiesloadoptions#company)|Obtient ou définit la compagnie du classeur.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesloadoptions#creationdate)|Obtient la date de création du classeur. En lecture seule.|
||[Mots clés](/javascript/api/excel/excel.documentpropertiesloadoptions#keywords)|Obtient ou définit les mots clés du classeur.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesloadoptions#lastauthor)|Obtient ou définit le dernier auteur du classeur. En lecture seule.|
||[dirigeant](/javascript/api/excel/excel.documentpropertiesloadoptions#manager)|Obtient ou définit le responsable du classeur.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesloadoptions#revisionnumber)|Obtient le numéro de révision du classeur. En lecture seule.|
||[subject](/javascript/api/excel/excel.documentpropertiesloadoptions#subject)|Obtient ou définit le sujet du classeur.|
||[title](/javascript/api/excel/excel.documentpropertiesloadoptions#title)|Obtient ou définit le titre du classeur.|
|[DocumentPropertiesUpdateData](/javascript/api/excel/excel.documentpropertiesupdatedata)|[créés](/javascript/api/excel/excel.documentpropertiesupdatedata#author)|Obtient ou définit l’auteur du classeur.|
||[catégories](/javascript/api/excel/excel.documentpropertiesupdatedata#category)|Obtient ou définit la catégorie du classeur.|
||[comments](/javascript/api/excel/excel.documentpropertiesupdatedata#comments)|Obtient ou définit les commentaires du classeur.|
||[company](/javascript/api/excel/excel.documentpropertiesupdatedata#company)|Obtient ou définit la compagnie du classeur.|
||[Mots clés](/javascript/api/excel/excel.documentpropertiesupdatedata#keywords)|Obtient ou définit les mots clés du classeur.|
||[dirigeant](/javascript/api/excel/excel.documentpropertiesupdatedata#manager)|Obtient ou définit le responsable du classeur.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesupdatedata#revisionnumber)|Obtient le numéro de révision du classeur. En lecture seule.|
||[subject](/javascript/api/excel/excel.documentpropertiesupdatedata#subject)|Obtient ou définit le sujet du classeur.|
||[title](/javascript/api/excel/excel.documentpropertiesupdatedata#title)|Obtient ou définit le titre du classeur.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Obtient ou définit la formule de l’élément nommé.  La formule commence toujours par un signe '='.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Renvoie un objet contenant les valeurs et les types de l’élément nommé. En lecture seule.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Représente les types de chaque élément dans le tableau d’éléments nommés|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Représente les valeurs de chaque élément dans le tableau élément nommé.|
|[NamedItemArrayValuesData](/javascript/api/excel/excel.nameditemarrayvaluesdata)|[types](/javascript/api/excel/excel.nameditemarrayvaluesdata#types)|Représente les types de chaque élément dans le tableau d’éléments nommés|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesdata#values)|Représente les valeurs de chaque élément dans le tableau élément nommé.|
|[NamedItemArrayValuesLoadOptions](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions)|[$all](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#$all)||
||[types](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#types)|Représente les types de chaque élément dans le tableau d’éléments nommés|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#values)|Représente les valeurs de chaque élément dans le tableau élément nommé.|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemcollectionloadoptions#arrayvalues)|Pour chaque élément de la collection: renvoie un objet contenant les valeurs et les types de l’élément nommé.|
||[formula](/javascript/api/excel/excel.nameditemcollectionloadoptions#formula)|Pour chaque élément de la collection: Obtient ou définit la formule de l’élément nommé.  La formule commence toujours par un signe '='.|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[arrayValues](/javascript/api/excel/excel.nameditemdata#arrayvalues)|Renvoie un objet contenant les valeurs et les types de l’élément nommé. En lecture seule.|
||[formula](/javascript/api/excel/excel.nameditemdata#formula)|Obtient ou définit la formule de l’élément nommé.  La formule commence toujours par un signe '='.|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemloadoptions#arrayvalues)|Renvoie un objet contenant les valeurs et les types de l’élément nommé.|
||[formula](/javascript/api/excel/excel.nameditemloadoptions#formula)|Obtient ou définit la formule de l’élément nommé.  La formule commence toujours par un signe '='.|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[formula](/javascript/api/excel/excel.nameditemupdatedata#formula)|Obtient ou définit la formule de l’élément nommé.  La formule commence toujours par un signe '='.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: nombre, numColumns: nombre)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Obtient un objet Plage avec la même cellule supérieure gauche que l’objet de Plage en cours, mais avec un nombre spécifié de lignes et colonnes.|
||[getImage ()](/javascript/api/excel/excel.range#getimage--)|Affiche la plage en tant qu’image png encodée au format Base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Renvoie un objet PLage qui représente la région environnante pour la cellule en haut à gauche de cette plage. Une région environnante est une plage délimitée par une combinaison de lignes et de colonnes vides par rapport à cette plage.|
||[lien hypertexte](/javascript/api/excel/excel.range#hyperlink)|Représente le lien hypertexte de la plage active.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Représente le code de format numérique d’Excel pour la plage donnée en tant que chaîne dans la langue de l’utilisateur.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Représente si la plage active est une colonne entière. En lecture seule.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Représente si la plage active est une ligne entière. En lecture seule.|
||[showCard ()](/javascript/api/excel/excel.range#showcard--)|Affiche la carte pour une cellule active si son contenu est riche en valeur.|
||[style](/javascript/api/excel/excel.range#style)|Représente le style de la plage actuelle.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[lien hypertexte](/javascript/api/excel/excel.rangedata#hyperlink)|Représente le lien hypertexte de la plage active.|
||[isEntireColumn](/javascript/api/excel/excel.rangedata#isentirecolumn)|Représente si la plage active est une colonne entière. En lecture seule.|
||[isEntireRow](/javascript/api/excel/excel.rangedata#isentirerow)|Représente si la plage active est une ligne entière. En lecture seule.|
||[numberFormatLocal](/javascript/api/excel/excel.rangedata#numberformatlocal)|Représente le code de format numérique d’Excel pour la plage donnée en tant que chaîne dans la langue de l’utilisateur.|
||[style](/javascript/api/excel/excel.rangedata#style)|Représente le style de la plage actuelle.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Obtient ou définit l’orientation du texte de toutes les cellules dans la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Indique si la largeur de la colonne de l’objet Range est égale à la largeur standard de la feuille.|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[textOrientation](/javascript/api/excel/excel.rangeformatdata#textorientation)|Obtient ou définit l’orientation du texte de toutes les cellules dans la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatdata#usestandardheight)|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatdata#usestandardwidth)|Indique si la largeur de la colonne de l’objet Range est égale à la largeur standard de la feuille.|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[textOrientation](/javascript/api/excel/excel.rangeformatloadoptions#textorientation)|Obtient ou définit l’orientation du texte de toutes les cellules dans la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatloadoptions#usestandardheight)|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatloadoptions#usestandardwidth)|Indique si la largeur de la colonne de l’objet Range est égale à la largeur standard de la feuille.|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[textOrientation](/javascript/api/excel/excel.rangeformatupdatedata#textorientation)|Obtient ou définit l’orientation du texte de toutes les cellules dans la plage.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatupdatedata#usestandardheight)|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatupdatedata#usestandardwidth)|Indique si la largeur de la colonne de l’objet Range est égale à la largeur standard de la feuille.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[adresse](/javascript/api/excel/excel.rangehyperlink#address)|Représente l’url cible du lien hypertexte.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Représente la cible de référence de document pour le lien hypertexte.|
||[Info](/javascript/api/excel/excel.rangehyperlink#screentip)|Représente la chaîne affichée lorsque vous pointez sur le lien hypertexte.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Représente la chaîne qui s’affiche dans la cellule en haut à gauche de la plage.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[lien hypertexte](/javascript/api/excel/excel.rangeloadoptions#hyperlink)|Représente le lien hypertexte de la plage active.|
||[isEntireColumn](/javascript/api/excel/excel.rangeloadoptions#isentirecolumn)|Représente si la plage active est une colonne entière. En lecture seule.|
||[isEntireRow](/javascript/api/excel/excel.rangeloadoptions#isentirerow)|Représente si la plage active est une ligne entière. En lecture seule.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeloadoptions#numberformatlocal)|Représente le code de format numérique d’Excel pour la plage donnée en tant que chaîne dans la langue de l’utilisateur.|
||[style](/javascript/api/excel/excel.rangeloadoptions#style)|Représente le style de la plage actuelle.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[lien hypertexte](/javascript/api/excel/excel.rangeupdatedata#hyperlink)|Représente le lien hypertexte de la plage active.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeupdatedata#numberformatlocal)|Représente le code de format numérique d’Excel pour la plage donnée en tant que chaîne dans la langue de l’utilisateur.|
||[style](/javascript/api/excel/excel.rangeupdatedata#style)|Représente le style de la plage actuelle.|
|[Style](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Supprime ce style.|
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
||[Set (propriétés: Excel. style)](/javascript/api/excel/excel.style#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. StyleUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.style#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Indique si le texte s’ajuste automatiquement pour tenir dans la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Représente l’alignement vertical du style. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Indique si Microsoft Excel renvoie le texte à la ligne dans l’objet.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Ajoute un nouveau style à la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Obtient un style par nom.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[$all](/javascript/api/excel/excel.stylecollectionloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.stylecollectionloadoptions#borders)|Pour chaque élément de la collection: une collection de bordures de quatre objets Border qui représentent le style des quatre bordures.|
||[builtIn](/javascript/api/excel/excel.stylecollectionloadoptions#builtin)|Pour chaque élément de la collection: indique si le style est un style prédéfini.|
||[fill](/javascript/api/excel/excel.stylecollectionloadoptions#fill)|Pour chaque élément de la collection: le remplissage du style.|
||[police](/javascript/api/excel/excel.stylecollectionloadoptions#font)|Pour chaque élément de la collection: objet font qui représente la police du style.|
||[formulaHidden](/javascript/api/excel/excel.stylecollectionloadoptions#formulahidden)|Pour chaque élément de la collection: indique si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#horizontalalignment)|Pour chaque élément de la collection: représente l’alignement horizontal du style. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[IncludeAlignment,](/javascript/api/excel/excel.stylecollectionloadoptions#includealignment)|Pour chaque élément de la collection: indique si le style inclut les propriétés autoindent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel et TextOrientation.|
||[IncludeBorder,](/javascript/api/excel/excel.stylecollectionloadoptions#includeborder)|Pour chaque élément de la collection: indique si le style inclut les propriétés Color, ColorIndex, LineStyle et Weight Border.|
||[IncludeFont,](/javascript/api/excel/excel.stylecollectionloadoptions#includefont)|Pour chaque élément de la collection: indique si le style inclut les propriétés de police background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript et Underline.|
||[IncludeNumber,](/javascript/api/excel/excel.stylecollectionloadoptions#includenumber)|Pour chaque élément de la collection: indique si le style inclut la propriété NumberFormat.|
||[IncludePatterns,](/javascript/api/excel/excel.stylecollectionloadoptions#includepatterns)|Pour chaque élément de la collection: indique si le style inclut les propriétés Color, ColorIndex, InvertIfNegative, pattern, PatternColor et PatternColorIndex, Interior.|
||[IncludeProtection,](/javascript/api/excel/excel.stylecollectionloadoptions#includeprotection)|Pour chaque élément de la collection: indique si le style inclut les propriétés FormulaHidden et protection verrouillée.|
||[indentLevel](/javascript/api/excel/excel.stylecollectionloadoptions#indentlevel)|Pour chaque élément de la collection: entier compris entre 0 et 250 qui indique le niveau de retrait pour le style.|
||[locked](/javascript/api/excel/excel.stylecollectionloadoptions#locked)|Pour chaque élément de la collection: indique si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[name](/javascript/api/excel/excel.stylecollectionloadoptions#name)|Pour chaque élément de la collection: le nom du style.|
||[numberFormat](/javascript/api/excel/excel.stylecollectionloadoptions#numberformat)|Pour chaque élément de la collection: code de format du format numérique pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.stylecollectionloadoptions#numberformatlocal)|Pour chaque élément de la collection: le code de format localisé du format numérique pour le style.|
||[readingOrder](/javascript/api/excel/excel.stylecollectionloadoptions#readingorder)|Pour chaque élément de la collection: le sens de lecture du style.|
||[shrinkToFit](/javascript/api/excel/excel.stylecollectionloadoptions#shrinktofit)|Pour chaque élément de la collection: indique si le texte s’ajuste automatiquement à la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#verticalalignment)|Pour chaque élément de la collection: représente l’alignement vertical du style. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.stylecollectionloadoptions#wraptext)|Pour chaque élément de la collection: indique si Microsoft Excel renvoie le texte à la ligne dans l’objet.|
|[StyleData](/javascript/api/excel/excel.styledata)|[Borders](/javascript/api/excel/excel.styledata#borders)|Collection de bordures de quatre objets qui représente le style des quatre bordures.|
||[builtIn](/javascript/api/excel/excel.styledata#builtin)|Indique si le style est un style intégré.|
||[fill](/javascript/api/excel/excel.styledata#fill)|Remplissage du style.|
||[police](/javascript/api/excel/excel.styledata#font)|Renvoie un objet Police qui représente la police du style.|
||[formulaHidden](/javascript/api/excel/excel.styledata#formulahidden)|Indique si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.styledata#horizontalalignment)|Représente l’alignement horizontal pour le style. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[IncludeAlignment,](/javascript/api/excel/excel.styledata#includealignment)|Indique si le style inclut les propriétés AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, et TextOrientation.|
||[IncludeBorder,](/javascript/api/excel/excel.styledata#includeborder)|Indique si le style inclut les propriétés dColor, ColorIndex, LineStyle, et Weight border.|
||[IncludeFont,](/javascript/api/excel/excel.styledata#includefont)|Indique si le style inclut les propriétés dBackground, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, et Underline font.|
||[IncludeNumber,](/javascript/api/excel/excel.styledata#includenumber)|Indique si le style inclut la propriété NumberFormat.|
||[IncludePatterns,](/javascript/api/excel/excel.styledata#includepatterns)|Indique si le style inclut les propriétés Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, et PatternColorIndex interior.|
||[IncludeProtection,](/javascript/api/excel/excel.styledata#includeprotection)|Indique si le style inclut les propriétés FormulaHidden et Locked protection.|
||[indentLevel](/javascript/api/excel/excel.styledata#indentlevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[locked](/javascript/api/excel/excel.styledata#locked)|Indique si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[name](/javascript/api/excel/excel.styledata#name)|Nom du style.|
||[numberFormat](/javascript/api/excel/excel.styledata#numberformat)|Le code de format du nombre format pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.styledata#numberformatlocal)|Le code de format localisé du nombre format pour le style.|
||[readingOrder](/javascript/api/excel/excel.styledata#readingorder)|L’ordre de lecture du style.|
||[shrinkToFit](/javascript/api/excel/excel.styledata#shrinktofit)|Indique si le texte s’ajuste automatiquement pour tenir dans la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.styledata#verticalalignment)|Représente l’alignement vertical du style. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.styledata#wraptext)|Indique si Microsoft Excel renvoie le texte à la ligne dans l’objet.|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[$all](/javascript/api/excel/excel.styleloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.styleloadoptions#borders)|Collection de bordures de quatre objets qui représente le style des quatre bordures.|
||[builtIn](/javascript/api/excel/excel.styleloadoptions#builtin)|Indique si le style est un style intégré.|
||[fill](/javascript/api/excel/excel.styleloadoptions#fill)|Remplissage du style.|
||[police](/javascript/api/excel/excel.styleloadoptions#font)|Renvoie un objet Police qui représente la police du style.|
||[formulaHidden](/javascript/api/excel/excel.styleloadoptions#formulahidden)|Indique si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.styleloadoptions#horizontalalignment)|Représente l’alignement horizontal pour le style. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[IncludeAlignment,](/javascript/api/excel/excel.styleloadoptions#includealignment)|Indique si le style inclut les propriétés AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, et TextOrientation.|
||[IncludeBorder,](/javascript/api/excel/excel.styleloadoptions#includeborder)|Indique si le style inclut les propriétés dColor, ColorIndex, LineStyle, et Weight border.|
||[IncludeFont,](/javascript/api/excel/excel.styleloadoptions#includefont)|Indique si le style inclut les propriétés dBackground, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, et Underline font.|
||[IncludeNumber,](/javascript/api/excel/excel.styleloadoptions#includenumber)|Indique si le style inclut la propriété NumberFormat.|
||[IncludePatterns,](/javascript/api/excel/excel.styleloadoptions#includepatterns)|Indique si le style inclut les propriétés Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, et PatternColorIndex interior.|
||[IncludeProtection,](/javascript/api/excel/excel.styleloadoptions#includeprotection)|Indique si le style inclut les propriétés FormulaHidden et Locked protection.|
||[indentLevel](/javascript/api/excel/excel.styleloadoptions#indentlevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[locked](/javascript/api/excel/excel.styleloadoptions#locked)|Indique si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[name](/javascript/api/excel/excel.styleloadoptions#name)|Nom du style.|
||[numberFormat](/javascript/api/excel/excel.styleloadoptions#numberformat)|Le code de format du nombre format pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.styleloadoptions#numberformatlocal)|Le code de format localisé du nombre format pour le style.|
||[readingOrder](/javascript/api/excel/excel.styleloadoptions#readingorder)|L’ordre de lecture du style.|
||[shrinkToFit](/javascript/api/excel/excel.styleloadoptions#shrinktofit)|Indique si le texte s’ajuste automatiquement pour tenir dans la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.styleloadoptions#verticalalignment)|Représente l’alignement vertical du style. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.styleloadoptions#wraptext)|Indique si Microsoft Excel renvoie le texte à la ligne dans l’objet.|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[Borders](/javascript/api/excel/excel.styleupdatedata#borders)|Collection de bordures de quatre objets qui représente le style des quatre bordures.|
||[fill](/javascript/api/excel/excel.styleupdatedata#fill)|Remplissage du style.|
||[police](/javascript/api/excel/excel.styleupdatedata#font)|Renvoie un objet Police qui représente la police du style.|
||[formulaHidden](/javascript/api/excel/excel.styleupdatedata#formulahidden)|Indique si la formule est masquée lorsque la feuille de calcul est protégée.|
||[horizontalAlignment](/javascript/api/excel/excel.styleupdatedata#horizontalalignment)|Représente l’alignement horizontal pour le style. Pour plus d’informations, voir Excel. HorizontalAlignment.|
||[IncludeAlignment,](/javascript/api/excel/excel.styleupdatedata#includealignment)|Indique si le style inclut les propriétés AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, et TextOrientation.|
||[IncludeBorder,](/javascript/api/excel/excel.styleupdatedata#includeborder)|Indique si le style inclut les propriétés dColor, ColorIndex, LineStyle, et Weight border.|
||[IncludeFont,](/javascript/api/excel/excel.styleupdatedata#includefont)|Indique si le style inclut les propriétés dBackground, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, et Underline font.|
||[IncludeNumber,](/javascript/api/excel/excel.styleupdatedata#includenumber)|Indique si le style inclut la propriété NumberFormat.|
||[IncludePatterns,](/javascript/api/excel/excel.styleupdatedata#includepatterns)|Indique si le style inclut les propriétés Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, et PatternColorIndex interior.|
||[IncludeProtection,](/javascript/api/excel/excel.styleupdatedata#includeprotection)|Indique si le style inclut les propriétés FormulaHidden et Locked protection.|
||[indentLevel](/javascript/api/excel/excel.styleupdatedata#indentlevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[locked](/javascript/api/excel/excel.styleupdatedata#locked)|Indique si l’objet est verrouillé lorsque la feuille de calcul est protégée.|
||[numberFormat](/javascript/api/excel/excel.styleupdatedata#numberformat)|Le code de format du nombre format pour le style.|
||[numberFormatLocal](/javascript/api/excel/excel.styleupdatedata#numberformatlocal)|Le code de format localisé du nombre format pour le style.|
||[readingOrder](/javascript/api/excel/excel.styleupdatedata#readingorder)|L’ordre de lecture du style.|
||[shrinkToFit](/javascript/api/excel/excel.styleupdatedata#shrinktofit)|Indique si le texte s’ajuste automatiquement pour tenir dans la largeur de colonne disponible.|
||[verticalAlignment](/javascript/api/excel/excel.styleupdatedata#verticalalignment)|Représente l’alignement vertical du style. Pour plus d’informations, voir Excel. VerticalAlignment.|
||[wrapText](/javascript/api/excel/excel.styleupdatedata#wraptext)|Indique si Microsoft Excel renvoie le texte à la ligne dans l’objet.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Survient lorsque les données des cellules changent sur une table spécifique.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Se produit lorsque la sélection change sur une table spécifique.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[adresse](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtient l’adresse qui représente la zone modifiée d’un tableau figurant dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié. Pour plus d’informations, voir Excel. DataChangeType.|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Obtient la cellule active du classeur.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Représente toutes les connexions de données du classeur. En lecture seule.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtient le nom du classeur. En lecture seule.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtient les propriétés du classeur. En lecture seule.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Renvoie un objet de protection de classeur pour un classeur. En lecture seule.|
||[proposés](/javascript/api/excel/excel.workbook#styles)|Représente une collection de styles associés au classeur. En lecture seule.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[name](/javascript/api/excel/excel.workbookdata#name)|Obtient le nom du classeur. En lecture seule.|
||[properties](/javascript/api/excel/excel.workbookdata#properties)|Obtient les propriétés du classeur. En lecture seule.|
||[protection](/javascript/api/excel/excel.workbookdata#protection)|Renvoie un objet de protection de classeur pour un classeur. En lecture seule.|
||[proposés](/javascript/api/excel/excel.workbookdata#styles)|Représente une collection de styles associés au classeur. En lecture seule.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[name](/javascript/api/excel/excel.workbookloadoptions#name)|Obtient le nom du classeur. En lecture seule.|
||[properties](/javascript/api/excel/excel.workbookloadoptions#properties)|Obtient les propriétés du classeur.|
||[protection](/javascript/api/excel/excel.workbookloadoptions#protection)|Renvoie un objet de protection de classeur pour un classeur.|
|[Objetworkbookprotection](/javascript/api/excel/excel.workbookprotection)|[Protect (Password?: String)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protège un classeur. Échoue si le classeur est protégé.|
||[sécurisé](/javascript/api/excel/excel.workbookprotection#protected)|Indique si le classeur est protégé. En lecture seule.|
||[Unprotect (Password?: String)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Annule la protection un classeur.|
|[WorkbookProtectionData](/javascript/api/excel/excel.workbookprotectiondata)|[sécurisé](/javascript/api/excel/excel.workbookprotectiondata#protected)|Indique si le classeur est protégé. En lecture seule.|
|[WorkbookProtectionLoadOptions](/javascript/api/excel/excel.workbookprotectionloadoptions)|[$all](/javascript/api/excel/excel.workbookprotectionloadoptions#$all)||
||[sécurisé](/javascript/api/excel/excel.workbookprotectionloadoptions#protected)|Indique si le classeur est protégé. En lecture seule.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[properties](/javascript/api/excel/excel.workbookupdatedata#properties)|Obtient les propriétés du classeur.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (positionType?: "none" \| "Before \| " "après \| " "Beginning \| " "End", relativeTo?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copie une feuille de calcul et la place à la position spécifiée. Renvoie la feuille de calcul copiée.|
||[Copy (positionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copie une feuille de calcul et la place à la position spécifiée. Renvoie la feuille de calcul copiée.|
||[getRangeByIndexes (startRow: nombre, ColonneDébut: nombre, rowCount: nombre, columnCount: nombre)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Obtient l’objet plage commençant à un index de ligne et de colonne particulier et couvrant un certain nombre de lignes et de colonnes.|
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
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est activée.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Survient lors de l’ajout d’une nouvelle feuille de calcul au classeur.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est désactivée.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Survient lors de la suppression d’une feuille de calcul du classeur.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[StandardHeight,](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardheight)|Pour chaque élément de la collection: renvoie la hauteur standard (par défaut) de toutes les lignes de la feuille de calcul, exprimée en points. En lecture seule.|
||[StandardWidth,](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardwidth)|Pour chaque élément de la collection: cette propriété renvoie ou définit la largeur standard (par défaut) de toutes les colonnes de la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheetcollectionloadoptions#tabcolor)|Pour chaque élément de la collection: Obtient ou définit la couleur de l’onglet de la feuille de calcul.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[StandardHeight,](/javascript/api/excel/excel.worksheetdata#standardheight)|Renvoie la hauteur standard (par défaut) de toutes les lignes dans la feuille de calcul, en points. En lecture seule.|
||[StandardWidth,](/javascript/api/excel/excel.worksheetdata#standardwidth)|Renvoie ou définit la largeur standard (par défaut) de toutes les colonnes dans la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheetdata#tabcolor)|Obtient ou modifie la couleur d’onglet de la feuille de calcul.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[type](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est desactivée.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est supprimée du classeur.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: chaîne \| de plage)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[freezeColumns (Count?: nombre)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Figer la/les première(s) colonne(s) de la feuille de calcul en place.|
||[freezeRows (Count?: nombre)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Figer la/les première(s) ligne(s) de la feuille de calcul en place.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|
||[Unfreeze ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Supprime tous les volets figés dans la feuille de calcul.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[StandardHeight,](/javascript/api/excel/excel.worksheetloadoptions#standardheight)|Renvoie la hauteur standard (par défaut) de toutes les lignes dans la feuille de calcul, en points. En lecture seule.|
||[StandardWidth,](/javascript/api/excel/excel.worksheetloadoptions#standardwidth)|Renvoie ou définit la largeur standard (par défaut) de toutes les colonnes dans la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheetloadoptions#tabcolor)|Obtient ou modifie la couleur d’onglet de la feuille de calcul.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[Unprotect (Password?: String)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Annule la protection d’une feuille de calcul.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Représente l’option de protection de feuille de calcul qui autorise la modification d’objets.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Représente l’option de protection de feuille de calcul qui autorise la modification de scénarios.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Représente l’option de protection de feuille de calcul qui autorise le mode sélection.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone sélectionnée dans une feuille de calcul spécifique.|
||[type](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[StandardWidth,](/javascript/api/excel/excel.worksheetupdatedata#standardwidth)|Renvoie ou définit la largeur standard (par défaut) de toutes les colonnes dans la feuille de calcul.|
||[tabColor](/javascript/api/excel/excel.worksheetupdatedata#tabcolor)|Obtient ou modifie la couleur d’onglet de la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
