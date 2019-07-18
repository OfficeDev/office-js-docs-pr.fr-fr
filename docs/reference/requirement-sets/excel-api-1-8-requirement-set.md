---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,8
description: Détails sur l’ensemble de conditions requises ExcelApi 1,8
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a5adcf56654070ca2a8336385f73062c34e90e1d
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772008"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Nouveautés de l’API JavaScript 1.8 pour Excel

L’ensemble de conditions requises Excel JavaScript API 1.8 incluent des API pour les tableaux croisés dynamiques, validation des données, graphiques, les événements pour les diagrammes, les options de performances et création de classeur.

## <a name="pivottable"></a>Tableau croisé dynamique

Vague 2 des APIs de tableau croisé dynamique permet aux compléments de définir les hiérarchies d’un tableau croisé dynamique. Vous pouvez désormais contrôler les données et comment elles sont regroupées. Notre [Article tableau croisé dynamique](/office/dev/add-ins/excel/excel-add-ins-pivottables) a plus d’informations sur les nouvelles fonctionnalités de tableau croisé dynamique.

## <a name="data-validation"></a>Validation des données

La validation des données vous donne le contrôle sur ce qu’un utilisateur insère dans une feuille de calcul. Vous pouvez limiter les cellules à des ensembles de réponse prédéfinie ou donner des avertissements contextuels concernant des entrées indésirables. En savoir plus maintenant sur [Ajout de validation des données à des plages](/office/dev/add-ins/excel/excel-add-ins-data-validation).

## <a name="charts"></a>Graphiques

Une autre série de graphiques API apporte un meilleur contrôle par programme des éléments de graphique. Vous avez à présent un meilleur accès à la légende, axes, courbe de tendance et zone de traçage.

## <a name="events"></a>Événements

Plus d’[événements](/office/dev/add-ins/excel/excel-add-ins-events) ont été ajoutés pour les graphiques. Votre complément réagit aux interactions des utilisateurs avec le graphique. Vous pouvez également [Activer ou désactiver les événements](/office/dev/add-ins/excel/performance#enable-and-disable-events) sur l’ensemble du classeur.

## <a name="api-list"></a>Liste des API

| Class | Champs | Description |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Spécifie l’opérande de droite lorsque la propriété Operator est définie sur un opérateur binaire tel que GreaterThan (l’opérande de gauche est la valeur que l’utilisateur tente d’entrer dans la cellule). Avec les opérateurs ternaires entre et NotBetween, spécifie l’opérande inférieur.|
||[Formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Avec les opérateurs ternaires entre et NotBetween, spécifie l’opérande supérieur. N’est pas utilisé avec les opérateurs binaires, tels que GreaterThan.|
||[is](/javascript/api/excel/excel.basicdatavalidation#operator)|L’opérateur à utiliser pour la validation des données.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Cette propriété renvoie ou définit une constante d’énumération ChartCategoryLabelLevel faisant référence à|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Renvoie ou définit la façon dont les cellules vides sont tracées sur un graphique. Lecture/écriture.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Renvoie ou spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Lecture/écriture.|
||[PlotVisibleOnly,](/javascript/api/excel/excel.chart#plotvisibleonly)|Vrai si seules les cellules visibles sont tracées.Faux si les deux cellules visibles et masquées sont tracées. Lecture/écriture.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Se produit lorsque le graphique est activé.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Se produit lorsque le graphique est désactivé.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Représente la zone de traçage pour le graphique.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Cette propriété renvoie ou définit une constante d’énumération ChartSeriesNameLevel faisant référence à|
||[ShowDataLabelsOverMaximum,](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Si vous voulez afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale sur l’axe de valeur.|
||[style](/javascript/api/excel/excel.chart#style)|Cette propriété renvoie ou définit le style de graphique pour le graphique. Lecture/écriture.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Obtient l’id du graphique qui est activé.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est activé.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Obtient l’id du graphique qui est ajouté à la feuille de calcul.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est ajouté.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[aligne](/javascript/api/excel/excel.chartaxis#alignment)|Représente l’alignement vertical de l’étiquette de la graduation de l’axe spécifié. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Représente si l’axe de valeur croise l’axe de catégorie entre catégories.|
||[Niveaux](/javascript/api/excel/excel.chartaxis#multilevel)|Représente si un axe est à plusieurs niveaux ou non.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Représente le code de format pour l’étiquette de graduation d’axe.|
||[compensé](/javascript/api/excel/excel.chartaxis#offset)|Représente la distance entre les niveaux d’étiquettes et la distance entre le premier niveau et la ligne d’axe. La valeur doit être un entier compris entre 0 et 1000.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Représente la position de l’axe spécifié où l’autre axe le croise. Pour plus d’informations, voir Excel. ChartAxisPosition.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Représente la position de l’axe spécifié où l’autre axe le croise. Vous devez utiliser la méthode SetPositionAt(double) pour définir cette propriété.|
||[setPositionAt (valeur: nombre)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Représente la position de l’axe spécifié où l’autre axe le croise.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Représente l’orientation du texte de l’étiquette de graduation de l’axe. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[aligne](/javascript/api/excel/excel.chartaxisdata#alignment)|Représente l’alignement vertical de l’étiquette de la graduation de l’axe spécifié. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisdata#isbetweencategories)|Représente si l’axe de valeur croise l’axe de catégorie entre catégories.|
||[Niveaux](/javascript/api/excel/excel.chartaxisdata#multilevel)|Représente si un axe est à plusieurs niveaux ou non.|
||[numberFormat](/javascript/api/excel/excel.chartaxisdata#numberformat)|Représente le code de format pour l’étiquette de graduation d’axe.|
||[compensé](/javascript/api/excel/excel.chartaxisdata#offset)|Représente la distance entre les niveaux d’étiquettes et la distance entre le premier niveau et la ligne d’axe. La valeur doit être un entier compris entre 0 et 1000.|
||[position](/javascript/api/excel/excel.chartaxisdata#position)|Représente la position de l’axe spécifié où l’autre axe le croise. Pour plus d’informations, voir Excel. ChartAxisPosition.|
||[positionAt](/javascript/api/excel/excel.chartaxisdata#positionat)|Représente la position de l’axe spécifié où l’autre axe le croise. Vous devez utiliser la méthode SetPositionAt(double) pour définir cette propriété.|
||[textOrientation](/javascript/api/excel/excel.chartaxisdata#textorientation)|Représente l’orientation du texte de l’étiquette de graduation de l’axe. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Représente la mise en forme de remplissage du graphique. En lecture seule.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[aligne](/javascript/api/excel/excel.chartaxisloadoptions#alignment)|Représente l’alignement vertical de l’étiquette de la graduation de l’axe spécifié. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisloadoptions#isbetweencategories)|Représente si l’axe de valeur croise l’axe de catégorie entre catégories.|
||[Niveaux](/javascript/api/excel/excel.chartaxisloadoptions#multilevel)|Représente si un axe est à plusieurs niveaux ou non.|
||[numberFormat](/javascript/api/excel/excel.chartaxisloadoptions#numberformat)|Représente le code de format pour l’étiquette de graduation d’axe.|
||[compensé](/javascript/api/excel/excel.chartaxisloadoptions#offset)|Représente la distance entre les niveaux d’étiquettes et la distance entre le premier niveau et la ligne d’axe. La valeur doit être un entier compris entre 0 et 1000.|
||[position](/javascript/api/excel/excel.chartaxisloadoptions#position)|Représente la position de l’axe spécifié où l’autre axe le croise. Pour plus d’informations, voir Excel. ChartAxisPosition.|
||[positionAt](/javascript/api/excel/excel.chartaxisloadoptions#positionat)|Représente la position de l’axe spécifié où l’autre axe le croise. Vous devez utiliser la méthode SetPositionAt(double) pour définir cette propriété.|
||[textOrientation](/javascript/api/excel/excel.chartaxisloadoptions#textorientation)|Représente l’orientation du texte de l’étiquette de graduation de l’axe. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (Formula: String)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Valeur de chaîne qui représente la formule de titre de l’axe graphique à l’aide de la notation de style A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[route](/javascript/api/excel/excel.chartaxistitleformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Représente la mise en forme de remplissage du graphique.|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[route](/javascript/api/excel/excel.chartaxistitleformatdata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[route](/javascript/api/excel/excel.chartaxistitleformatloadoptions#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[route](/javascript/api/excel/excel.chartaxistitleformatupdatedata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[aligne](/javascript/api/excel/excel.chartaxisupdatedata#alignment)|Représente l’alignement vertical de l’étiquette de la graduation de l’axe spécifié. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisupdatedata#isbetweencategories)|Représente si l’axe de valeur croise l’axe de catégorie entre catégories.|
||[Niveaux](/javascript/api/excel/excel.chartaxisupdatedata#multilevel)|Représente si un axe est à plusieurs niveaux ou non.|
||[numberFormat](/javascript/api/excel/excel.chartaxisupdatedata#numberformat)|Représente le code de format pour l’étiquette de graduation d’axe.|
||[compensé](/javascript/api/excel/excel.chartaxisupdatedata#offset)|Représente la distance entre les niveaux d’étiquettes et la distance entre le premier niveau et la ligne d’axe. La valeur doit être un entier compris entre 0 et 1000.|
||[position](/javascript/api/excel/excel.chartaxisupdatedata#position)|Représente la position de l’axe spécifié où l’autre axe le croise. Pour plus d’informations, voir Excel. ChartAxisPosition.|
||[textOrientation](/javascript/api/excel/excel.chartaxisupdatedata#textorientation)|Représente l’orientation du texte de l’étiquette de graduation de l’axe. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Désactiver le format de bordure d’un élément de graphique.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Se produit lorsqu’un graphique est activé.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Survient lors de l’ajout d’un nouveau graphique à la feuille de calcul.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Se produit lorsqu’un graphique est désactivé.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Survient lors de la suppression d’un graphique.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartcollectionloadoptions#categorylabellevel)|Pour chaque élément de la collection: renvoie ou définit une constante d’énumération ChartCategoryLabelLevel faisant référence à|
||[displayBlanksAs](/javascript/api/excel/excel.chartcollectionloadoptions#displayblanksas)|Pour chaque élément de la collection: cette propriété renvoie ou définit la façon dont les cellules vides sont tracées dans un graphique. Lecture/écriture.|
||[plotArea](/javascript/api/excel/excel.chartcollectionloadoptions#plotarea)|Pour chaque élément de la collection: représente le plotArea du graphique.|
||[plotBy](/javascript/api/excel/excel.chartcollectionloadoptions#plotby)|Pour chaque élément de la collection: cette propriété renvoie ou définit la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Lecture/écriture.|
||[PlotVisibleOnly,](/javascript/api/excel/excel.chartcollectionloadoptions#plotvisibleonly)|Pour chaque élément de la collection: true si seules les cellules visibles sont tracées.Faux si les deux cellules visibles et masquées sont tracées. Lecture/écriture.|
||[seriesNameLevel](/javascript/api/excel/excel.chartcollectionloadoptions#seriesnamelevel)|Pour chaque élément de la collection: renvoie ou définit une constante d’énumération ChartSeriesNameLevel faisant référence à|
||[ShowDataLabelsOverMaximum,](/javascript/api/excel/excel.chartcollectionloadoptions#showdatalabelsovermaximum)|Pour chaque élément de la collection: indique s’il faut afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale de l’axe des ordonnées.|
||[style](/javascript/api/excel/excel.chartcollectionloadoptions#style)|Pour chaque élément de la collection: cette propriété renvoie ou définit le style de graphique pour le graphique. Lecture/écriture.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[categoryLabelLevel](/javascript/api/excel/excel.chartdata#categorylabellevel)|Cette propriété renvoie ou définit une constante d’énumération ChartCategoryLabelLevel faisant référence à|
||[displayBlanksAs](/javascript/api/excel/excel.chartdata#displayblanksas)|Renvoie ou définit la façon dont les cellules vides sont tracées sur un graphique. Lecture/écriture.|
||[plotArea](/javascript/api/excel/excel.chartdata#plotarea)|Représente la zone de traçage pour le graphique.|
||[plotBy](/javascript/api/excel/excel.chartdata#plotby)|Renvoie ou spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Lecture/écriture.|
||[PlotVisibleOnly,](/javascript/api/excel/excel.chartdata#plotvisibleonly)|Vrai si seules les cellules visibles sont tracées.Faux si les deux cellules visibles et masquées sont tracées. Lecture/écriture.|
||[seriesNameLevel](/javascript/api/excel/excel.chartdata#seriesnamelevel)|Cette propriété renvoie ou définit une constante d’énumération ChartSeriesNameLevel faisant référence à|
||[ShowDataLabelsOverMaximum,](/javascript/api/excel/excel.chartdata#showdatalabelsovermaximum)|Si vous voulez afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale sur l’axe de valeur.|
||[style](/javascript/api/excel/excel.chartdata#style)|Cette propriété renvoie ou définit le style de graphique pour le graphique. Lecture/écriture.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabel#autotext)|Valeur booléenne représentant si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Représente le format d’étiquette de données graphique.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Représente la hauteur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Représente la largeur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Chaîne représentant le texte d’étiquette de données dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Représente l’orientation du texte de l’étiquette de données du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabeldata#autotext)|Valeur booléenne représentant si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|
||[format](/javascript/api/excel/excel.chartdatalabeldata#format)|Représente le format d’étiquette de données graphique.|
||[formula](/javascript/api/excel/excel.chartdatalabeldata#formula)|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|
||[height](/javascript/api/excel/excel.chartdatalabeldata#height)|Représente la hauteur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabeldata#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.chartdatalabeldata#left)|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabeldata#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|
||[text](/javascript/api/excel/excel.chartdatalabeldata#text)|Chaîne représentant le texte d’étiquette de données dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabeldata#textorientation)|Représente l’orientation du texte de l’étiquette de données du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.chartdatalabeldata#top)|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabeldata#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
||[width](/javascript/api/excel/excel.chartdatalabeldata#width)|Représente la largeur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[route](/javascript/api/excel/excel.chartdatalabelformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[route](/javascript/api/excel/excel.chartdatalabelformatdata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[route](/javascript/api/excel/excel.chartdatalabelformatloadoptions#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[route](/javascript/api/excel/excel.chartdatalabelformatupdatedata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabelloadoptions#autotext)|Valeur booléenne représentant si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|
||[format](/javascript/api/excel/excel.chartdatalabelloadoptions#format)|Représente le format d’étiquette de données graphique.|
||[formula](/javascript/api/excel/excel.chartdatalabelloadoptions#formula)|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|
||[height](/javascript/api/excel/excel.chartdatalabelloadoptions#height)|Représente la hauteur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.chartdatalabelloadoptions#left)|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|
||[text](/javascript/api/excel/excel.chartdatalabelloadoptions#text)|Chaîne représentant le texte d’étiquette de données dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelloadoptions#textorientation)|Représente l’orientation du texte de l’étiquette de données du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.chartdatalabelloadoptions#top)|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
||[width](/javascript/api/excel/excel.chartdatalabelloadoptions#width)|Représente la largeur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible.|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabelupdatedata#autotext)|Valeur booléenne représentant si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|
||[format](/javascript/api/excel/excel.chartdatalabelupdatedata#format)|Représente le format d’étiquette de données graphique.|
||[formula](/javascript/api/excel/excel.chartdatalabelupdatedata#formula)|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.chartdatalabelupdatedata#left)|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|
||[text](/javascript/api/excel/excel.chartdatalabelupdatedata#text)|Chaîne représentant le texte d’étiquette de données dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelupdatedata#textorientation)|Représente l’orientation du texte de l’étiquette de données du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.chartdatalabelupdatedata#top)|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabels#autotext)|Représente si des étiquettes de données génèrent automatiquement le texte approprié en fonction du contexte.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Représente le code de format pour les étiquettes de données.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Représente l’orientation du texte des étiquettes de données. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabelsdata#autotext)|Représente si des étiquettes de données génèrent automatiquement le texte approprié en fonction du contexte.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsdata#numberformat)|Représente le code de format pour les étiquettes de données.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsdata#textorientation)|Représente l’orientation du texte des étiquettes de données. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabelsloadoptions#autotext)|Représente si des étiquettes de données génèrent automatiquement le texte approprié en fonction du contexte.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#numberformat)|Représente le code de format pour les étiquettes de données.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsloadoptions#textorientation)|Représente l’orientation du texte des étiquettes de données. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabelsupdatedata#autotext)|Représente si des étiquettes de données génèrent automatiquement le texte approprié en fonction du contexte.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#numberformat)|Représente le code de format pour les étiquettes de données.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsupdatedata#textorientation)|Représente l’orientation du texte des étiquettes de données. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Obtient l’id du graphique qui est desactivé.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est desactivé.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Obtient l’id du graphique qui est supprimé de la feuille de calcul.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est supprimé.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Représente la hauteur de legendEntry sur la légende du graphique.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Représente l’index de legendEntry sur la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Représente la partie gauche d’un graphique legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Représente la partie supérieure d’un graphique legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Représente la largeur de legendEntry sur la légende d’un graphique.|
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[height](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#height)|Pour chaque élément de la collection: représente la hauteur de legendEntry sur la légende du graphique.|
||[index](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#index)|Pour chaque élément de la collection: représente l’index de legendEntry dans la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#left)|Pour chaque élément de la collection: représente la gauche d’un graphique legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#top)|Pour chaque élément de la collection: représente le bord supérieur d’un graphique legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#width)|Pour chaque élément de la collection: représente la largeur de legendEntry sur la légende du graphique.|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[height](/javascript/api/excel/excel.chartlegendentrydata#height)|Représente la hauteur de legendEntry sur la légende du graphique.|
||[index](/javascript/api/excel/excel.chartlegendentrydata#index)|Représente l’index de legendEntry sur la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegendentrydata#left)|Représente la partie gauche d’un graphique legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentrydata#top)|Représente la partie supérieure d’un graphique legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentrydata#width)|Représente la largeur de legendEntry sur la légende d’un graphique.|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[height](/javascript/api/excel/excel.chartlegendentryloadoptions#height)|Représente la hauteur de legendEntry sur la légende du graphique.|
||[index](/javascript/api/excel/excel.chartlegendentryloadoptions#index)|Représente l’index de legendEntry sur la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegendentryloadoptions#left)|Représente la partie gauche d’un graphique legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentryloadoptions#top)|Représente la partie supérieure d’un graphique legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentryloadoptions#width)|Représente la largeur de legendEntry sur la légende d’un graphique.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[route](/javascript/api/excel/excel.chartlegendformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[route](/javascript/api/excel/excel.chartlegendformatdata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[route](/javascript/api/excel/excel.chartlegendformatloadoptions#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[route](/javascript/api/excel/excel.chartlegendformatupdatedata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartloadoptions#categorylabellevel)|Cette propriété renvoie ou définit une constante d’énumération ChartCategoryLabelLevel faisant référence à|
||[displayBlanksAs](/javascript/api/excel/excel.chartloadoptions#displayblanksas)|Renvoie ou définit la façon dont les cellules vides sont tracées sur un graphique. Lecture/écriture.|
||[plotArea](/javascript/api/excel/excel.chartloadoptions#plotarea)|Représente la zone de traçage pour le graphique.|
||[plotBy](/javascript/api/excel/excel.chartloadoptions#plotby)|Renvoie ou spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Lecture/écriture.|
||[PlotVisibleOnly,](/javascript/api/excel/excel.chartloadoptions#plotvisibleonly)|Vrai si seules les cellules visibles sont tracées.Faux si les deux cellules visibles et masquées sont tracées. Lecture/écriture.|
||[seriesNameLevel](/javascript/api/excel/excel.chartloadoptions#seriesnamelevel)|Cette propriété renvoie ou définit une constante d’énumération ChartSeriesNameLevel faisant référence à|
||[ShowDataLabelsOverMaximum,](/javascript/api/excel/excel.chartloadoptions#showdatalabelsovermaximum)|Si vous voulez afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale sur l’axe de valeur.|
||[style](/javascript/api/excel/excel.chartloadoptions#style)|Cette propriété renvoie ou définit le style de graphique pour le graphique. Lecture/écriture.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Représente la valeur de hauteur de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Représente la valeur insideHeight de plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Représente la valeur insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Représente la valeur insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Représente la valeur insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Représente la valeur gauche de plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Représentant la position de plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Représente la mise en forme d’un graphique plotArea.|
||[Set (propriétés: Excel. ChartPlotArea)](/javascript/api/excel/excel.chartplotarea#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartPlotAreaUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartplotarea#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Représente la valeur supérieure de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Représente la valeur de largeur de plotArea.|
|[ChartPlotAreaData](/javascript/api/excel/excel.chartplotareadata)|[format](/javascript/api/excel/excel.chartplotareadata#format)|Représente la mise en forme d’un graphique plotArea.|
||[height](/javascript/api/excel/excel.chartplotareadata#height)|Représente la valeur de hauteur de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareadata#insideheight)|Représente la valeur insideHeight de plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareadata#insideleft)|Représente la valeur insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareadata#insidetop)|Représente la valeur insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareadata#insidewidth)|Représente la valeur insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotareadata#left)|Représente la valeur gauche de plotArea.|
||[position](/javascript/api/excel/excel.chartplotareadata#position)|Représentant la position de plotArea.|
||[top](/javascript/api/excel/excel.chartplotareadata#top)|Représente la valeur supérieure de plotArea.|
||[width](/javascript/api/excel/excel.chartplotareadata#width)|Représente la valeur de largeur de plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[route](/javascript/api/excel/excel.chartplotareaformat#border)|Représente les attributs de bordure d’un graphique plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[Set (propriétés: Excel. ChartPlotAreaFormat)](/javascript/api/excel/excel.chartplotareaformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartPlotAreaFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.chartplotareaformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartPlotAreaFormatData](/javascript/api/excel/excel.chartplotareaformatdata)|[route](/javascript/api/excel/excel.chartplotareaformatdata#border)|Représente les attributs de bordure d’un graphique plotArea.|
|[ChartPlotAreaFormatLoadOptions](/javascript/api/excel/excel.chartplotareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartplotareaformatloadoptions#$all)||
||[route](/javascript/api/excel/excel.chartplotareaformatloadoptions#border)|Représente les attributs de bordure d’un graphique plotArea.|
|[ChartPlotAreaFormatUpdateData](/javascript/api/excel/excel.chartplotareaformatupdatedata)|[route](/javascript/api/excel/excel.chartplotareaformatupdatedata#border)|Représente les attributs de bordure d’un graphique plotArea.|
|[ChartPlotAreaLoadOptions](/javascript/api/excel/excel.chartplotarealoadoptions)|[$all](/javascript/api/excel/excel.chartplotarealoadoptions#$all)||
||[format](/javascript/api/excel/excel.chartplotarealoadoptions#format)|Représente la mise en forme d’un graphique plotArea.|
||[height](/javascript/api/excel/excel.chartplotarealoadoptions#height)|Représente la valeur de hauteur de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarealoadoptions#insideheight)|Représente la valeur insideHeight de plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarealoadoptions#insideleft)|Représente la valeur insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarealoadoptions#insidetop)|Représente la valeur insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarealoadoptions#insidewidth)|Représente la valeur insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarealoadoptions#left)|Représente la valeur gauche de plotArea.|
||[position](/javascript/api/excel/excel.chartplotarealoadoptions#position)|Représentant la position de plotArea.|
||[top](/javascript/api/excel/excel.chartplotarealoadoptions#top)|Représente la valeur supérieure de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarealoadoptions#width)|Représente la valeur de largeur de plotArea.|
|[ChartPlotAreaUpdateData](/javascript/api/excel/excel.chartplotareaupdatedata)|[format](/javascript/api/excel/excel.chartplotareaupdatedata#format)|Représente la mise en forme d’un graphique plotArea.|
||[height](/javascript/api/excel/excel.chartplotareaupdatedata#height)|Représente la valeur de hauteur de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareaupdatedata#insideheight)|Représente la valeur insideHeight de plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareaupdatedata#insideleft)|Représente la valeur insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareaupdatedata#insidetop)|Représente la valeur insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareaupdatedata#insidewidth)|Représente la valeur insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotareaupdatedata#left)|Représente la valeur gauche de plotArea.|
||[position](/javascript/api/excel/excel.chartplotareaupdatedata#position)|Représentant la position de plotArea.|
||[top](/javascript/api/excel/excel.chartplotareaupdatedata#top)|Représente la valeur supérieure de plotArea.|
||[width](/javascript/api/excel/excel.chartplotareaupdatedata#width)|Représente la valeur de largeur de plotArea.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Renvoie ou définit le groupe de la série spécifiée. Lecture/Écriture|
||[barrage](/javascript/api/excel/excel.chartseries#explosion)|Renvoie ou définit la valeur d’explosion pour une coupe de graphique en secteurs ou de graphique en anneaux. Renvoie 0 (zéro) s’il n’y a aucune explosion (la pointe de la coupe est dans le centre du graphique). Lecture/écriture.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Renvoie ou définit l’angle de la première coupe graphique en secteurs ou graphique en anneaux, en degrés (dans le sens des aiguilles d’une montre, vertical). S’applique uniquement aux graphiques en secteurs, graphiques en secteurs 3D et graphiques en anneaux. Peut être une valeur comprise entre 0 et 360. Lecture/Écriture|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|Vrai si Microsoft Excel inverse le motif dans l’élément lorsqu’il correspond à un nombre négatif. Lecture/écriture.|
||[coin](/javascript/api/excel/excel.chartseries#overlap)|Spécifie comment barres et colonnes sont positionnées. Peut être une valeur comprise entre – 100 et 100. S’applique uniquement aux graphiques en barres 2D et en colonnes 2D. Lecture/écriture.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Représente la collection de tous les dataLabels de la série.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Renvoie ou définit la taille de la section secondaire d’un secteur de graphique en secteurs ou d’une barre de graphique en secteurs, sous forme de pourcentage de la taille du graphique principal. Peut être une valeur comprise entre 5 et 200. Lecture/écriture.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Renvoie ou définit la façon dont les deux sections d’un secteur de graphique en secteurs ou de barre d’un graphique en secteurs sont réparties. Lecture/écriture.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|Vrai si Microsoft Excel affecte une couleur ou un motif différentes à chaque marqueur de données. Le graphique ne doit contenir qu’une seule série. Lecture/écriture.|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriescollectionloadoptions#axisgroup)|Pour chaque élément de la collection: cette propriété renvoie ou définit le groupe de la série spécifiée. Lecture/Écriture|
||[dataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#datalabels)|Pour chaque élément de la collection: représente une collection de tous les dataLabels de la série.|
||[barrage](/javascript/api/excel/excel.chartseriescollectionloadoptions#explosion)|Pour chaque élément de la collection: cette propriété renvoie ou définit la valeur d’éclatement d’un graphique en secteurs ou d’un graphique en anneau. Renvoie 0 (zéro) s’il n’y a aucune explosion (la pointe de la coupe est dans le centre du graphique). Lecture/écriture.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriescollectionloadoptions#firstsliceangle)|Pour chaque élément de la collection: cette propriété renvoie ou définit l’angle du premier secteur de graphique en secteurs ou graphique en anneau, en degrés (dans le sens des aiguilles d’une montre). S’applique uniquement aux graphiques en secteurs, graphiques en secteurs 3D et graphiques en anneaux. Peut être une valeur comprise entre 0 et 360. Lecture/Écriture|
||[invertIfNegative](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertifnegative)|Pour chaque élément de la collection: true si Microsoft Excel inverse le motif de l’élément lorsqu’il correspond à un nombre négatif. Lecture/écriture.|
||[coin](/javascript/api/excel/excel.chartseriescollectionloadoptions#overlap)|Pour chaque élément de la collection: spécifie la position des barres et des colonnes. Peut être une valeur comprise entre – 100 et 100. S’applique uniquement aux graphiques en barres 2D et en colonnes 2D. Lecture/écriture.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#secondplotsize)|Pour chaque élément de la collection: cette propriété renvoie ou définit la taille de la section secondaire d’un graphique en secteurs de secteur ou une barre de graphique en secteurs, sous la forme d’un pourcentage de la taille du secteur principal. Peut être une valeur comprise entre 5 et 200. Lecture/écriture.|
||[splitType](/javascript/api/excel/excel.chartseriescollectionloadoptions#splittype)|Pour chaque élément de la collection: cette propriété renvoie ou définit la façon dont les deux sections d’un secteur d’un graphique en secteurs ou d’une barre d’un graphique en secteurs sont fractionnées. Lecture/écriture.|
||[varyByCategories](/javascript/api/excel/excel.chartseriescollectionloadoptions#varybycategories)|Pour chaque élément de la collection: cette propriété a la valeur true si Microsoft Excel affecte une couleur ou un motif différent à chaque marqueur de données. Le graphique ne doit contenir qu’une seule série. Lecture/écriture.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[axisGroup](/javascript/api/excel/excel.chartseriesdata#axisgroup)|Renvoie ou définit le groupe de la série spécifiée. Lecture/Écriture|
||[dataLabels](/javascript/api/excel/excel.chartseriesdata#datalabels)|Représente la collection de tous les dataLabels de la série.|
||[barrage](/javascript/api/excel/excel.chartseriesdata#explosion)|Renvoie ou définit la valeur d’explosion pour une coupe de graphique en secteurs ou de graphique en anneaux. Renvoie 0 (zéro) s’il n’y a aucune explosion (la pointe de la coupe est dans le centre du graphique). Lecture/écriture.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesdata#firstsliceangle)|Renvoie ou définit l’angle de la première coupe graphique en secteurs ou graphique en anneaux, en degrés (dans le sens des aiguilles d’une montre, vertical). S’applique uniquement aux graphiques en secteurs, graphiques en secteurs 3D et graphiques en anneaux. Peut être une valeur comprise entre 0 et 360. Lecture/Écriture|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesdata#invertifnegative)|Vrai si Microsoft Excel inverse le motif dans l’élément lorsqu’il correspond à un nombre négatif. Lecture/écriture.|
||[coin](/javascript/api/excel/excel.chartseriesdata#overlap)|Spécifie comment barres et colonnes sont positionnées. Peut être une valeur comprise entre – 100 et 100. S’applique uniquement aux graphiques en barres 2D et en colonnes 2D. Lecture/écriture.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesdata#secondplotsize)|Renvoie ou définit la taille de la section secondaire d’un secteur de graphique en secteurs ou d’une barre de graphique en secteurs, sous forme de pourcentage de la taille du graphique principal. Peut être une valeur comprise entre 5 et 200. Lecture/écriture.|
||[splitType](/javascript/api/excel/excel.chartseriesdata#splittype)|Renvoie ou définit la façon dont les deux sections d’un secteur de graphique en secteurs ou de barre d’un graphique en secteurs sont réparties. Lecture/écriture.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesdata#varybycategories)|Vrai si Microsoft Excel affecte une couleur ou un motif différentes à chaque marqueur de données. Le graphique ne doit contenir qu’une seule série. Lecture/écriture.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriesloadoptions#axisgroup)|Renvoie ou définit le groupe de la série spécifiée. Lecture/Écriture|
||[dataLabels](/javascript/api/excel/excel.chartseriesloadoptions#datalabels)|Représente la collection de tous les dataLabels de la série.|
||[barrage](/javascript/api/excel/excel.chartseriesloadoptions#explosion)|Renvoie ou définit la valeur d’explosion pour une coupe de graphique en secteurs ou de graphique en anneaux. Renvoie 0 (zéro) s’il n’y a aucune explosion (la pointe de la coupe est dans le centre du graphique). Lecture/écriture.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesloadoptions#firstsliceangle)|Renvoie ou définit l’angle de la première coupe graphique en secteurs ou graphique en anneaux, en degrés (dans le sens des aiguilles d’une montre, vertical). S’applique uniquement aux graphiques en secteurs, graphiques en secteurs 3D et graphiques en anneaux. Peut être une valeur comprise entre 0 et 360. Lecture/Écriture|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesloadoptions#invertifnegative)|Vrai si Microsoft Excel inverse le motif dans l’élément lorsqu’il correspond à un nombre négatif. Lecture/écriture.|
||[coin](/javascript/api/excel/excel.chartseriesloadoptions#overlap)|Spécifie comment barres et colonnes sont positionnées. Peut être une valeur comprise entre – 100 et 100. S’applique uniquement aux graphiques en barres 2D et en colonnes 2D. Lecture/écriture.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesloadoptions#secondplotsize)|Renvoie ou définit la taille de la section secondaire d’un secteur de graphique en secteurs ou d’une barre de graphique en secteurs, sous forme de pourcentage de la taille du graphique principal. Peut être une valeur comprise entre 5 et 200. Lecture/écriture.|
||[splitType](/javascript/api/excel/excel.chartseriesloadoptions#splittype)|Renvoie ou définit la façon dont les deux sections d’un secteur de graphique en secteurs ou de barre d’un graphique en secteurs sont réparties. Lecture/écriture.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesloadoptions#varybycategories)|Vrai si Microsoft Excel affecte une couleur ou un motif différentes à chaque marqueur de données. Le graphique ne doit contenir qu’une seule série. Lecture/écriture.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[axisGroup](/javascript/api/excel/excel.chartseriesupdatedata#axisgroup)|Renvoie ou définit le groupe de la série spécifiée. Lecture/Écriture|
||[dataLabels](/javascript/api/excel/excel.chartseriesupdatedata#datalabels)|Représente la collection de tous les dataLabels de la série.|
||[barrage](/javascript/api/excel/excel.chartseriesupdatedata#explosion)|Renvoie ou définit la valeur d’explosion pour une coupe de graphique en secteurs ou de graphique en anneaux. Renvoie 0 (zéro) s’il n’y a aucune explosion (la pointe de la coupe est dans le centre du graphique). Lecture/écriture.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesupdatedata#firstsliceangle)|Renvoie ou définit l’angle de la première coupe graphique en secteurs ou graphique en anneaux, en degrés (dans le sens des aiguilles d’une montre, vertical). S’applique uniquement aux graphiques en secteurs, graphiques en secteurs 3D et graphiques en anneaux. Peut être une valeur comprise entre 0 et 360. Lecture/Écriture|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesupdatedata#invertifnegative)|Vrai si Microsoft Excel inverse le motif dans l’élément lorsqu’il correspond à un nombre négatif. Lecture/écriture.|
||[coin](/javascript/api/excel/excel.chartseriesupdatedata#overlap)|Spécifie comment barres et colonnes sont positionnées. Peut être une valeur comprise entre – 100 et 100. S’applique uniquement aux graphiques en barres 2D et en colonnes 2D. Lecture/écriture.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesupdatedata#secondplotsize)|Renvoie ou définit la taille de la section secondaire d’un secteur de graphique en secteurs ou d’une barre de graphique en secteurs, sous forme de pourcentage de la taille du graphique principal. Peut être une valeur comprise entre 5 et 200. Lecture/écriture.|
||[splitType](/javascript/api/excel/excel.chartseriesupdatedata#splittype)|Renvoie ou définit la façon dont les deux sections d’un secteur de graphique en secteurs ou de barre d’un graphique en secteurs sont réparties. Lecture/écriture.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesupdatedata#varybycategories)|Vrai si Microsoft Excel affecte une couleur ou un motif différentes à chaque marqueur de données. Le graphique ne doit contenir qu’une seule série. Lecture/écriture.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Représente le nombre de points que la courbe de tendance étend en arrière.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Représente le nombre de points que la courbe de tendance étend en avant.|
||[labellisé](/javascript/api/excel/excel.charttrendline#label)|Représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#backwardperiod)|Pour chaque élément de la collection: représente le nombre de périodes que la courbe de tendance étend en rétrospective.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#forwardperiod)|Pour chaque élément de la collection: représente le nombre de périodes que la courbe de tendance étend en prospective.|
||[labellisé](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#label)|Pour chaque élément de la collection: représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showequation)|Pour chaque élément de la collection: true si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showrsquared)|Pour chaque élément de la collection: true si le carré de la courbe de tendance est affiché sur le graphique.|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinedata#backwardperiod)|Représente le nombre de points que la courbe de tendance étend en arrière.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinedata#forwardperiod)|Représente le nombre de points que la courbe de tendance étend en avant.|
||[labellisé](/javascript/api/excel/excel.charttrendlinedata#label)|Représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendlinedata#showequation)|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendlinedata#showrsquared)|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[Insertion automatique](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Valeur booléenne représentant si l’étiquette de tendances génère automatiquement le texte approprié en fonction du contexte.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Valeur de chaîne qui représente la formule de l’étiquette de tendances du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Représente la distance en points, du bord gauche de l’étiquette de tendances graphique au bord gauche de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de tendances.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Représente le format d’étiquette de tendances du graphique.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Représente la hauteur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Représente la largeur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible.|
||[Set (propriétés: Excel. ChartTrendlineLabel)](/javascript/api/excel/excel.charttrendlinelabel#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartTrendlineLabelUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabel#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Chaîne représentant le texte d’étiquette de tendances dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Représente l’orientation du texte de l’étiquette de tendances du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Représente la distance en points, du bord supérieur de l’étiquette de tendances du graphique au bord supérieur de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Représente l’alignement vertical de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[Insertion automatique](/javascript/api/excel/excel.charttrendlinelabeldata#autotext)|Valeur booléenne représentant si l’étiquette de tendances génère automatiquement le texte approprié en fonction du contexte.|
||[format](/javascript/api/excel/excel.charttrendlinelabeldata#format)|Représente le format d’étiquette de tendances du graphique.|
||[formula](/javascript/api/excel/excel.charttrendlinelabeldata#formula)|Valeur de chaîne qui représente la formule de l’étiquette de tendances du graphique à l’aide de la notation de style A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabeldata#height)|Représente la hauteur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.charttrendlinelabeldata#left)|Représente la distance en points, du bord gauche de l’étiquette de tendances graphique au bord gauche de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de tendances.|
||[text](/javascript/api/excel/excel.charttrendlinelabeldata#text)|Chaîne représentant le texte d’étiquette de tendances dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabeldata#textorientation)|Représente l’orientation du texte de l’étiquette de tendances du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttrendlinelabeldata#top)|Représente la distance en points, du bord supérieur de l’étiquette de tendances du graphique au bord supérieur de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#verticalalignment)|Représente l’alignement vertical de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
||[width](/javascript/api/excel/excel.charttrendlinelabeldata#width)|Représente la largeur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[route](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Représente le format de remplissage de l’étiquette de tendances du graphique actuel.|
||[police](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de tendances de graphique.|
||[Set (propriétés: Excel. ChartTrendlineLabelFormat)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. ChartTrendlineLabelFormatUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[ChartTrendlineLabelFormatData](/javascript/api/excel/excel.charttrendlinelabelformatdata)|[route](/javascript/api/excel/excel.charttrendlinelabelformatdata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
||[police](/javascript/api/excel/excel.charttrendlinelabelformatdata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de tendances de graphique.|
|[ChartTrendlineLabelFormatLoadOptions](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#$all)||
||[route](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
||[police](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de tendances de graphique.|
|[ChartTrendlineLabelFormatUpdateData](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata)|[route](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
||[police](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#font)|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de tendances de graphique.|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelloadoptions#$all)||
||[Insertion automatique](/javascript/api/excel/excel.charttrendlinelabelloadoptions#autotext)|Valeur booléenne représentant si l’étiquette de tendances génère automatiquement le texte approprié en fonction du contexte.|
||[format](/javascript/api/excel/excel.charttrendlinelabelloadoptions#format)|Représente le format d’étiquette de tendances du graphique.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelloadoptions#formula)|Valeur de chaîne qui représente la formule de l’étiquette de tendances du graphique à l’aide de la notation de style A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabelloadoptions#height)|Représente la hauteur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.charttrendlinelabelloadoptions#left)|Représente la distance en points, du bord gauche de l’étiquette de tendances graphique au bord gauche de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de tendances.|
||[text](/javascript/api/excel/excel.charttrendlinelabelloadoptions#text)|Chaîne représentant le texte d’étiquette de tendances dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelloadoptions#textorientation)|Représente l’orientation du texte de l’étiquette de tendances du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttrendlinelabelloadoptions#top)|Représente la distance en points, du bord supérieur de l’étiquette de tendances du graphique au bord supérieur de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#verticalalignment)|Représente l’alignement vertical de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
||[width](/javascript/api/excel/excel.charttrendlinelabelloadoptions#width)|Représente la largeur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible.|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[Insertion automatique](/javascript/api/excel/excel.charttrendlinelabelupdatedata#autotext)|Valeur booléenne représentant si l’étiquette de tendances génère automatiquement le texte approprié en fonction du contexte.|
||[format](/javascript/api/excel/excel.charttrendlinelabelupdatedata#format)|Représente le format d’étiquette de tendances du graphique.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelupdatedata#formula)|Valeur de chaîne qui représente la formule de l’étiquette de tendances du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextHorizontalAlignment.|
||[left](/javascript/api/excel/excel.charttrendlinelabelupdatedata#left)|Représente la distance en points, du bord gauche de l’étiquette de tendances graphique au bord gauche de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de tendances.|
||[text](/javascript/api/excel/excel.charttrendlinelabelupdatedata#text)|Chaîne représentant le texte d’étiquette de tendances dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelupdatedata#textorientation)|Représente l’orientation du texte de l’étiquette de tendances du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|
||[top](/javascript/api/excel/excel.charttrendlinelabelupdatedata#top)|Représente la distance en points, du bord supérieur de l’étiquette de tendances du graphique au bord supérieur de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#verticalalignment)|Représente l’alignement vertical de l’étiquette de tendances du graphique. Pour plus d’informations, voir Excel. ChartTextVerticalAlignment.|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#backwardperiod)|Représente le nombre de points que la courbe de tendance étend en arrière.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#forwardperiod)|Représente le nombre de points que la courbe de tendance étend en avant.|
||[labellisé](/javascript/api/excel/excel.charttrendlineloadoptions#label)|Représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendlineloadoptions#showequation)|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendlineloadoptions#showrsquared)|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#backwardperiod)|Représente le nombre de points que la courbe de tendance étend en arrière.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#forwardperiod)|Représente le nombre de points que la courbe de tendance étend en avant.|
||[labellisé](/javascript/api/excel/excel.charttrendlineupdatedata#label)|Représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendlineupdatedata#showequation)|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendlineupdatedata#showrsquared)|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[categoryLabelLevel](/javascript/api/excel/excel.chartupdatedata#categorylabellevel)|Cette propriété renvoie ou définit une constante d’énumération ChartCategoryLabelLevel faisant référence à|
||[displayBlanksAs](/javascript/api/excel/excel.chartupdatedata#displayblanksas)|Renvoie ou définit la façon dont les cellules vides sont tracées sur un graphique. Lecture/écriture.|
||[plotArea](/javascript/api/excel/excel.chartupdatedata#plotarea)|Représente la zone de traçage pour le graphique.|
||[plotBy](/javascript/api/excel/excel.chartupdatedata#plotby)|Renvoie ou spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Lecture/écriture.|
||[PlotVisibleOnly,](/javascript/api/excel/excel.chartupdatedata#plotvisibleonly)|Vrai si seules les cellules visibles sont tracées.Faux si les deux cellules visibles et masquées sont tracées. Lecture/écriture.|
||[seriesNameLevel](/javascript/api/excel/excel.chartupdatedata#seriesnamelevel)|Cette propriété renvoie ou définit une constante d’énumération ChartSeriesNameLevel faisant référence à|
||[ShowDataLabelsOverMaximum,](/javascript/api/excel/excel.chartupdatedata#showdatalabelsovermaximum)|Si vous voulez afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale sur l’axe de valeur.|
||[style](/javascript/api/excel/excel.chartupdatedata#style)|Cette propriété renvoie ou définit le style de graphique pour le graphique. Lecture/écriture.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Une formule de validation des données personnalisée. Cette opération crée des règles d’entrée spéciales, telles que la prévention des doublons ou la limitation du total dans une plage de cellules.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Nom de la DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Format de nombre de la DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Position de la DataPivotHierarchy.|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|Renvoie les PivotFields associés à la DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID de la DataPivotHierarchy.|
||[Set (propriétés: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchy#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. DataPivotHierarchyUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.datapivothierarchy#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Restaurer la DataPivotHierarchy à ses valeurs par défaut.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Détermine si les données doivent apparaître sous forme d’un calcul de synthèse spécifique ou non.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Détermine si vous voulez afficher tous les éléments de la DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[Ajouter (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Ajoute le PivotHierarchy à l’axe en cours.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Obtient une DataPivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Obtient une DataPivotHierarchy par nom. Si la DataPivotHierarchy n’existe pas, cela renvoie un objet null.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[Remove (DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Supprime le PivotHierarchy de l’axe en cours.|
|[DataPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#field)|Pour chaque élément de la collection: renvoie les PivotFields associés au DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#id)|Pour chaque élément de la collection: ID du DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#name)|Pour chaque élément de la collection: nom du DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#numberformat)|Pour chaque élément de la collection: format de numéro du DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#position)|Pour chaque élément de la collection: position du DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#showas)|Pour chaque élément de la collection: détermine si les données doivent être affichées sous la forme d’un calcul de synthèse spécifique ou non.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#summarizeby)|Pour chaque élément de la collection: détermine s’il faut afficher tous les éléments du DataPivotHierarchy.|
|[DataPivotHierarchyData](/javascript/api/excel/excel.datapivothierarchydata)|[field](/javascript/api/excel/excel.datapivothierarchydata#field)|Renvoie les PivotFields associés à la DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchydata#id)|ID de la DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchydata#name)|Nom de la DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchydata#numberformat)|Format de nombre de la DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchydata#position)|Position de la DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchydata#showas)|Détermine si les données doivent apparaître sous forme d’un calcul de synthèse spécifique ou non.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchydata#summarizeby)|Détermine si vous voulez afficher tous les éléments de la DataPivotHierarchy.|
|[DataPivotHierarchyLoadOptions](/javascript/api/excel/excel.datapivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchyloadoptions#$all)||
||[field](/javascript/api/excel/excel.datapivothierarchyloadoptions#field)|Renvoie les PivotFields associés à la DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchyloadoptions#id)|ID de la DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyloadoptions#name)|Nom de la DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyloadoptions#numberformat)|Format de nombre de la DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyloadoptions#position)|Position de la DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyloadoptions#showas)|Détermine si les données doivent apparaître sous forme d’un calcul de synthèse spécifique ou non.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyloadoptions#summarizeby)|Détermine si vous voulez afficher tous les éléments de la DataPivotHierarchy.|
|[DataPivotHierarchyUpdateData](/javascript/api/excel/excel.datapivothierarchyupdatedata)|[field](/javascript/api/excel/excel.datapivothierarchyupdatedata#field)|Renvoie les PivotFields associés à la DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyupdatedata#name)|Nom de la DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyupdatedata#numberformat)|Format de nombre de la DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyupdatedata#position)|Position de la DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyupdatedata#showas)|Détermine si les données doivent apparaître sous forme d’un calcul de synthèse spécifique ou non.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyupdatedata#summarizeby)|Détermine si vous voulez afficher tous les éléments de la DataPivotHierarchy.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Efface la validation des données de la plage active.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Alerte d’erreur lorsque l’utilisateur entre des données non valides.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Ignorer les espaces vides : aucune validation des données ne sera exécutée sur les cellules vides, la valeur par défaut est vrai.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Invite lorsque les utilisateurs sélectionnent une cellule.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Type de validation des données, voir Excel.DataValidationType pour plus d’informations.|
||[validation](/javascript/api/excel/excel.datavalidation#valid)|Représente si toutes les valeurs de cellule sont valides selon les règles de validation des données.|
||[sous](/javascript/api/excel/excel.datavalidation#rule)|Règle de validation des données qui contient différents critères de validation de données.|
||[Set (propriétés: Excel. DataValidation)](/javascript/api/excel/excel.datavalidation#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. DataValidationUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.datavalidation#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[DataValidationData](/javascript/api/excel/excel.datavalidationdata)|[errorAlert](/javascript/api/excel/excel.datavalidationdata#erroralert)|Alerte d’erreur lorsque l’utilisateur entre des données non valides.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationdata#ignoreblanks)|Ignorer les espaces vides : aucune validation des données ne sera exécutée sur les cellules vides, la valeur par défaut est vrai.|
||[prompt](/javascript/api/excel/excel.datavalidationdata#prompt)|Invite lorsque les utilisateurs sélectionnent une cellule.|
||[sous](/javascript/api/excel/excel.datavalidationdata#rule)|Règle de validation des données qui contient différents critères de validation de données.|
||[type](/javascript/api/excel/excel.datavalidationdata#type)|Type de validation des données, voir Excel.DataValidationType pour plus d’informations.|
||[validation](/javascript/api/excel/excel.datavalidationdata#valid)|Représente si toutes les valeurs de cellule sont valides selon les règles de validation des données.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Représente le message d’alerte d’erreur.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Détermine si vous voulez afficher un dialogue Alerte d’erreur ou pas lorsqu’un utilisateur entre des données non valides. La valeur par défaut est True.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Représente un type d’alerte de validation des données, voir Excel.DataValidationAlertStyle pour plus d’informations.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Représente le titre de dialogue d’alerte d’erreur.|
|[DataValidationLoadOptions](/javascript/api/excel/excel.datavalidationloadoptions)|[$all](/javascript/api/excel/excel.datavalidationloadoptions#$all)||
||[errorAlert](/javascript/api/excel/excel.datavalidationloadoptions#erroralert)|Alerte d’erreur lorsque l’utilisateur entre des données non valides.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationloadoptions#ignoreblanks)|Ignorer les espaces vides : aucune validation des données ne sera exécutée sur les cellules vides, la valeur par défaut est vrai.|
||[prompt](/javascript/api/excel/excel.datavalidationloadoptions#prompt)|Invite lorsque les utilisateurs sélectionnent une cellule.|
||[sous](/javascript/api/excel/excel.datavalidationloadoptions#rule)|Règle de validation des données qui contient différents critères de validation de données.|
||[type](/javascript/api/excel/excel.datavalidationloadoptions#type)|Type de validation des données, voir Excel.DataValidationType pour plus d’informations.|
||[validation](/javascript/api/excel/excel.datavalidationloadoptions#valid)|Représente si toutes les valeurs de cellule sont valides selon les règles de validation des données.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Représente le message de l’invite.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Détermine d’afficher ou non l’invite lorsqu’un utilisateur sélectionne une cellule avec validation des données.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Représente le titre de l’invite.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[personnalisé](/javascript/api/excel/excel.datavalidationrule#custom)|Critères de validation des données personnalisés.|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|Critères de validation des données de date.|
||[fixe](/javascript/api/excel/excel.datavalidationrule#decimal)|Critères de validation des données décimales.|
||[liste](/javascript/api/excel/excel.datavalidationrule#list)|Critères de validation des données de liste.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Critères de validation des données textLength.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Critères de validation des données de temps.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Critères de validation des données WholeNumber.|
|[DataValidationUpdateData](/javascript/api/excel/excel.datavalidationupdatedata)|[errorAlert](/javascript/api/excel/excel.datavalidationupdatedata#erroralert)|Alerte d’erreur lorsque l’utilisateur entre des données non valides.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationupdatedata#ignoreblanks)|Ignorer les espaces vides : aucune validation des données ne sera exécutée sur les cellules vides, la valeur par défaut est vrai.|
||[prompt](/javascript/api/excel/excel.datavalidationupdatedata#prompt)|Invite lorsque les utilisateurs sélectionnent une cellule.|
||[sous](/javascript/api/excel/excel.datavalidationupdatedata#rule)|Règle de validation des données qui contient différents critères de validation de données.|
|[Objetdatetimedatavalidation](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Spécifie l’opérande de droite lorsque la propriété Operator est définie sur un opérateur binaire tel que GreaterThan (l’opérande de gauche est la valeur que l’utilisateur tente d’entrer dans la cellule). Avec les opérateurs ternaires entre et NotBetween, spécifie l’opérande inférieur.|
||[Formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Avec les opérateurs ternaires entre et NotBetween, spécifie l’opérande supérieur. N’est pas utilisé avec les opérateurs binaires, tels que GreaterThan.|
||[is](/javascript/api/excel/excel.datetimedatavalidation#operator)|L’opérateur à utiliser pour la validation des données.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Détermine si vous voulez autoriser plusieurs éléments de filtre.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Nom du filterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Position du filterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Renvoie les PivotFields associés à la FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID du filterPivotHierarchy.|
||[Set (propriétés: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchy#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. FilterPivotHierarchyUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.filterpivothierarchy#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Restaurer la FilterPivotHierarchy à ses valeurs par défaut.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[Ajouter (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Ajoute le PivotHierarchy à l’axe en cours. Si la hiérarchie est présente ailleurs sur la ligne, colonne,|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Obtient une FilterPivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Obtient un FilterPivotHierarchy par nom. Si la FilterPivotHierarchy n’existe pas, cela renvoie un objet null.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[Remove (filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Supprime le PivotHierarchy de l’axe en cours.|
|[FilterPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#enablemultiplefilteritems)|Pour chaque élément de la collection: détermine s’il faut autoriser plusieurs éléments de filtre.|
||[id](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#id)|Pour chaque élément de la collection: ID du FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#name)|Pour chaque élément de la collection: nom du FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#position)|Pour chaque élément de la collection: position du FilterPivotHierarchy.|
|[FilterPivotHierarchyData](/javascript/api/excel/excel.filterpivothierarchydata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchydata#enablemultiplefilteritems)|Détermine si vous voulez autoriser plusieurs éléments de filtre.|
||[fields](/javascript/api/excel/excel.filterpivothierarchydata#fields)|Renvoie les PivotFields associés à la FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchydata#id)|ID du filterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchydata#name)|Nom du filterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchydata#position)|Position du filterPivotHierarchy.|
|[FilterPivotHierarchyLoadOptions](/javascript/api/excel/excel.filterpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchyloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyloadoptions#enablemultiplefilteritems)|Détermine si vous voulez autoriser plusieurs éléments de filtre.|
||[id](/javascript/api/excel/excel.filterpivothierarchyloadoptions#id)|ID du filterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchyloadoptions#name)|Nom du filterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyloadoptions#position)|Position du filterPivotHierarchy.|
|[FilterPivotHierarchyUpdateData](/javascript/api/excel/excel.filterpivothierarchyupdatedata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyupdatedata#enablemultiplefilteritems)|Détermine si vous voulez autoriser plusieurs éléments de filtre.|
||[name](/javascript/api/excel/excel.filterpivothierarchyupdatedata#name)|Nom du filterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyupdatedata#position)|Position du filterPivotHierarchy.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[InCellDropdown,](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Affiche la liste dans la cellule déroulante ou non, sa valeur par défaut est true.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Source de la liste de validation des données|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Nom du champ PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID du champ PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Renvoie les PivotItems qui composent avec le champ de tableau croisé dynamique.|
||[Set (propriétés: Excel. PivotField)](/javascript/api/excel/excel.pivotfield#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. PivotFieldUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.pivotfield#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[ShowAllItems,](/javascript/api/excel/excel.pivotfield#showallitems)|Détermine si vous voulez afficher tous les éléments de PivotField.|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Trie le PivotField. Si une DataPivotHierarchy est spécifiée, le tri sera appliqué en fonction de celle-ci, sinon le tri sera basé sur le PivotField lui-même.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Sous-totaux du champ PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Obtient le nombre de champs de tableau croisé dynamique dans la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Obtient un champ de tableau croisé dynamique par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Obtient un champ de tableau croisé dynamique par nom. Si le champ PivotField n’existe pas, un objet null est renvoyé.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotFieldCollectionLoadOptions](/javascript/api/excel/excel.pivotfieldcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#id)|Pour chaque élément de la collection: ID du champ de tableau croisé dynamique.|
||[name](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#name)|Pour chaque élément de la collection: nom du champ de tableau croisé dynamique.|
||[ShowAllItems,](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#showallitems)|Pour chaque élément de la collection: détermine si tous les éléments du champ PivotField doivent être affichés.|
||[subtotals](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#subtotals)|Pour chaque élément de la collection: sous-totaux du champ de tableau croisé dynamique.|
|[PivotFieldData](/javascript/api/excel/excel.pivotfielddata)|[id](/javascript/api/excel/excel.pivotfielddata#id)|ID du champ PivotField.|
||[items](/javascript/api/excel/excel.pivotfielddata#items)|Renvoie les PivotFields associés à PivotField.|
||[name](/javascript/api/excel/excel.pivotfielddata#name)|Nom du champ PivotField.|
||[ShowAllItems,](/javascript/api/excel/excel.pivotfielddata#showallitems)|Détermine si vous voulez afficher tous les éléments de PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfielddata#subtotals)|Sous-totaux du champ PivotField.|
|[PivotFieldLoadOptions](/javascript/api/excel/excel.pivotfieldloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldloadoptions#id)|ID du champ PivotField.|
||[name](/javascript/api/excel/excel.pivotfieldloadoptions#name)|Nom du champ PivotField.|
||[ShowAllItems,](/javascript/api/excel/excel.pivotfieldloadoptions#showallitems)|Détermine si vous voulez afficher tous les éléments de PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldloadoptions#subtotals)|Sous-totaux du champ PivotField.|
|[PivotFieldUpdateData](/javascript/api/excel/excel.pivotfieldupdatedata)|[name](/javascript/api/excel/excel.pivotfieldupdatedata#name)|Nom du champ PivotField.|
||[ShowAllItems,](/javascript/api/excel/excel.pivotfieldupdatedata#showallitems)|Détermine si vous voulez afficher tous les éléments de PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldupdatedata#subtotals)|Sous-totaux du champ PivotField.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Nom de la PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Renvoie les PivotFields associés à la PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID de la PivotHierarchy.|
||[Set (propriétés: Excel. PivotHierarchy)](/javascript/api/excel/excel.pivothierarchy#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. PivotHierarchyUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.pivothierarchy#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Obtient une PivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Obtient une PivotHierarchy par nom. Si la PivotHierarchy n’existe pas, cela renvoie un objet null.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.pivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#id)|Pour chaque élément de la collection: ID du PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#name)|Pour chaque élément de la collection: nom du PivotHierarchy.|
|[PivotHierarchyData](/javascript/api/excel/excel.pivothierarchydata)|[fields](/javascript/api/excel/excel.pivothierarchydata#fields)|Renvoie les PivotFields associés à la PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchydata#id)|ID de la PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchydata#name)|Nom de la PivotHierarchy.|
|[PivotHierarchyLoadOptions](/javascript/api/excel/excel.pivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchyloadoptions#id)|ID de la PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchyloadoptions#name)|Nom de la PivotHierarchy.|
|[PivotHierarchyUpdateData](/javascript/api/excel/excel.pivothierarchyupdatedata)|[name](/javascript/api/excel/excel.pivothierarchyupdatedata#name)|Nom de la PivotHierarchy.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[IsExpanded,](/javascript/api/excel/excel.pivotitem#isexpanded)|Détermine si l’élément est développé pour afficher les éléments enfants ou si ce dernier est réduit et les éléments enfants sont masqués.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Nom du champ PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID du champ PivotItem.|
||[Set (propriétés: Excel. PivotItem)](/javascript/api/excel/excel.pivotitem#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. PivotItemUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.pivotitem#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Détermine si le PivotItem est visible ou non.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Obtient le nombre d’éléments de tableau croisé dynamique dans la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Obtient un PivotItem par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Obtient un PivotItem par nom. Si PivotItem n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotItemCollectionLoadOptions](/javascript/api/excel/excel.pivotitemcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotitemcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemcollectionloadoptions#id)|Pour chaque élément de la collection: ID de PivotItem.|
||[IsExpanded,](/javascript/api/excel/excel.pivotitemcollectionloadoptions#isexpanded)|Pour chaque élément de la collection: détermine si l’élément est développé pour afficher les éléments enfants ou s’il est réduit et que les éléments enfants sont masqués.|
||[name](/javascript/api/excel/excel.pivotitemcollectionloadoptions#name)|Pour chaque élément de la collection: nom de l’objet PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemcollectionloadoptions#visible)|Pour chaque élément de la collection: détermine si l’objet PivotItem est visible ou non.|
|[PivotItemData](/javascript/api/excel/excel.pivotitemdata)|[id](/javascript/api/excel/excel.pivotitemdata#id)|ID du champ PivotItem.|
||[IsExpanded,](/javascript/api/excel/excel.pivotitemdata#isexpanded)|Détermine si l’élément est développé pour afficher les éléments enfants ou si ce dernier est réduit et les éléments enfants sont masqués.|
||[name](/javascript/api/excel/excel.pivotitemdata#name)|Nom du champ PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemdata#visible)|Détermine si le PivotItem est visible ou non.|
|[PivotItemLoadOptions](/javascript/api/excel/excel.pivotitemloadoptions)|[$all](/javascript/api/excel/excel.pivotitemloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemloadoptions#id)|ID du champ PivotItem.|
||[IsExpanded,](/javascript/api/excel/excel.pivotitemloadoptions#isexpanded)|Détermine si l’élément est développé pour afficher les éléments enfants ou si ce dernier est réduit et les éléments enfants sont masqués.|
||[name](/javascript/api/excel/excel.pivotitemloadoptions#name)|Nom du champ PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemloadoptions#visible)|Détermine si le PivotItem est visible ou non.|
|[PivotItemUpdateData](/javascript/api/excel/excel.pivotitemupdatedata)|[IsExpanded,](/javascript/api/excel/excel.pivotitemupdatedata#isexpanded)|Détermine si l’élément est développé pour afficher les éléments enfants ou si ce dernier est réduit et les éléments enfants sont masqués.|
||[name](/javascript/api/excel/excel.pivotitemupdatedata#name)|Nom du champ PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemupdatedata#visible)|Détermine si le PivotItem est visible ou non.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Renvoie la plage où les étiquettes de colonnes de tableau croisé dynamique se trouvent.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Renvoie la plage où les valeurs de données de tableau croisé dynamique se trouvent.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Renvoie la plage de la zone de filtre de tableau croisé dynamique.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Renvoie la plage sur laquelle le tableau croisé dynamique existe, à l’exception de la zone de filtre.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Renvoie la plage où les étiquettes de lignes de tableau croisé dynamique se trouvent.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Cette propriété indique le PivotLayoutType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
||[Set (propriétés: Excel. PivotLayout)](/javascript/api/excel/excel.pivotlayout#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. PivotLayoutUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.pivotlayout#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des colonnes.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des lignes.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Cette propriété indique le SubtotalLocationType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[layoutType](/javascript/api/excel/excel.pivotlayoutdata#layouttype)|Cette propriété indique le PivotLayoutType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showcolumngrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des colonnes.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showrowgrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des lignes.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutdata#subtotallocation)|Cette propriété indique le SubtotalLocationType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[$all](/javascript/api/excel/excel.pivotlayoutloadoptions#$all)||
||[layoutType](/javascript/api/excel/excel.pivotlayoutloadoptions#layouttype)|Cette propriété indique le PivotLayoutType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showcolumngrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des colonnes.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showrowgrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des lignes.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutloadoptions#subtotallocation)|Cette propriété indique le SubtotalLocationType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[layoutType](/javascript/api/excel/excel.pivotlayoutupdatedata#layouttype)|Cette propriété indique le PivotLayoutType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showcolumngrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des colonnes.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showrowgrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des lignes.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutupdatedata#subtotallocation)|Cette propriété indique le SubtotalLocationType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Supprime le tableau croisé dynamique.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|Les hiérarchies de colonne de tableau croisé dynamique.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|Les hiérarchies de données de tableau croisé dynamique.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|Les hiérarchies de filtre de tableau croisé dynamique.|
||[hiérarchies](/javascript/api/excel/excel.pivottable#hierarchies)|Les hiérarchies Pivot de tableau croisé dynamique.|
||[disposition](/javascript/api/excel/excel.pivottable#layout)|Le PivotLayout décrivant la disposition et la structure visuelle de tableau croisé dynamique.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|Les hiérarchies de lignes de tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (Name: chaîne, source: table \| de \| chaînes de plage, destination \| : chaîne de plage)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Ajoute un tableau croisé dynamique en fonction des données sources spécifiées et les insère à la cellule supérieure gauche de la plage de destination.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[disposition](/javascript/api/excel/excel.pivottablecollectionloadoptions#layout)|Pour chaque élément de la collection: PivotLayout décrivant la disposition et la structure visuelle du tableau croisé dynamique.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[columnHierarchies](/javascript/api/excel/excel.pivottabledata#columnhierarchies)|Les hiérarchies de colonne de tableau croisé dynamique.|
||[dataHierarchies](/javascript/api/excel/excel.pivottabledata#datahierarchies)|Les hiérarchies de données de tableau croisé dynamique.|
||[filterHierarchies](/javascript/api/excel/excel.pivottabledata#filterhierarchies)|Les hiérarchies de filtre de tableau croisé dynamique.|
||[hiérarchies](/javascript/api/excel/excel.pivottabledata#hierarchies)|Les hiérarchies Pivot de tableau croisé dynamique.|
||[rowHierarchies](/javascript/api/excel/excel.pivottabledata#rowhierarchies)|Les hiérarchies de lignes de tableau croisé dynamique.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[disposition](/javascript/api/excel/excel.pivottableloadoptions#layout)|Le PivotLayout décrivant la disposition et la structure visuelle de tableau croisé dynamique.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Renvoie un objet de validation des données.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[dataValidation](/javascript/api/excel/excel.rangedata#datavalidation)|Renvoie un objet de validation des données.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[dataValidation](/javascript/api/excel/excel.rangeloadoptions#datavalidation)|Renvoie un objet de validation des données.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeupdatedata#datavalidation)|Renvoie un objet de validation des données.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Nom de la RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Position de la RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Renvoie les PivotFields associés à la RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID de la RowColumnPivotHierarchy.|
||[Set (propriétés: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties-)|Définit plusieurs propriétés de l’objet en même temps, en fonction d’un objet chargé existant.|
||[Set (propriétés: interfaces. RowColumnPivotHierarchyUpdateData, Options?: objet officeextension. UpdateOptions)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties--options-)|Définit plusieurs propriétés d’un objet en même temps. Vous pouvez transmettre un objet plain avec les propriétés appropriées, ou un autre objet API du même type.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Restaurer la RowColumnPivotHierarchy à ses valeurs par défaut.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Ajouter (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Ajoute le PivotHierarchy à l’axe en cours. Si la hiérarchie est présente ailleurs sur la ligne, colonne,|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Obtient une RowColumnPivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Obtient une RowColumnPivotHierarchy par nom. Si la RowColumnPivotHierarchy n’existe pas, cela renvoie un objet null.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[Remove (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Supprime le PivotHierarchy de l’axe en cours.|
|[RowColumnPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#id)|Pour chaque élément de la collection: ID du RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#name)|Pour chaque élément de la collection: nom du RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#position)|Pour chaque élément de la collection: position du RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyData](/javascript/api/excel/excel.rowcolumnpivothierarchydata)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchydata#fields)|Renvoie les PivotFields associés à la RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchydata#id)|ID de la RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchydata#name)|Nom de la RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchydata#position)|Position de la RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#id)|ID de la RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#name)|Nom de la RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#position)|Position de la RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyUpdateData](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#name)|Nom de la RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#position)|Position de la RowColumnPivotHierarchy.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Activer/désactiver les événements JavaScript dans le volet Office actuel ou le complément de contenu.|
|[RuntimeData](/javascript/api/excel/excel.runtimedata)|[enableEvents](/javascript/api/excel/excel.runtimedata#enableevents)|Activer/désactiver les événements JavaScript dans le volet Office actuel ou le complément de contenu.|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[enableEvents](/javascript/api/excel/excel.runtimeloadoptions#enableevents)|Activer/désactiver les événements JavaScript dans le volet Office actuel ou le complément de contenu.|
|[RuntimeUpdateData](/javascript/api/excel/excel.runtimeupdatedata)|[enableEvents](/javascript/api/excel/excel.runtimeupdatedata#enableevents)|Activer/désactiver les événements JavaScript dans le volet Office actuel ou le complément de contenu.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[BaseField,](/javascript/api/excel/excel.showasrule#basefield)|La base PivotField pour le calcul ShowAs, le cas échéant en fonction du type ShowAsCalculation, sinon null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|La base Item pour le calcul ShowAs, le cas échéant en fonction du type ShowAsCalculation, sinon null.|
||[opération](/javascript/api/excel/excel.showasrule#calculation)|Le calcul ShowAs à utiliser pour le champ de données PivotField. Pour plus d’informations, voir Excel. du showascalculation.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|L’orientation du texte pour le style.|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[autoIndent](/javascript/api/excel/excel.stylecollectionloadoptions#autoindent)|Pour chaque élément de la collection: indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur une distribution égale.|
||[textOrientation](/javascript/api/excel/excel.stylecollectionloadoptions#textorientation)|Pour chaque élément de la collection: orientation du texte.|
|[StyleData](/javascript/api/excel/excel.styledata)|[autoIndent](/javascript/api/excel/excel.styledata#autoindent)|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|
||[textOrientation](/javascript/api/excel/excel.styledata#textorientation)|L’orientation du texte pour le style.|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[autoIndent](/javascript/api/excel/excel.styleloadoptions#autoindent)|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|
||[textOrientation](/javascript/api/excel/excel.styleloadoptions#textorientation)|L’orientation du texte pour le style.|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[autoIndent](/javascript/api/excel/excel.styleupdatedata#autoindent)|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|
||[textOrientation](/javascript/api/excel/excel.styleupdatedata#textorientation)|L’orientation du texte pour le style.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Si Automatic est défini sur true, toutes les autres valeurs seront ignorées lorsque vous configurez les sous-totaux.|
||[temps](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[Max](/javascript/api/excel/excel.subtotals#max)||
||[mn](/javascript/api/excel/excel.subtotals#min)||
||[techniques](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[synthèse](/javascript/api/excel/excel.subtotals#sum)||
||[variante](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Renvoie un ID numérique.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Obtient la plage qui représente la zone modifiée d’un tableau dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Obtient la plage qui représente la zone modifiée d’un tableau dans une feuille de calcul spécifique. Il peut renvoyer un objet null.|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[legacyId](/javascript/api/excel/excel.tablecollectionloadoptions#legacyid)|Pour chaque élément de la collection: renvoie un ID numérique.|
|[TableData](/javascript/api/excel/excel.tabledata)|[legacyId](/javascript/api/excel/excel.tabledata#legacyid)|Renvoie un ID numérique.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[legacyId](/javascript/api/excel/excel.tableloadoptions#legacyid)|Renvoie un ID numérique.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True si le classeur est ouvert en mode lecture seule. En lecture seule.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[WorkbookData](/javascript/api/excel/excel.workbookdata)|[readOnly](/javascript/api/excel/excel.workbookdata#readonly)|True si le classeur est ouvert en mode lecture seule. En lecture seule.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[readOnly](/javascript/api/excel/excel.workbookloadoptions#readonly)|True si le classeur est ouvert en mode lecture seule. En lecture seule.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Se produit lors du calcul de la feuille de calcul.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui est calculée.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique. Il peut renvoyer un objet null.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est calculée.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[showGridlines](/javascript/api/excel/excel.worksheetcollectionloadoptions#showgridlines)|Pour chaque élément de la collection: Obtient ou définit l’indicateur de quadrillage de la feuille de calcul.|
||[showHeadings](/javascript/api/excel/excel.worksheetcollectionloadoptions#showheadings)|Pour chaque élément de la collection: Obtient ou définit l’indicateur des titres de la feuille de calcul.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[showGridlines](/javascript/api/excel/excel.worksheetdata#showgridlines)|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul.|
||[showHeadings](/javascript/api/excel/excel.worksheetdata#showheadings)|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[showGridlines](/javascript/api/excel/excel.worksheetloadoptions#showgridlines)|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul.|
||[showHeadings](/javascript/api/excel/excel.worksheetloadoptions#showheadings)|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[showGridlines](/javascript/api/excel/excel.worksheetupdatedata#showgridlines)|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul.|
||[showHeadings](/javascript/api/excel/excel.worksheetupdatedata#showheadings)|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul.|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
