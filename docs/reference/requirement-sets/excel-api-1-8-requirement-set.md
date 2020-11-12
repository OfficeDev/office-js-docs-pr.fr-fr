---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,8
description: Détails sur l’ensemble de conditions requises ExcelApi 1,8.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6454a7429276148e36431bfaffdf929a19a36d76
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996206"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Nouveautés de l’API JavaScript pour Excel 1,8

L’ensemble de conditions requises Excel JavaScript API 1.8 incluent des API pour les tableaux croisés dynamiques, validation des données, graphiques, les événements pour les diagrammes, les options de performances et création de classeur.

## <a name="pivottable"></a>Tableau croisé dynamique

Vague 2 des APIs de tableau croisé dynamique permet aux compléments de définir les hiérarchies d’un tableau croisé dynamique. Vous pouvez désormais contrôler les données et comment elles sont regroupées. Notre [Article tableau croisé dynamique](../../excel/excel-add-ins-pivottables.md) a plus d’informations sur les nouvelles fonctionnalités de tableau croisé dynamique.

## <a name="data-validation"></a>Validation des données

La validation des données vous donne le contrôle sur ce qu’un utilisateur insère dans une feuille de calcul. Vous pouvez limiter les cellules à des ensembles de réponse prédéfinie ou donner des avertissements contextuels concernant des entrées indésirables. En savoir plus maintenant sur [Ajout de validation des données à des plages](../../excel/excel-add-ins-data-validation.md).

## <a name="charts"></a>Graphiques

Une autre série de graphiques API apporte un meilleur contrôle par programme des éléments de graphique. Vous avez à présent un meilleur accès à la légende, axes, courbe de tendance et zone de traçage.

## <a name="events"></a>Événements

Plus d’[événements](../../excel/excel-add-ins-events.md) ont été ajoutés pour les graphiques. Votre complément réagit aux interactions des utilisateurs avec le graphique. Vous pouvez également [Activer ou désactiver les événements](../../excel/performance.md#enable-and-disable-events) sur l’ensemble du classeur.

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,8. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,8 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,8 ou version antérieure](/javascript/api/excel?view=excel-js-1.8&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Spécifie l’opérande de droite lorsque la propriété Operator est définie sur un opérateur binaire tel que GreaterThan (l’opérande de gauche est la valeur que l’utilisateur tente d’entrer dans la cellule).|
||[Formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Avec les opérateurs ternaires entre et NotBetween, spécifie l’opérande supérieur.|
||[opérateur](/javascript/api/excel/excel.basicdatavalidation#operator)|L’opérateur à utiliser pour la validation des données.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Spécifie une constante d’énumération ChartCategoryLabelLevel faisant référence à|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Indique la façon dont les cellules vides sont tracées dans un graphique.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique.|
||[PlotVisibleOnly,](/javascript/api/excel/excel.chart#plotvisibleonly)|Vrai si seules les cellules visibles sont tracées. Faux si les deux cellules visibles et masquées sont tracées.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Se produit lorsque le graphique est activé.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Se produit lorsque le graphique est désactivé.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Représente la zone de traçage pour le graphique.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Spécifie une constante d’énumération ChartSeriesNameLevel faisant référence à|
||[ShowDataLabelsOverMaximum,](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Indique s’il faut afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale de l’axe des ordonnées.|
||[style](/javascript/api/excel/excel.chart#style)|Cette énumération spécifie le style de graphique du graphique.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Obtient l’id du graphique qui est activé.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est activé.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Obtient l’id du graphique qui est ajouté à la feuille de calcul.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est ajouté.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[aligne](/javascript/api/excel/excel.chartaxis#alignment)|Spécifie l’alignement de l’étiquette de graduation de l’axe spécifié.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Indique si l’axe des ordonnées coupe l’axe des abscisses entre les catégories.|
||[Niveaux](/javascript/api/excel/excel.chartaxis#multilevel)|Indique si un axe est à plusieurs niveaux.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Spécifie le code du format de l’étiquette de graduation de l’axe.|
||[compensé](/javascript/api/excel/excel.chartaxis#offset)|Indique la distance entre les niveaux des étiquettes et la distance entre le premier niveau et la ligne de l’axe.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Spécifie la position de l’axe spécifié où l’autre axe croise.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Spécifie la position de l’axe spécifié où l’autre axe se coupe.|
||[setPositionAt (valeur : nombre)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Définit la position de l’axe spécifié où l’autre axe se coupe.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Cette énumération spécifie l’angle auquel le texte est orienté pour l’étiquette de graduation de l’axe du graphique.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Cette énumération spécifie la mise en forme du remplissage du graphique.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula (Formula : String)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Valeur de chaîne qui représente la formule de titre de l’axe graphique à l’aide de la notation de style A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[bordure](/javascript/api/excel/excel.chartaxistitleformat#border)|Indique le format de bordure du titre de l’axe du graphique, qui inclut la couleur, le style de style et le poids.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Cette énumération spécifie la mise en forme du titre de l’axe du graphique.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Désactiver le format de bordure d’un élément de graphique.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Se produit lorsqu’un graphique est activé.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Survient lors de l’ajout d’un nouveau graphique à la feuille de calcul.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Se produit lorsqu’un graphique est désactivé.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Survient lors de la suppression d’un graphique.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabel#autotext)|Spécifie si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de données du graphique.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Représente le format d’étiquette de données graphique.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Représente la hauteur, exprimée en points, de l’étiquette de données du graphique.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Représente la largeur, exprimée en points, de l’étiquette de données du graphique.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Chaîne représentant le texte d’étiquette de données dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Représente l’angle auquel le texte est orienté pour l’étiquette de données du graphique.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[bordure](/javascript/api/excel/excel.chartdatalabelformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[Insertion automatique](/javascript/api/excel/excel.chartdatalabels#autotext)|Indique si les étiquettes de données génèrent automatiquement le texte approprié en fonction du contexte.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Indique l’alignement horizontal de l’étiquette de données du graphique.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Spécifie le code de format des étiquettes de données.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Représente l’angle de l’orientation du texte pour les étiquettes de données.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Représente l’alignement vertical de l’étiquette de données du graphique.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Obtient l’id du graphique qui est desactivé.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est desactivé.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Obtient l’id du graphique qui est supprimé de la feuille de calcul.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est supprimé.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Cette énumération spécifie la hauteur de la fonction legendEntry sur la légende du graphique.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Spécifie l’index de legendEntry dans la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Spécifie la partie gauche d’un graphique legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Spécifie le bord supérieur d’un graphique legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Représente la largeur de legendEntry sur la légende d’un graphique.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[bordure](/javascript/api/excel/excel.chartlegendformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Spécifie la valeur de la hauteur de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Spécifie la valeur insideHeight de plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Spécifie la valeur insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Spécifie la valeur insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Spécifie la valeur insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Spécifie la valeur gauche de plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Spécifie la position de plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Cette énumération spécifie la mise en forme d’un graphique plotArea.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Spécifie la valeur la plus haute de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Spécifie la valeur de la largeur de plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[bordure](/javascript/api/excel/excel.chartplotareaformat#border)|Cette énumération spécifie les attributs de bordure d’un graphique plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Spécifie le format de remplissage d’un objet, qui inclut les informations de mise en forme de l’arrière-plan.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Cette énumération spécifie le groupe de la série spécifiée.|
||[barrage](/javascript/api/excel/excel.chartseries#explosion)|Indique la valeur d’explosion d’un graphique en secteurs ou d’un graphique en anneau.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Spécifie l’angle du premier secteur de graphique en secteurs ou en anneau, en degrés (dans le sens des aiguilles d’une montre verticale).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|La valeur true si Excel inverse le motif dans l’élément lorsqu’il correspond à un nombre négatif.|
||[coin](/javascript/api/excel/excel.chartseries#overlap)|Spécifie comment barres et colonnes sont positionnées.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Représente la collection de tous les dataLabels de la série.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Indique la taille de la section secondaire d’un graphique en secteurs de secteur ou un graphique en barres de secteur, sous la forme d’un pourcentage de la taille du secteur principal.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Indique la façon dont les deux sections d’un graphique en secteurs de secteur ou un graphique en barres de secteur sont fractionnées.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|La valeur true si Excel affecte une couleur ou un motif différent à chaque marqueur de données.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Représente le nombre de points que la courbe de tendance étend en arrière.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Représente le nombre de points que la courbe de tendance étend en avant.|
||[labellisé](/javascript/api/excel/excel.charttrendline#label)|Représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showequation)|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[Insertion automatique](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Indique si l’étiquette de courbe de tendance génère automatiquement le texte approprié en fonction du contexte.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Valeur de chaîne qui représente la formule de l’étiquette de tendances du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Représente l’alignement horizontal de l’étiquette de tendances du graphique.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Représente la distance en points, du bord gauche de l’étiquette de tendances graphique au bord gauche de la zone de graphique.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Valeur de chaîne qui représente le code de format pour l’étiquette de tendances.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Format de l’étiquette de courbe de tendance du graphique.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Représente la hauteur, exprimée en points, de l’étiquette de tendances du graphique.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Représente la largeur, exprimée en points, de l’étiquette de tendances du graphique.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Chaîne représentant le texte d’étiquette de tendances dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Représente l’angle auquel le texte est orienté pour l’étiquette de courbe de tendance du graphique.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Représente la distance en points, du bord supérieur de l’étiquette de tendances du graphique au bord supérieur de la zone de graphique.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Représente l’alignement vertical de l’étiquette de tendances du graphique.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[bordure](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Cette énumération spécifie le format de bordure, qui inclut couleur, LineStyle et Weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Cette énumération spécifie le format de remplissage de l’étiquette de courbe de tendance du graphique actif.|
||[police](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Spécifie les attributs de police (nom de police, taille de police, couleur, etc.) pour une étiquette de courbe de tendance de graphique.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Une formule de validation des données personnalisée.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Nom de la DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Format de nombre de la DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Position de la DataPivotHierarchy.|
||[field](/javascript/api/excel/excel.datapivothierarchy#field)|Renvoie les PivotFields associés à la DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID de la DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Restaurer la DataPivotHierarchy à ses valeurs par défaut.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Indique si les données doivent être affichées sous la forme d’un calcul de synthèse spécifique.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Indique si tous les éléments de l’DataPivotHierarchy sont affichés.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[Ajouter (pivotHierarchy : Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Ajoute le PivotHierarchy à l’axe en cours.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Obtient une DataPivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Obtient une DataPivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[Remove (DataPivotHierarchy : Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Supprime le PivotHierarchy de l’axe en cours.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Efface la validation des données de la plage active.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Alerte d’erreur lorsque l’utilisateur entre des données non valides.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Indique si la validation des données est effectuée sur des cellules vides, la valeur par défaut est true.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Invite lorsque les utilisateurs sélectionnent une cellule.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Type de validation des données, voir Excel.DataValidationType pour plus d’informations.|
||[validation](/javascript/api/excel/excel.datavalidation#valid)|Représente si toutes les valeurs de cellule sont valides selon les règles de validation des données.|
||[sous](/javascript/api/excel/excel.datavalidation#rule)|Règle de validation des données qui contient différents critères de validation de données.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Représente le message d’alerte d’erreur.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Indique si une boîte de dialogue d’alerte d’erreur doit s’afficher lorsqu’un utilisateur entre des données non valides.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Type d’alerte de validation des données, consultez Excel. DataValidationAlertStyle pour plus d’informations.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Représente le titre de dialogue d’alerte d’erreur.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Spécifie le message de l’invite.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Indique si une invite s’affiche lorsqu’un utilisateur sélectionne une cellule avec validation des données.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Spécifie le titre de l’invite.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[personnalisé](/javascript/api/excel/excel.datavalidationrule#custom)|Critères de validation des données personnalisés.|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|Critères de validation des données de date.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Critères de validation des données décimales.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Critères de validation des données de liste.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Critères de validation des données textLength.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Critères de validation des données de temps.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Critères de validation des données WholeNumber.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Spécifie l’opérande de droite lorsque la propriété Operator est définie sur un opérateur binaire tel que GreaterThan (l’opérande de gauche est la valeur que l’utilisateur tente d’entrer dans la cellule).|
||[Formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Avec les opérateurs ternaires entre et NotBetween, spécifie l’opérande supérieur.|
||[opérateur](/javascript/api/excel/excel.datetimedatavalidation#operator)|L’opérateur à utiliser pour la validation des données.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Détermine si vous voulez autoriser plusieurs éléments de filtre.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Nom du filterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Position du filterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Renvoie les PivotFields associés à la FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID du filterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Restaurer la FilterPivotHierarchy à ses valeurs par défaut.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[Ajouter (pivotHierarchy : Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Ajoute le PivotHierarchy à l’axe en cours.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Obtient une FilterPivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Obtient un FilterPivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[Remove (filterPivotHierarchy : Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Supprime le PivotHierarchy de l’axe en cours.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[InCellDropdown,](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Affiche la liste dans la cellule déroulante ou non, sa valeur par défaut est true.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Source de la liste de validation des données|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Nom du champ PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID du champ PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Renvoie les PivotFields associés à PivotField.|
||[ShowAllItems,](/javascript/api/excel/excel.pivotfield#showallitems)|Détermine si vous voulez afficher tous les éléments de PivotField.|
||[sortByLabels (sortBy : SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Trie le PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Sous-totaux du champ PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Obtient le nombre de champs de tableau croisé dynamique dans la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Obtient un champ de tableau croisé dynamique par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Obtient un champ de tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Nom de la PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Renvoie les PivotFields associés à la PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID de la PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Obtient une PivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Obtient une PivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[IsExpanded,](/javascript/api/excel/excel.pivotitem#isexpanded)|Détermine si l’élément est développé pour afficher les éléments enfants ou si ce dernier est réduit et les éléments enfants sont masqués.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Nom du champ PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID du champ PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Indique si la PivotItem est visible.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Obtient le nombre de PivotItems dans la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Obtient un PivotItem par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Obtient un PivotItem par nom.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Renvoie la plage où les étiquettes de colonnes de tableau croisé dynamique se trouvent.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Renvoie la plage où les valeurs de données de tableau croisé dynamique se trouvent.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Renvoie la plage de la zone de filtre de tableau croisé dynamique.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Renvoie la plage sur laquelle le tableau croisé dynamique existe, à l’exception de la zone de filtre.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Renvoie la plage où les étiquettes de lignes de tableau croisé dynamique se trouvent.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Cette propriété indique le PivotLayoutType de tous les champs sur le tableau croisé dynamique.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Indique si le rapport de tableau croisé dynamique affiche les totaux généraux des colonnes.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Indique si le rapport de tableau croisé dynamique affiche des totaux généraux de lignes.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Cette propriété indique le SubtotalLocationType de tous les champs sur le tableau croisé dynamique.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Supprime le tableau croisé dynamique.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|Les hiérarchies de colonne de tableau croisé dynamique.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#datahierarchies)|Les hiérarchies de données de tableau croisé dynamique.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|Les hiérarchies de filtre de tableau croisé dynamique.|
||[hiérarchies](/javascript/api/excel/excel.pivottable#hierarchies)|Les hiérarchies Pivot de tableau croisé dynamique.|
||[disposition](/javascript/api/excel/excel.pivottable#layout)|Le PivotLayout décrivant la disposition et la structure visuelle de tableau croisé dynamique.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowhierarchies)|Les hiérarchies de lignes de tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (Name : chaîne, source : \| \| table de chaînes de plage, destination : chaîne de plage \| )](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Ajoutez un tableau croisé dynamique basé sur les données sources spécifiées et insérez-le dans la cellule située en haut à gauche de la plage de destination.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Renvoie un objet de validation des données.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Nom de la RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Position de la RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Renvoie les PivotFields associés à la RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID de la RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Restaurer la RowColumnPivotHierarchy à ses valeurs par défaut.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Ajouter (pivotHierarchy : Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Ajoute le PivotHierarchy à l’axe en cours.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Obtient une RowColumnPivotHierarchy par son nom ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Obtient une RowColumnPivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[Remove (rowColumnPivotHierarchy : Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Supprime le PivotHierarchy de l’axe en cours.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Activer/désactiver les événements JavaScript dans le volet Office actuel ou le complément de contenu.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[BaseField,](/javascript/api/excel/excel.showasrule#basefield)|La base PivotField pour le calcul ShowAs, le cas échéant en fonction du type ShowAsCalculation, sinon null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|La base Item pour le calcul ShowAs, le cas échéant en fonction du type ShowAsCalculation, sinon null.|
||[opération](/javascript/api/excel/excel.showasrule#calculation)|Le calcul ShowAs à utiliser pour le champ de données PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur une distribution égale.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|L’orientation du texte pour le style.|
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
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Obtient la plage qui représente la zone modifiée d’un tableau dans une feuille de calcul spécifique.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True si le classeur est ouvert en mode lecture seule.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#oncalculated)|Se produit lors du calcul de la feuille de calcul.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showgridlines)|Indique si le quadrillage est visible par l’utilisateur.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showheadings)|Indique si les en-têtes sont visibles par l’utilisateur.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le calcul s’est produit.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Cet événement se produit lorsqu’une feuille de calcul du classeur est calculée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
