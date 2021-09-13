---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.8
description: Détails sur l’ensemble de conditions requises ExcelApi 1.8.
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: e97dd98d024b27aa58ca6f0c76fdee17b657c7c9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153687"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Nouveautés de l Excel API JavaScript 1.8

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

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.8. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.8 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Spécifie l’opérande de droite lorsque la propriété de l’opérateur est définie sur un opérateur binaire tel que GreaterThan (l’opérande gauche est la valeur que l’utilisateur tente d’entrer dans la cellule).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Avec les opérateurs ternaires Between et NotBetween, spécifie l’opérande lié supérieur.|
||[opérateur](/javascript/api/excel/excel.basicdatavalidation#operator)|L’opérateur à utiliser pour la validation des données.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categoryLabelLevel)|Spécifie une constante d’éumération au niveau des étiquettes de catégorie de graphique, faisant référence au niveau des étiquettes de catégorie source.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayBlanksAs)|Spécifie la façon dont les cellules vides sont tracées sur un graphique.|
||[plotBy](/javascript/api/excel/excel.chart#plotBy)|Spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotVisibleOnly)|Vrai si seules les cellules visibles sont tracées.|
||[onActivated](/javascript/api/excel/excel.chart#onActivated)|Se produit lorsque le graphique est activé.|
||[onDeactivated](/javascript/api/excel/excel.chart#onDeactivated)|Se produit lorsque le graphique est désactivé.|
||[plotArea](/javascript/api/excel/excel.chart#plotArea)|Représente la zone de traçage du graphique.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesNameLevel)|Spécifie une constante d’éumération de niveau de nom de série de graphique, faisant référence au niveau des noms de série source.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showDataLabelsOverMaximum)|Spécifie s’il faut afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale sur l’axe des valeurs.|
||[style](/javascript/api/excel/excel.chart#style)|Spécifie le style de graphique pour le graphique.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#chartId)|Obtient l’ID du graphique activé.|
||[type](/javascript/api/excel/excel.chartactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le graphique est activé.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#chartId)|Obtient l’ID du graphique qui est ajouté à la feuille de calcul.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.chartaddedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le graphique est ajouté.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignement](/javascript/api/excel/excel.chartaxis#alignment)|Spécifie l’alignement de l’étiquette de la coche de l’axe spécifiée.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isBetweenCategories)|Spécifie si l’axe des valeurs croise l’axe des catégories entre les catégories.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#multiLevel)|Spécifie si un axe est à plusieurs niveaux.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberFormat)|Spécifie le code de format de l’étiquette de la coche de l’axe.|
||[offset](/javascript/api/excel/excel.chartaxis#offset)|Spécifie la distance entre les niveaux d’étiquettes et la distance entre le premier niveau et la ligne d’axe.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Spécifie la position de l’axe spécifié à l’endroit où l’autre axe croise.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionAt)|Spécifie la position de l’axe où l’autre axe croise.|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#setPositionAt_value_)|Définit la position de l’axe spécifié à l’endroit où l’autre axe le croise.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textOrientation)|Spécifie l’angle vers lequel le texte est orienté pour l’étiquette de la tick de l’axe du graphique.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Spécifie la mise en forme du remplissage du graphique.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#setFormula_formula_)|Valeur de chaîne qui représente la formule de titre de l’axe graphique à l’aide de la notation de style A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[bordure](/javascript/api/excel/excel.chartaxistitleformat#border)|Spécifie le format de bordure du titre de l’axe du graphique, qui inclut la couleur, le style de trait et l’pondération.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Spécifie la mise en forme du remplissage du titre de l’axe du graphique.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear__)|Désactiver le format de bordure d’un élément de graphique.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onActivated)|Se produit lorsqu’un graphique est activé.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onAdded)|Se produit lorsqu’un nouveau graphique est ajouté à la feuille de calcul.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#onDeactivated)|Se produit lorsqu’un graphique est désactivé.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#onDeleted)|Se produit lorsqu’un graphique est supprimé.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#autoText)|Spécifie si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalAlignment)|Représente l’alignement horizontal de l’étiquette de données du graphique.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberFormat)|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Représente le format d’étiquette de données graphique.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Représente la hauteur, exprimée en points, de l’étiquette de données du graphique.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Représente la largeur, exprimée en points, de l’étiquette de données du graphique.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Chaîne représentant le texte d’étiquette de données dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textOrientation)|Représente l’angle vers lequel le texte est orienté pour l’étiquette de données du graphique.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalAlignment)|Représente l’alignement vertical de l’étiquette de données du graphique.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[bordure](/javascript/api/excel/excel.chartdatalabelformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#autoText)|Spécifie si les étiquettes de données génèrent automatiquement le texte approprié en fonction du contexte.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalAlignment)|Spécifie l’alignement horizontal pour l’étiquette de données du graphique.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberFormat)|Spécifie le code de format pour les étiquettes de données.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textOrientation)|Représente l’angle vers lequel le texte est orienté pour les étiquettes de données.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalAlignment)|Représente l’alignement vertical de l’étiquette de données du graphique.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#chartId)|Obtient l’ID du graphique qui est désactivé.|
||[type](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le graphique est désactivé.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#chartId)|Obtient l’ID du graphique qui est supprimé de la feuille de calcul.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.chartdeletedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le graphique est supprimé.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Spécifie la hauteur de l’entrée de légende sur la légende du graphique.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Spécifie l’index de l’entrée de légende dans la légende du graphique.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Spécifie la valeur gauche d’une entrée de légende de graphique.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Spécifie le haut d’une entrée de légende de graphique.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Représente la largeur de l’entrée de légende sur la légende du graphique.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[bordure](/javascript/api/excel/excel.chartlegendformat#border)|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Spécifie la valeur de hauteur d’une zone de traçage.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideHeight)|Spécifie la valeur de hauteur intérieure d’une zone de traçage.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideLeft)|Spécifie la valeur à l’intérieur gauche d’une zone de traçage.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insideTop)|Spécifie la valeur supérieure intérieure d’une zone de traçage.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insideWidth)|Spécifie la valeur de largeur intérieure d’une zone de traçage.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Spécifie la valeur gauche d’une zone de traçage.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Spécifie la position d’une zone de traçage.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Spécifie la mise en forme d’une zone de traçage de graphique.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Spécifie la valeur supérieure d’une zone de traçage.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Spécifie la valeur de largeur d’une zone de traçage.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[bordure](/javascript/api/excel/excel.chartplotareaformat#border)|Spécifie les attributs de bordure d’une zone de traçage de graphique.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Spécifie le format de remplissage d’un objet, qui inclut des informations de mise en forme d’arrière-plan.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisGroup)|Spécifie le groupe de la série spécifiée.|
||[explosion](/javascript/api/excel/excel.chartseries#explosion)|Spécifie la valeur d’explosion d’un graphique en secteurs ou d’une tranche de graphique en doughnuts.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstSliceAngle)|Spécifie l’angle de la première tranche de graphique en secteurs ou de graphique en doughnuts, en degrés (dans le sens des aiguilles d’une montre à partir de la verticale).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertIfNegative)|True si Excel inverse le motif dans l’élément lorsqu’il correspond à un nombre négatif.|
||[overlap](/javascript/api/excel/excel.chartseries#overlap)|Spécifie comment barres et colonnes sont positionnées.|
||[dataLabels](/javascript/api/excel/excel.chartseries#dataLabels)|Représente une collection de toutes les étiquettes de données de la série.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondPlotSize)|Spécifie la taille de la section secondaire d’un graphique en secteurs de secteur ou d’un graphique en barres de secteur, sous forme de pourcentage de la taille du secteur principal.|
||[splitType](/javascript/api/excel/excel.chartseries#splitType)|Spécifie le mode de fractionnement des deux sections d’un graphique en secteurs de secteur ou d’un graphique en barres de secteur.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varyByCategories)|True si Excel affecte une couleur ou un motif différent à chaque marqueur de données.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardPeriod)|Représente le nombre de points que la courbe de tendance étend en arrière.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardPeriod)|Représente le nombre de points que la courbe de tendance étend en avant.|
||[label](/javascript/api/excel/excel.charttrendline#label)|Représente l’étiquette d’une courbe de tendance de graphique.|
||[showEquation](/javascript/api/excel/excel.charttrendline#showEquation)|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showRSquared)|True si la valeur r-squared de la courbe de tendance est affichée sur le graphique.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#autoText)|Spécifie si l’étiquette de courbe de tendance génère automatiquement le texte approprié en fonction du contexte.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Valeur de chaîne qui représente la formule de l’étiquette de courbe de tendance du graphique à l’aide de la notation de style A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalAlignment)|Représente l’alignement horizontal de l’étiquette de tendances du graphique.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Représente la distance, en points, entre le bord gauche de l’étiquette de la courbe de tendance du graphique et le bord gauche de la zone de graphique.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberFormat)|Valeur de chaîne qui représente le code de format de l’étiquette de courbe de tendance.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Format de l’étiquette de courbe de tendance du graphique.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Représente la hauteur, exprimée en points, de l’étiquette de tendances du graphique.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Représente la largeur, exprimée en points, de l’étiquette de tendances du graphique.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Chaîne représentant le texte d’étiquette de tendances dans un graphique.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textOrientation)|Représente l’angle vers lequel le texte est orienté pour l’étiquette de courbe de tendance du graphique.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Représente la distance, en points, entre le bord supérieur de l’étiquette de tendances du graphique et le haut de la zone de graphique.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalAlignment)|Représente l’alignement vertical de l’étiquette de tendances du graphique.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[bordure](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Spécifie le format de bordure, qui inclut la couleur, le style de trait et l’pondération.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Spécifie le format de remplissage de l’étiquette de tendances du graphique actuel.|
||[police](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Spécifie les attributs de police (tels que le nom de la police, la taille et la couleur) d’une étiquette de tendances de graphique.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Une formule de validation des données personnalisée.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Nom de la DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberFormat)|Format de nombre de la DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Position de la DataPivotHierarchy.|
||[champ](/javascript/api/excel/excel.datapivothierarchy#field)|Renvoie les PivotFields associés à la DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID de la DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#setToDefault__)|Restaurer la DataPivotHierarchy à ses valeurs par défaut.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showAs)|Spécifie si les données doivent être affichées en tant que calcul récapitulatif spécifique.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeBy)|Spécifie si tous les éléments de la DataPivotHierarchy sont affichés.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add_pivotHierarchy_)|Ajoute le PivotHierarchy à l’axe en cours.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getCount__)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getItem_name_)|Obtient une DataPivotHierarchy par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getItemOrNullObject_name_)|Obtient une DataPivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[remove(DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove_DataPivotHierarchy_)|Supprime le PivotHierarchy de l’axe en cours.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear__)|Efface la validation des données de la plage active.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#errorAlert)|Alerte d’erreur lorsque l’utilisateur entre des données non valides.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreBlanks)|Spécifie si la validation des données sera effectuée sur des cellules vides.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Invite lorsque les utilisateurs sélectionnent une cellule.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Type de validation des données, voir `Excel.DataValidationType` pour plus d’informations.|
||[valide](/javascript/api/excel/excel.datavalidation#valid)|Représente si toutes les valeurs de cellule sont valides selon les règles de validation des données.|
||[rule](/javascript/api/excel/excel.datavalidation#rule)|Règle de validation des données qui contient différents types de critères de validation des données.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Représente le message d’alerte d’erreur.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#showAlert)|Spécifie s’il faut afficher une boîte de dialogue d’alerte d’erreur lorsqu’un utilisateur entre des données non valides.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Type d’alerte de validation des données. Pour `Excel.DataValidationAlertStyle` plus d’informations, voir.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Représente le titre de la boîte de dialogue d’alerte d’erreur.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Spécifie le message de l’invite.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#showPrompt)|Spécifie si une invite s’affiche lorsqu’un utilisateur sélectionne une cellule avec validation des données.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Spécifie le titre de l’invite.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#custom)|Critères de validation des données personnalisés.|
||[date](/javascript/api/excel/excel.datavalidationrule#date)|Critères de validation des données de date.|
||[décimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Critères de validation des données décimales.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Critères de validation des données de liste.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textLength)|Critères de validation des données de longueur du texte.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Critères de validation des données de temps.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholeNumber)|Critères de validation des données de nombre entier.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Spécifie l’opérande de droite lorsque la propriété de l’opérateur est définie sur un opérateur binaire tel que GreaterThan (l’opérande gauche est la valeur que l’utilisateur tente d’entrer dans la cellule).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Avec les opérateurs ternaires Between et NotBetween, spécifie l’opérande lié supérieur.|
||[opérateur](/javascript/api/excel/excel.datetimedatavalidation#operator)|L’opérateur à utiliser pour la validation des données.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enableMultipleFilterItems)|Détermine si vous voulez autoriser plusieurs éléments de filtre.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Nom du filterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Position du filterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Renvoie les PivotFields associés à la FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID de la FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#setToDefault__)|Restaurer la FilterPivotHierarchy à ses valeurs par défaut.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add_pivotHierarchy_)|Ajoute le PivotHierarchy à l’axe en cours.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getCount__)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getItem_name_)|Obtient une FilterPivotHierarchy par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getItemOrNullObject_name_)|Obtient un FilterPivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[remove(filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove_filterPivotHierarchy_)|Supprime le PivotHierarchy de l’axe en cours.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#inCellDropDown)|Spécifie s’il faut afficher la liste dans une liste de cellules.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Source de la liste de validation des données|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Nom du champ PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID du champ de tableau croisé dynamique.|
||[items](/javascript/api/excel/excel.pivotfield#items)|Renvoie les PivotFields associés à PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showAllItems)|Détermine si vous voulez afficher tous les éléments de PivotField.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortByLabels_sortBy_)|Trie le PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Sous-totaux du champ PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getCount__)|Obtient le nombre de champs de tableau croisé dynamique dans la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getItem_name_)|Obtient un champ de tableau croisé dynamique par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getItemOrNullObject_name_)|Obtient un champ de tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Nom de la PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Renvoie les PivotFields associés à la PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID de la PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getCount__)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getItem_name_)|Obtient une PivotHierarchy par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getItemOrNullObject_name_)|Obtient une PivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isExpanded)|Détermine si l’élément est développé pour afficher les éléments enfants ou si ce dernier est réduit et les éléments enfants sont masqués.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Nom du champ PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID du PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Spécifie si l’pivotItem est visible.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getCount__)|Obtient le nombre d’pivotItems dans la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getItem_name_)|Obtient un PivotItem par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getItemOrNullObject_name_)|Obtient un pivotItem par nom.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getColumnLabelRange__)|Renvoie la plage où les étiquettes de colonnes de tableau croisé dynamique se trouvent.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getDataBodyRange__)|Renvoie la plage où les valeurs de données de tableau croisé dynamique se trouvent.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getFilterAxisRange__)|Renvoie la plage de la zone de filtre de tableau croisé dynamique.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getRange__)|Renvoie la plage sur laquelle le tableau croisé dynamique existe, à l’exception de la zone de filtre.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getRowLabelRange__)|Renvoie la plage où les étiquettes de lignes de tableau croisé dynamique se trouvent.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#layoutType)|Cette propriété indique le PivotLayoutType de tous les champs sur le tableau croisé dynamique.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showColumnGrandTotals)|Spécifie si le rapport de tableau croisé dynamique affiche les totaux grands des colonnes.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showRowGrandTotals)|Spécifie si le rapport de tableau croisé dynamique affiche les totaux complets des lignes.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotalLocation)|Cette propriété indique tous les `SubtotalLocationType` champs du tableau croisé dynamique.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete__)|Supprime le tableau croisé dynamique.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnHierarchies)|Les hiérarchies de colonne de tableau croisé dynamique.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#dataHierarchies)|Les hiérarchies de données de tableau croisé dynamique.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterHierarchies)|Les hiérarchies de filtre de tableau croisé dynamique.|
||[hiérarchies](/javascript/api/excel/excel.pivottable#hierarchies)|Les hiérarchies Pivot de tableau croisé dynamique.|
||[disposition](/javascript/api/excel/excel.pivottable#layout)|Le PivotLayout décrivant la disposition et la structure visuelle de tableau croisé dynamique.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#rowHierarchies)|Les hiérarchies de lignes de tableau croisé dynamique.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#add_name__source__destination_)|Ajoutez un tableau croisé dynamique basé sur les données sources spécifiées et insérez-le dans la cellule supérieure gauche de la plage de destination.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#dataValidation)|Renvoie un objet de validation des données.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Nom de la RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Position de la RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Renvoie les PivotFields associés à la RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID de la RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#setToDefault__)|Restaurer la RowColumnPivotHierarchy à ses valeurs par défaut.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add_pivotHierarchy_)|Ajoute le PivotHierarchy à l’axe en cours.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getCount__)|Obtient le nombre de hiérarchies croisées de la collection.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItem_name_)|Obtient une RowColumnPivotHierarchy par son nom ou son ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getItemOrNullObject_name_)|Obtient une RowColumnPivotHierarchy par nom.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[remove(rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove_rowColumnPivotHierarchy_)|Supprime le PivotHierarchy de l’axe en cours.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableEvents)|Basculez les événements JavaScript dans le volet Des tâches ou le module de contenu actuel.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#baseField)|PivotField sur qui baser le calcul, le cas échéant `ShowAs` en fonction du `ShowAsCalculation` type, sinon `null` .|
||[baseItem](/javascript/api/excel/excel.showasrule#baseItem)|Élément sur quoi baser le calcul, le cas échéant en `ShowAs` fonction du `ShowAsCalculation` type, sinon `null` .|
||[calcul](/javascript/api/excel/excel.showasrule#calculation)|Calcul `ShowAs` à utiliser pour le champ de tableau croisé dynamique.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoIndent)|Spécifie si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est égal à la distribution.|
||[textOrientation](/javascript/api/excel/excel.style#textOrientation)|L’orientation du texte pour le style.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Si `Automatic` la valeur est définie sur , toutes les autres `true` valeurs seront ignorées lors de la définition de `Subtotals` la propriété .|
||[moyenne](/javascript/api/excel/excel.subtotals#average)||
||[count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countNumbers)||
||[max](/javascript/api/excel/excel.subtotals#max)||
||[min](/javascript/api/excel/excel.subtotals#min)||
||[produit](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standardDeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standardDeviationP)||
||[sum](/javascript/api/excel/excel.subtotals#sum)||
||[variance](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#varianceP)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyId)|Renvoie un ID numérique.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRange_ctx_)|Obtient la plage qui représente la zone modifiée d’un tableau dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getRangeOrNullObject_ctx_)|Obtient la plage qui représente la zone modifiée d’un tableau dans une feuille de calcul spécifique.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readOnly)|Renvoie `true` si le workbook est ouvert en mode lecture seule.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#onCalculated)|Se produit lorsque la feuille de calcul est calculée.|
||[showGridlines](/javascript/api/excel/excel.worksheet#showGridlines)|Spécifie si le quadrillage est visible par l’utilisateur.|
||[showHeadings](/javascript/api/excel/excel.worksheet#showHeadings)|Spécifie si les titres sont visibles pour l’utilisateur.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[type](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le calcul s’est produit.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRange_ctx_)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getRangeOrNullObject_ctx_)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#onCalculated)|Se produit lorsqu’une feuille de calcul du manuel est calculée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
