---
title: Ensembles de conditions requises de l’API JavaScript pour Excel
description: ''
ms.date: 10/09/2018
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 0cb01bffb2947c4269e377c39ff3c329c48591d9
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742470"
---
# <a name="excel-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Excel

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Les compléments Excel peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 pour Windows, Office pour iPad, Office pour Mac et Office Online. Le tableau suivant répertorie les ensembles de conditions requises pour Excel, les applications hôtes Office qui prennent en charge chaque ensemble de conditions et les versions ou numéro de build de ces applications.

> [!NOTE]
> Toute API qui est répertoriée en tant que version**Bêta** n’est pas prête pour la production pour utilisateur final. Elles sont rendues disponibles afin que les développeurs puissent les tester dans les environnements de test et de développement. Elles ne sont pas destinées à être utilisées avec des documents essentiels de production/d’entreprise.
> 
> Pour les ensembles de conditions marqués comme étant **Bêta**, utilisez la version spécifiée (ou version ultérieure) du logiciel Office et la bibliothèque bêta sur le CDN : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Les entrées non marquées comme **Bêta** sont généralement disponibles et vous pouvez utiliser la bibliothèque Production sur le CDN : https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Ensemble de conditions requises  |  Office 365 pour Windows\*  |  Office 365 pour iPad  |  Office 365 pour Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Bêta  | Veuillez [visiter notre page ouvrir la spécification API JavaScript Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.8  | Version 1808 (build 10730.20102) ou ultérieure | 2.17 ou version ultérieure | 16.17 ou version ultérieure | Septembre 2018 | Bientôt disponible |
| ExcelApi1.7  | Version 1801 (build 9001.2171) ou ultérieure   | 2.9 ou version ultérieure | 16.9 ou version ultérieure | Avril 2018 | Bientôt disponible |
| ExcelApi1.6  | Version 1704 (Build 8201.2001) ou version ultérieure   | 2.2 ou version ultérieure |15.36 ou version ultérieure| Avril 2017 | Bientôt disponible|
| ExcelApi1.5  | Version 1703 (Build 8067.2070) ou version ultérieure   | 2.2 ou version ultérieure |15.36 ou version ultérieure| Mars 2017 | Bientôt disponible|
| ExcelApi1.4  | Version 1701 (Build 7870.2024) ou version ultérieure   | 2.2 ou version ultérieure |15.36 ou version ultérieure| Janvier 2017 | Bientôt disponible|
| ExcelApi1.3  | Version 1608 (Build 7369.2055) ou version ultérieure | 1.27 ou version ultérieure |  15.27 ou version ultérieure| Septembre 2016 | Version 1608 (Build 7601.6800) ou version ultérieure|
| ExcelApi1.2  | Version 1601 (Build 6741.2088) ou version ultérieure | 1.21 ou version ultérieure | 15.22 ou version ultérieure| Janvier 2016 ||
| ExcelApi1.1  | Version 1509 (Build 4266.1001) ou version ultérieure | 1.19 ou version ultérieure | 15.20 ou version ultérieure| Janvier 2016 ||

> [!NOTE]
> Le numéro de build d’Office 2016 installé via MSI est 16.0.4266.1001. Cette version ne contient que l’ensemble de conditions requises de l’ExcelApi 1.1.

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Où trouver le numéro de version et de build pour une application cliente Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Présentation d’Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-18"></a>Nouveautés de l’API JavaScript 1.8 pour Excel

L’ensemble de conditions requises Excel JavaScript API 1.8 incluent des API pour les tableaux croisés dynamiques, validation des données, graphiques, les événements pour les diagrammes, les options de performances et création de classeur.

### <a name="pivottable"></a>Tableau croisé dynamique

Vague 2 des APIs de tableau croisé dynamique permet aux compléments de définir les hiérarchies d’un tableau croisé dynamique. Vous pouvez désormais contrôler les données et comment elles sont regroupées. Notre [Article tableau croisé dynamique](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) a plus d’informations sur les nouvelles fonctionnalités de tableau croisé dynamique.

### <a name="data-validation"></a>Validation des données

La validation des données vous donne le contrôle sur ce qu’un utilisateur insère dans une feuille de calcul. Vous pouvez limiter les cellules à des ensembles de réponse prédéfinie ou donner des avertissements contextuels concernant des entrées indésirables. En savoir plus maintenant sur [Ajout de validation des données à des plages](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation).

### <a name="charts"></a>Graphiques

Une autre série de graphiques API apporte un meilleur contrôle par programme des éléments de graphique. Vous avez à présent un meilleur accès à la légende, axes, courbe de tendance et zone de traçage.

### <a name="events"></a>Événements

Plus d’[événements](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) ont été ajoutés pour les graphiques. Votre complément réagit aux interactions des utilisateurs avec le graphique. Vous pouvez également [Activer ou désactiver les événements](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) sur l’ensemble du classeur.


|Objet| Nouveautés| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Méthode_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|Crée un nouveau classeur masqué à l’aide d’un fichier facultatif .xlsx codé en base 64.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Propriété_ > formula1|Obtient ou définit Formula1, c'est-à-dire la valeur minimale ou la valeur en fonction de l’opérateur.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Propriété_ > formula2|Obtient ou définit Formula2, c'est-à-dire la valeur maximale ou la valeur en fonction de l’opérateur.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Relation_ > opérateur|L’opérateur à utiliser pour la validation des données.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > categoryLabelLevel|Renvoie ou définit une constante énumération ChartCategoryLabelLevel faisant référence au niveau de l’endroit d’où les étiquettes de catégorie proviennent. Lecture/écriture.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > plotVisibleOnly|Vrai si seules les cellules visibles sont tracées. Faux si les deux cellules visibles et masquées sont tracées. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > seriesNameLevel|Renvoie ou définit une constante énumération ChartSeriesNameLevel faisant référence au niveau de l’endroit d’où les noms de séries proviennent. Lecture/écriture.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > showDataLabelsOverMaximum|Si vous voulez afficher les étiquettes de données lorsque la valeur est supérieure à la valeur maximale sur l’axe de valeur.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > style|Cette propriété renvoie ou définit le style de graphique pour le graphique. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relation_ > displayBlanksAs|Renvoie ou définit la façon dont les cellules vides sont tracées sur un graphique. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relation_ > plotArea|Représente la zone de traçage pour le graphique. En lecture seule.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relation_ > plotby|Renvoie ou spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. ReadWrite.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriété_ > chartId|Obtient l’id du graphique qui est activé.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriété_ > type|Obtient le type de l’événement.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle le graphique est activé.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriété_ > chartId|Obtient l’id du graphique qui est ajouté à la feuille de calcul.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriété_ > type|Obtient le type de l’événement.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle le graphique est ajouté.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Relation_ > source|Obtient la source de l’événement.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > isBetweenCategories|Représente si l’axe de valeur croise l’axe de catégorie entre catégories.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > multiLevel|Représente si un axe est à plusieurs niveaux ou non.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > numberFormat|Représente le code de format pour l’étiquette de graduation d’axe.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > offset|Représente la distance entre les niveaux d’étiquettes et la distance entre le premier niveau et la ligne d’axe. La valeur doit être un entier compris entre 0 et 1000.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > positionAt|Représente la position de l’axe spécifié où l’autre axe le croise. Vous devez utiliser la méthode SetPositionAt(double) pour définir cette propriété. En lecture seule.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > textOrientation|Représente l’orientation du texte de l’étiquette de graduation de l’axe. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relation_ > alignment|Représente l’alignement vertical de l’étiquette de la graduation de l’axe spécifié.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relation_ > position|Représente la position de l’axe spécifié où l’autre axe le croise.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Méthode_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|Représente la position de l’axe spécifié où l’autre axe le croise.|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_Relation_ > fill|Représente la mise en forme de remplissage du graphique. En lecture seule.|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_Méthode_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|Valeur de chaîne qui représente la formule de titre de l’axe graphique à l’aide de la notation de style A1.|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relation_ > border|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relation_ > fill|Représente la mise en forme de remplissage du graphique. En lecture seule.|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Méthode_ > [clear()](/javascript/api/excel/excel.chartborder)|Désactiver le format de bordure d’un élément de graphique.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > autoText|Valeur booléenne représentant si l’étiquette de données génère automatiquement le texte approprié en fonction du contexte.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > formula|Valeur de chaîne qui représente la formule de l’étiquette de données du graphique à l’aide de la notation de style A1.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > height|Représente la hauteur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible. En lecture seule.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > left|Représente la distance en points, du bord gauche de l’étiquette de données graphique au bord gauche de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > numberFormat|Valeur de chaîne qui représente le code de format pour l’étiquette de données.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > text|Chaîne représentant le texte d’étiquette de données dans un graphique.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > textOrientation|Représente l’orientation du texte de l’étiquette de données du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > top|Représente la distance en points, du bord supérieur de l’étiquette de données graphique au bord supérieur de la zone de graphique. Null si l’étiquette de données graphique n’est pas visible.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > width|Représente la largeur, exprimée en points, de l’étiquette de données du graphique. En lecture seule. Null si l’étiquette de données graphique n’est pas visible. En lecture seule.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relation_ > format|Représente le format d’étiquette de données graphique. En lecture seule.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relation_ > horizontalAlignment|Représente l’alignement horizontal de l’étiquette de données du graphique.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relation_ > verticalAlignment|Représente l’alignement vertical de l’étiquette de données du graphique.|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_Relation_ > border|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriété_ > autoText|Représente si des étiquettes de données génèrent automatiquement le texte approprié en fonction du contexte.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriété_ > numberFormat|Représente le code de format pour les étiquettes de données.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriété_ > textOrientation|Représente l’orientation du texte des étiquettes de données. La valeur doit être un entier soit de -90 à 90, soit de 0 à 180 pour le texte orienté verticalement.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relation_ > horizontalAlignment|Représente l’alignement horizontal de l’étiquette de données du graphique.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relation_ > verticalAlignment|Représente l’alignement vertical de l’étiquette de données du graphique.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriété_ > chartId|Obtient l’id du graphique qui est desactivé.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriété_ > type|Obtient le type de l’événement.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle le graphique est desactivé.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriété_ > chartId|Obtient l’id du graphique qui est supprimé de la feuille de calcul.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriété_ > type|Obtient le type de l’événement.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle le graphique est supprimé.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Relation_ > source|Obtient la source de l’événement.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriété_ > height|Représente la hauteur de legendEntry sur la légende du graphique. En lecture seule.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriété_ > index|Représente l’index de legendEntry sur la légende du graphique. En lecture seule.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriété_ > left|Représente la partie gauche d’un graphique legendEntry. En lecture seule.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriété_ > top|Représente la partie supérieure d’un graphique legendEntry. En lecture seule.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriété_ > width|Représente la largeur de legendEntry sur la légende d’un graphique. En lecture seule.|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_Relation_ > border|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > height|Représente la valeur de hauteur de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > insideHeight|Représente la valeur insideHeight de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > insideLeft|Représente la valeur insideLeft de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > insideTop|Représente la valeur insideTop de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > insideWidth|Représente la valeur insideWidth de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > left|Représente la valeur gauche de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > top|Représente la valeur supérieure de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriété_ > width|Représente la valeur de largeur de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relation_ > format|Représente la mise en forme d’un graphique plotArea. En lecture seule.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relation_ > position|Représentant la position de plotArea.|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relation_ > border|Représente les attributs de bordure d’un graphique plotArea. En lecture seule.|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relation_ > fill|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan. En lecture seule.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > explosion|Renvoie ou définit la valeur d’explosion pour une coupe de graphique en secteurs ou de graphique en anneaux. Renvoie 0 (zéro) s’il n’y a aucune explosion (la pointe de la coupe est dans le centre du graphique). ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > firstSliceAngle|Renvoie ou définit l’angle de la première coupe graphique en secteurs ou graphique en anneaux, en degrés (dans le sens des aiguilles d’une montre, vertical). S’applique uniquement aux graphiques en secteurs, graphiques en secteurs 3D et graphiques en anneaux. Peut être une valeur comprise entre 0 et 360. ReadWrite|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > invertIfNegative|Vrai si Microsoft Excel inverse le motif dans l’élément lorsqu’il correspond à un nombre négatif. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > overlap|Spécifie comment barres et colonnes sont positionnées. Peut être une valeur comprise entre -100 et 100. S’applique uniquement aux graphiques en barres 2D et en colonnes 2D. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > secondPlotSize|Renvoie ou définit la taille de la section secondaire d’un secteur de graphique en secteurs ou d’une barre de graphique en secteurs, sous forme de pourcentage de la taille du graphique principal. Peut être une valeur comprise entre 5 et 200. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > varyByCategories|Vrai si Microsoft Excel affecte une couleur ou un motif différentes à chaque marqueur de données. Le graphique ne doit contenir qu’une seule série. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relation_ > axisGroup|Renvoie ou définit le groupe pour la série spécifiée. ReadWrite|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relation_ > dataLabels|Représente la collection de tous les dataLabels de la série. En lecture seule.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relation_ > splitType|Renvoie ou définit la façon dont les deux sections d’un secteur de graphique en secteurs ou de barre d’un graphique en secteurs sont réparties. ReadWrite.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > backwardPeriod|Représente le nombre de points que la courbe de tendance étend en arrière.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > forwardPeriod|Représente le nombre de points que la courbe de tendance étend en avant.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > showEquation|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > showRSquared|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relation_ > label|Représente l’étiquette d’une courbe de tendance de graphique. En lecture seule.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > autoText|Valeur booléenne représentant si l’étiquette de tendances génère automatiquement le texte approprié en fonction du contexte.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > formula|Valeur de chaîne qui représente la formule de l’étiquette de tendances du graphique à l’aide de la notation de style A1.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > height|Représente la hauteur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible. En lecture seule.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > left|Représente la distance en points, du bord gauche de l’étiquette de tendances graphique au bord gauche de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > numberFormat|Valeur de chaîne qui représente le code de format pour l’étiquette de tendances.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > text|Chaîne représentant le texte d’étiquette de tendances dans un graphique.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > textOrientation|Représente l’orientation du texte de l’étiquette de tendances du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > top|Représente la distance en points, du bord supérieur de l’étiquette de tendances du graphique au bord supérieur de la zone de graphique. Null si l’étiquette de tendances graphique n’est pas visible.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriété_ > width|Représente la largeur, exprimée en points, de l’étiquette de tendances du graphique. En lecture seule. Null si l’étiquette de tendances graphique n’est pas visible. En lecture seule.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relation_ > format|Représente le format d’étiquette de tendances du graphique. En lecture seule.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relation_ > horizontalAlignment|Représente l’alignement horizontal de l’étiquette de tendances du graphique.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relation_ > verticalAlignment|Représente l’alignement vertical de l’étiquette de tendances du graphique.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relation_ > border|Représente le format bordure, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relation_ > fill|Représente le format de remplissage de l’étiquette de tendances du graphique actuel. En lecture seule.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relation_ > font|Représente les attributs de police (nom de la police, taille de police, couleur, etc.) d’une étiquette de tendances de graphique. En lecture seule.|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_Propriété_ > formula| Une formule de validation des données personnalisée. Cette opération crée des règles d’entrée spéciales, comme empêcher les doublons ou limiter le total dans une plage de cellules.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriété_ > id|ID de la DataPivotHierarchy. En lecture seule.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriété_ > name|Nom de la DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriété_ > numberFormat|Format de nombre de la DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriété_ > position|Position de la DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relation_ > field|Renvoie les PivotFields associés à la DataPivotHierarchy. En lecture seule.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relation_ > showAs|Détermine si les données doivent apparaître sous forme d’un calcul de synthèse spécifique ou non.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relation_ > summarizeBy|Détermine si vous voulez afficher tous les éléments de la DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Méthode_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Restaurer la DataPivotHierarchy à ses valeurs par défaut.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Propriété_ > items|Une collection d’objets dataPivotHierarchy. En lecture seule.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Méthode_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Ajoute le PivotHierarchy à l’axe en cours.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Méthode_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|Obtient le nombre de hiérarchies croisées de la collection.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Méthode_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtient une DataPivotHierarchy par son nom ou id.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Méthode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtient une DataPivotHierarchy par nom. Si la DataPivotHierarchy n’existe pas, cela renvoie un objet null.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Supprime le PivotHierarchy de l’axe en cours.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Propriété_ > ignoreBlanks|Ignorer les espaces vides : aucune validation des données ne sera exécutée sur les cellules vides, la valeur par défaut est vrai.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Propriété_ > valid|Représente si toutes les valeurs de cellule sont valides selon les règles de validation des données. En lecture seule.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relation_ > errorAlert|Alerte d’erreur lorsque l’utilisateur entre des données non valides.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relation_ > prompt|Invite lorsque les utilisateurs sélectionnent une cellule.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relation_ > rule|Règle de validation des données qui contient différents types de critères de validation des données.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relation_ > type|Type de validation des données, voir [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) pour plus d’informations. En lecture seule.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Méthode_ > [clear()](/javascript/api/excel/excel.datavalidation)|Efface la validation des données de la plage active.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriété_ > message|Représente le message d’alerte d’erreur.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriété_ > showAlert|Détermine si vous voulez afficher un dialogue Alerte d’erreur ou pas lorsqu’un utilisateur entre des données non valides. La valeur par défaut est True.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriété_ > title|Représente le titre de dialogue d’alerte d’erreur.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Relation_ > style|Représente un type d’alerte de validation des données, voir [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) pour plus d’informations.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Propriété_ > message|Représente le message de l’invite.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Propriété_ > showPrompt|Détermine d’afficher ou non l’invite lorsqu’un utilisateur sélectionne une cellule avec validation des données.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Propriété_ > title|Représente le titre de l’invite.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relation_ > custom|Critères de validation des données personnalisés.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relation_ > date|Critères de validation des données de date.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relation_ > decimal|Critères de validation des données décimales.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relation_ > list|Critères de validation des données de liste.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relation_ > textLength|Critères de validation des données textLength.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relation_ > time|Critères de validation des données de temps.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relation_ > wholeNumber|Critères de validation des données WholeNumber.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Propriété_ > formula1|Obtient ou définit Formula1, c'est-à-dire la valeur minimale ou la valeur en fonction de l’opérateur.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Propriété_ > formula2|Obtient ou définit Formula2, c'est-à-dire la valeur maximale ou la valeur en fonction de l’opérateur.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Relation_ > opérateur|L’opérateur à utiliser pour la validation des données.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriété_ > enableMultipleFilterItems|Détermine si vous voulez autoriser plusieurs éléments de filtre.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriété_ > id|ID du filterPivotHierarchy. En lecture seule.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriété_ > name|Nom du filterPivotHierarchy.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriété_ > position|Position du filterPivotHierarchy.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Relation_ > fields|Renvoie les PivotFields associés à la FilterPivotHierarchy. En lecture seule.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Méthode_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|Restaurer la FilterPivotHierarchy à ses valeurs par défaut.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Propriété_ > items|Une collection d’objets filterPivotHierarchy. En lecture seule.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Méthode_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Ajoute le PivotHierarchy à l’axe en cours. Si la hiérarchie apparaît ailleurs sur la ligne, colonne ou axe de filtre, celle-ci est supprimée de cet emplacement.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Méthode_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtient le nombre de hiérarchies croisées de la collection.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Méthode_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtient une FilterPivotHierarchy par son nom ou id.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Méthode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtient un FilterPivotHierarchy par nom. Si la FilterPivotHierarchy n’existe pas, cela renvoie un objet null.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Méthode_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Supprime le PivotHierarchy de l’axe en cours.|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Propriété_ > inCellDropDown|Affiche la liste dans la cellule déroulante ou non, sa valeur par défaut est true.|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Propriété_ > source|Source de la liste de validation des données|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriété_ > id|ID du champ PivotField. En lecture seule.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriété_ > name|Nom du champ PivotField.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriété_ > showAllItems|Détermine si vous voulez afficher tous les éléments de PivotField.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relation_ > items|Renvoie les PivotFields associés à PivotField. En lecture seule.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relation_ > subtotals|Sous-totaux du champ PivotField.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Méthode_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|Trie le PivotField. Si une DataPivotHierarchy est spécifiée, le tri sera appliqué en fonction de celle-ci, sinon le tri sera basé sur le PivotField lui-même.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Propriété_ > items|Collection d’objets pivotField. En lecture seule.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Méthode_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|Obtient le nombre de hiérarchies croisées de la collection.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Méthode_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtient une PivotHierarchy par son nom ou id.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Méthode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtient une PivotHierarchy par nom. Si la PivotHierarchy n’existe pas, cela renvoie un objet null.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Propriété_ > id|ID de la PivotHierarchy. En lecture seule.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Propriété_ > name|Nom de la PivotHierarchy.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Relation_ > fields|Renvoie les PivotFields associés à la PivotHierarchy. En lecture seule.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Propriété_ > items|Une collection d’objets PivotHierarchy. En lecture seule.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Méthode_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|Obtient le nombre de hiérarchies croisées de la collection.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Méthode_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtient une PivotHierarchy par son nom ou id.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Méthode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtient une PivotHierarchy par nom. Si la PivotHierarchy n’existe pas, cela renvoie un objet null.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriété_ > id|ID du champ PivotItem. En lecture seule.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriété_ > isExpanded|Détermine si l’élément est développé pour afficher les éléments enfants ou si ce dernier est réduit et les éléments enfants sont masqués.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriété_ > name|Nom du champ PivotItem.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriété_ > visible|Détermine si le PivotItem est visible ou non.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Propriété_ > items|Collection d’objets pivotItem. En lecture seule.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Méthode_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|Obtient le nombre de hiérarchies croisées de la collection.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Méthode_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtient une PivotHierarchy par son nom ou id.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Méthode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtient une PivotHierarchy par nom. Si la PivotHierarchy n’existe pas, cela renvoie un objet null.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriété_ > showColumnGrandTotals|True si le rapport de tableau croisé dynamique affiche les grands totaux des colonnes.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriété_ > showRowGrandTotals|True si le rapport de tableau croisé dynamique affiche les grands totaux des lignes.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriété_ > subtotalLocation|Cette propriété indique le SubtotalLocationType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null. Les valeurs possibles sont AtTop, AtBottom.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Relation_ > layoutType|Cette propriété indique le PivotLayoutType de tous les champs sur le tableau croisé dynamique. Si les champs ont des états différents, la valeur sera null.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Méthode_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|Renvoie la plage où les étiquettes de colonnes de tableau croisé dynamique se trouvent.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Méthode_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|Renvoie la plage où les valeurs de données de tableau croisé dynamique se trouvent.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Méthode_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|Renvoie la plage de la zone de filtre de tableau croisé dynamique.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Méthode_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|Renvoie la plage sur laquelle le tableau croisé dynamique existe, à l’exception de la zone de filtre.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Méthode_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|Renvoie la plage où les étiquettes de lignes de tableau croisé dynamique se trouvent.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relation_ > columnHierarchies|Les hiérarchies de colonne de tableau croisé dynamique. En lecture seule.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relation_ > dataHierarchies|Les hiérarchies de données de tableau croisé dynamique. En lecture seule.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relation_ > filterHierarchies|Les hiérarchies de filtre de tableau croisé dynamique. En lecture seule.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relation_ > hiérarchies|Les hiérarchies Pivot de tableau croisé dynamique. En lecture seule.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relation_ > layout|Le PivotLayout décrivant la disposition et la structure visuelle de tableau croisé dynamique. En lecture seule.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relation_ > rowHierarchies|Les hiérarchies de lignes de tableau croisé dynamique. En lecture seule.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Méthode_ > [delete()](/javascript/api/excel/excel.pivottable)|Supprime le tableau croisé dynamique.|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Méthode_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|Ajoute un tableau croisé dynamique en fonction des données sources spécifiées et les insère à la cellule supérieure gauche de la plage de destination.|1.8|
|[range](/javascript/api/excel/excel.range)|_Relation_ > dataValidation|Renvoie un objet de validation des données. En lecture seule.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriété_ > id|ID de la RowColumnPivotHierarchy. En lecture seule.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriété_ > name|Nom de la RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriété_ > position|Position de la RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Relation_ > fields|Renvoie les PivotFields associés à la RowColumnPivotHierarchy. En lecture seule.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Méthode_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|Restaurer la RowColumnPivotHierarchy à ses valeurs par défaut.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Propriété_ > items|Une collection d’objets rowColumnPivotHierarchy. En lecture seule.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Méthode_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Ajoute le PivotHierarchy à l’axe en cours. Si la hiérarchie est présente ailleurs sur la ligne, colonne,|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Méthode_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtient le nombre de hiérarchies croisées de la collection.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Méthode_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtient une RowColumnPivotHierarchy par son nom ou id.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Méthode_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtient une RowColumnPivotHierarchy par nom. Si la RowColumnPivotHierarchy n’existe pas, cela renvoie un objet null.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Méthode_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Supprime le PivotHierarchy de l’axe en cours.|1.8|
|[runtime](/javascript/api/excel/excel.runtime)|_Propriété_ > enableEvents|Activer/désactiver les événements JavaScript dans le volet de tâches en cours ou complément de contenu.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relation_ > baseField|La base PivotField pour le calcul ShowAs, le cas échéant en fonction du type ShowAsCalculation, sinon null.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relation_ > baseItem|La base Item pour le calcul ShowAs, le cas échéant en fonction du type ShowAsCalculation, sinon null.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relation_ > calculation|Le calcul ShowAs à utiliser pour le champ de données PivotField.|1.8|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > autoIndent|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|1.8|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > textOrientation|L’orientation du texte pour le style.|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > automatic|Si Automatic est défini sur true, toutes les autres valeurs seront ignorées lorsque vous configurez les sous-totaux.|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > average| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > count| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > countNumbers| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > max| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > min| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > product| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > standardDeviation| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > standardDeviationP| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > sum| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > variance| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriété_ > varianceP| |1.8|
|[table](/javascript/api/excel/excel.table)|_Propriété_ > legacyId|Renvoie un ID numérique. Lecture seule.|1.8|
|[workbook](/javascript/api/excel/excel.workbook)|_Propriété_ > readOnly|True si le classeur est ouvert en mode lecture seule. En lecture seule.|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Propriété_ > id|Renvoie une valeur qui permet d’identifier l’objet WorkbookCreated de manière unique. En lecture seule.|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Méthode_ > [open()](/javascript/api/excel/excel.workbookcreated)|Ouvrez le classeur.|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriété_ > showGridlines|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul.|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriété_ > showHeadings|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul.|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Propriété_ > type|Obtient le type de l’événement.|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul qui est calculée.|1.8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Nouveautés de l’API JavaScript 1.7 pour Excel

Les fonctionnalités Excel JavaScript API ensemble de conditions 1.7 incluent des API pour les graphiques, événements, feuilles de calcul, plages, propriétés de document, éléments nommés, options de protection et styles.

### <a name="customize-charts"></a>Personnaliser des graphiques

Avec le nouvel API graphique, vous pouvez créer des types de graphiques supplémentaires, ajouter une série de données à un graphique, définir le titre du graphique, ajouter un titre d’axe, ajouter une unité d’affichage, ajouter une courbe de tendance avec moyenne mobile, modifier une courbe de tendance en ligne, et bien plus encore. Voici quelques exemples :

* Axe du graphique - obtenir, définir, mettre en forme et supprimer une unité d’axe, une étiquette et un titre dans un graphique.
* Série de graphique - ajouter, configurer et supprimer une série dans un graphique.  Modifier les marqueurs de série, les commandes traçage et le redimensionnement.
* Courbes de tendance de graphique - ajouter, obtenir et mettre en forme des courbes de tendance dans un graphique.
* Légende de graphique - mettre en forme la police de légende dans un graphique.
* Point de graphique - définir la couleur du point de graphique.
* Sous-chaîne de titre du graphique - obtenir et définir une sous-chaîne de titre d’un graphique.
* Type de graphique - option pour créer plusieurs types de graphiques.

### <a name="events"></a>Événements

Les API Événements pour Excel fournissent un grand nombre de gestionnaires d’événements autorisant votre complément à exécuter automatiquement une fonction désignée lorsqu’un événement spécifique se produit. Vous pouvez créer cette fonction pour effectuer n’importe quelle action nécessaire à votre scénario. Pour une liste des événements qui sont actuellement disponibles, voir [Manipuler des Événements à l’aide de l’API JavaScript Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personnaliser l’apparence de feuilles de calcul et des plages

À l’aide des nouveaux API, vous pouvez personnaliser l’apparence de feuilles de calcul de plusieurs façons :

* Figer les volets pour conserver certaines lignes ou colonnes visibles lorsque vous faites défiler la feuille de calcul. Par exemple, si la première ligne dans votre feuille de calcul contient des en-têtes, vous pouvez figer cette ligne de sorte que les en-têtes de colonne restent visibles pendant le défilement vers le bas de la feuille de calcul.
* Modifier la couleur d’onglet de la feuille de calcul.
* Ajouter des en-têtes de feuille de calcul.


Vous pouvez personnaliser l’apparence des plages de plusieurs façons :

* Définir le style de cellule pour une plage pour vous assurer que toutes les cellules dans la plage ont une mise en forme cohérente. Un style de cellule est un ensemble défini de caractéristiques de mise en forme, comme les polices et les tailles de police, formats des nombres, bordures de cellule et ombrage de cellule. Utilisez un des styles de cellule intégrés d’Excel ou créer votre propre style de cellule personnalisé.
* Définit l’orientation du texte pour une plage.
* Ajouter ou modifier un lien hypertexte sur une plage qui permet d’accéder à un autre emplacement dans le classeur ou à un emplacement externe.

### <a name="manage-document-properties"></a>Gérer les propriétés du document

À l’aide des API de propriétés du document, vous pouvez accéder aux propriétés de document intégrées et également créer et gérer les propriétés de document personnalisées pour stocker l’état du classeur et lire le flux de travail et la logique d’entreprise.

### <a name="copy-worksheets"></a>Obtenir des feuilles de calcul

À l’aide des API de copie de feuille de calcul , vous pouvez copier les données et le format à partir d’une feuille de calcul dans une nouvelle feuille de calcul au sein du même classeur et réduire la quantité de transfert de données nécessaire.

### <a name="handle-ranges-with-ease"></a>Gérer les plages en toute simplicité

À l’aide des API de plage différente, vous pouvez effectuer des actions telles qu’obtenir la région environnante, obtenir une plage redimensionnée et bien plus encore. Ces API doivent rendre des tâches telles que la manipulation de plage et l’adressage beaucoup plus efficaces.

De plus :

* Options de protection de classeur et feuille de calcul : utilisez ces API pour protéger les données dans une feuille de calcul et la structure du classeur.
* Mettre à jour un élément nommé : utilisez cet API pour mettre à jour un élément nommé.
* Obtenir la cellule active : utilisez cet API pour obtenir la cellule active d’un classeur.

|Objet| Quelles sont les nouveautés ?| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > chartType|Représente le type d’un graphique. Les valeurs possibles sont : ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > id|ID unique du graphique. En lecture seule.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > showAllFieldButtons|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Relation_ > border|Représente le format de bordure d’une zone de graphique, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Méthode_ > getItem(type: string, group: string)|Renvoie l’axe spécifique identifié par type et par groupe.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > axisBetweenCategories|Représente si l’axe de valeur croise l’axe de catégorie entre catégories.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > axisGroup|Représente le groupe pour l’axe spécifié. En lecture seule. Les valeurs possibles sont : primaire, secondaire.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > categoryType|Renvoie ou définit le type d’axe de catégorie. Les valeurs possibles sont : Automatic, TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > crosses|Représente l’axe spécifié où l’autre axe le croise. Les valeurs possibles sont : Automatic, Maximum, Minimum, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > crossesAt|Représente l’axe spécifié où l’autre axe le croise. En lecture seule. La configuration pour cette propriété doit utiliser la méthode SetCrossesAt(double). En lecture seule.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > customDisplayUnit|Représente la valeur unité d’affichage personnalisé d’axe. En lecture seule. Pour définir cette propriété, utilisez la méthode SetCustomDisplayUnit(double). En lecture seule.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > displayUnit|Représente l’unité d’affichage de l’axe. Les valeurs possibles sont : None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillions, Billions, Trillions, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > height|Représente la hauteur, exprimée en points, de l’axe de graphique. Null si l’axe n’est pas visible. En lecture seule.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > left|Représente la distance en points, du bord gauche de l’axe au bord gauche de la zone de graphique. Null si l’axe n’est pas visible. En lecture seule.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > logBase|Représente la base du logarithme lorsque vous utilisez des échelles logarithmiques.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > reversePlotOrder|Représente si Microsoft Excel trace des points de données à du dernier au premier.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > scaleType|Représente le type d’échelle de l’axe des ordonnées. Les valeurs possibles sont : linéaire, logarithmique.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > showDisplayUnitLabel|Représente si l’étiquette d’unité de l’affichage d’axe est visible.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > tickLabelSpacing|Représente le nombre de catégories ou séries entre les étiquettes de graduation. Peut être une valeur de 1 à 31999 ou une chaîne vide pour le paramètre automatique. La valeur renvoyée est toujours un nombre.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > tickMarkSpacing|Représente le nombre de catégories ou séries entre les graduations.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > top|Représente la distance en points, du bord supérieur de l’axe au bord supérieur de la zone de graphique. Null si l’axe n’est pas visible. En lecture seule.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > type|Représente le type d’axe. En lecture seule. Les valeurs possibles sont : non valide, catégorie, valeur, série.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > visible|Valeur booléenne qui représente la visibilité d’un axe.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriété_ > width|Représente la largeur, en points, de l’axe de graphique. Null si l’axe n’est pas visible. En lecture seule.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relation_ > baseTimeUnit|Renvoie ou définit l’unité de base pour l’axe des abscisses spécifié.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relation_ > majorTickMark|Représente le type de graduation principale pour l’axe spécifié.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > majorTimeUnitScale|Renvoie ou définit la valeur d’échelle d’unité principale pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relation_ > minorTickMark|Représente le type de graduation secondaire pour l’axe spécifié.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relation_ > minorTimeUnitScale|Renvoie ou définit la valeur d’échelle d’unité secondaire pour l’axe des abscisses lorsque la propriété CategoryType est définie sur échelle de temps.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relation_ > tickLabelPosition|Représentant la position des étiquettes de graduation sur l’axe spécifié.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Méthode_ > setCategoryNames(sourceData: Range)|Définit tous les noms de catégorie pour l’axe spécifié.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Méthode_ > setCrossesAt(value: double)|Représente l’axe spécifié où l’autre axe le croise.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Méthode_ > setCustomDisplayUnit(value: double)|Définit l’unité d’affichage axe sur une valeur personnalisée.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriété_ > color|Code couleur HTML qui représente la couleur des bordures dans le graphique.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriété_ > weight|Représente l’épaisseur de bordure, en points.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Relation_ > lineStyle|Représente le style de trait de la bordure.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > position|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Les valeurs possibles sont les suivantes : None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > separator|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > showBubbleSize|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > showCategoryName|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > showLegendKey|Valeur booléenne indiquant si la clé de légende des étiquettes de données est visible ou non.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > showPercentage|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > showSeriesName|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriété_ > showValue|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriété_ > height|Représente la hauteur de la légende sur le graphique.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriété_ > left|Représente la partie gauche d’une légende de graphique.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriété_ > showShadow|Représente si la légende possède une ombre sur le graphique.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriété_ > top|Représente la partie supérieure d’une légende de graphique.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriété_ > width|Représente la largeur de la légende sur le graphique.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Relation_ > legendEntries|Représente une collection de legendEntries dans la légende. En lecture seule.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriété_ > visible|Représente la partie visible d’une entrée de légende de graphique.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Propriété_ > items|Une collection d’objets chartLegendEntry. En lecture seule.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Méthode_ > getCount()|Renvoie le nombre de legendEntry de la collection.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Méthode_ > getItemAt(index: number)|Renvoie un legendEntry à l’index donné.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriété_ > hasDataLabel|Représente si un point de données a datalabel. Non applicable pour les graphiques en surface.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriété_ > markerBackgroundColor|Représentation de code de couleur HTML de la couleur d’arrière-plan du marqueur de données. Par exemple, #FF0000 représente le rouge.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriété_ > markerForegroundColor|Représentation de code de couleur HTML de la couleur de premier plan du point de données. Par exemple, #FF0000 représente le rouge.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriété_ > markerSize|Représente la taille du marqueur du point de données.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Propriété_ > markerStyle|Représente le style du marqueur du point de données de graphique. Les valeurs possibles sont : Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Relation_ > dataLabel|Renvoie l’étiquette de données d’un point du graphique. En lecture seule.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Relation_ > border|Représente le format de bordure d’un point de données de graphique, qui inclut les informations de couleur, style de ligne et épaisseur. En lecture seule.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > chartType|Représente le type de graphique d’une série. Les valeurs possibles sont : ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > doughnutHoleSize|Représente la taille du centre d’une série de graphiques en anneaux.  Valide uniquement dans les graphiques en anneaux et doughnutExploded.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > filtered|Valeur booléenne représentant si la série est filtrée ou non. Non applicable pour les graphiques en surface.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > gapWidth|Représente la largeur de l’intervalle d’une série de graphique.  Valide uniquement sur les graphiques en barres et colonnes, ainsi que|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > hasDataLabels|Valeur booléenne représentant si la série a des étiquettes de données ou non.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > markerBackgroundColor|Représente la couleur d’arrière-plan de marqueurs d’une série de graphiques.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > markerForegroundColor|Représente la couleur de premier plan de marqueurs d’une série de graphiques.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > markerSize|Représente la taille du marqueur d’une série de graphiques.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > markerStyle|Représente le style du marqueur d’une série de graphiques. Les valeurs possibles sont : Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > plotOrder|Représente l’ordre de traçage d’une série de graphiques au sein du groupe de graphiques.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > showShadow|Valeur booléenne représentant si la série a une ombre ou non.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriété_ > smooth|Valeur booléenne représentant si la série est fluide ou non. Uniquement pour les graphiques en lignes et en nuages de points.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relation_ > dataLabels|Représente la collection de tous les dataLabels de la série. En lecture seule.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relation_ > trendlines|Représente la collection de courbes de tendance de la série. En lecture seule.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Méthode_ > delete()|Supprime la série graphique.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Méthode_ > setBubbleSizes(sourceData: Range)|Définit des tailles de bulles pour une série de graphiques. Fonctionne uniquement pour les graphiques en bulles.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Méthode_ > setValues(sourceData: Range)|Définit des valeurs pour une série de graphiques. Pour un graphique en nuages de points, cela signifie des valeurs de l’axe Y.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Méthode_ > setXAxisValues(sourceData: Range)|Définit des valeurs d’axe Y pour une série de graphiques. Fonctionne uniquement pour les graphiques en nuages de points.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Méthode_ > add(name: string, index: number)|Ajouter une nouvelle série à la collection.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > height|Représente la hauteur, exprimée en points, du titre du graphique. En lecture seule. Null si le titre du graphique n’est pas visible. En lecture seule.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > horizontalAlignment|Représente l’alignement horizontal du titre du graphique. Les valeurs possibles sont : Center, Left, Justify, Distributed, Right.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > left|Représente la distance en points, du bord gauche du titre du graphique au bord gauche de la zone de graphique. Null si le titre du graphique n’est pas visible.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > position|Représente la position du titre du graphique. Les valeurs possibles sont : Top, Automatic, Bottom, Right, Left.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > showShadow|Représente une valeur booléenne qui détermine si le titre du graphique possède une ombre.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > textOrientation|Représente l’orientation du texte du titre du graphique. La valeur doit être un entier soit de -90 à 90, soit 180 pour le texte orienté verticalement.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > top|Représente la distance en points, du bord supérieur du titre du graphique au bord supérieur de la zone de graphique. Null si le titre du graphique n’est pas visible.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > verticalAlignment|Représente l’alignement vertical du titre du graphique. Les valeurs possibles sont : Center, Bottom, Top, Justify, Distributed.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriété_ > width|Représente la largeur, exprimée en points, du titre du graphique. En lecture seule. Null si le titre du graphique n’est pas visible. En lecture seule.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Méthode_ > setFormula(formula: string)|Définit une valeur de chaîne qui représente la formule de titre de graphique à l’aide de la notation de style A1.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Relation_ > border|Représente le format de bordure du titre de graphique, qui inclut couleur, style de ligne et épaisseur. En lecture seule.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > backward|Représente le nombre de points que la courbe de tendance étend en arrière.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > displayEquation|Vrai si l’équation de la courbe de tendance est affichée sur le graphique.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > displayRSquared|Vrai si le coefficient de détermination de la courbe de tendance est affiché sur le graphique.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > forward|Représente le nombre de points que la courbe de tendance étend en avant.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > intercept|Représente la valeur intercept de la courbe de tendance. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs automatiques). La valeur renvoyée est toujours un nombre.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > movingAveragePeriod|Représente le point d’une courbe de tendance graphique, uniquement pour les courbes de tendance avec type MovingAverage.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > name|Représente le nom de la courbe de tendance. Peut être configurée pour une valeur de chaîne, ou peut être configurée pour que la valeur null représente les valeurs automatiques. La valeur renvoyée est toujours une chaîne.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > polynomialOrder|Représente l’ordre d’une courbe de tendance graphique, uniquement pour les courbes de tendance avec type Polynomial.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriété_ > type|Représente le type de courbe de tendance de graphique. Les valeurs possibles sont : Linear, Exponential, Logarithmic, MovingAverage, Polynomial, Power.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relation_ > format|Représente la mise en forme de courbe de tendance de graphique. En lecture seule.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Méthode_ > delete()|Supprime l’objet courbe de tendance.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Propriété_ > items|Collection d’objets chartTrendline. En lecture seule.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Méthode_ > add(type: string)|Ajoute une nouvelle courbe de tendance à la collection de courbes de tendance.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Méthode_ > getCount()|Renvoie le nombre de courbes de tendance de la collection.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Méthode_ > getItem(index: number)|Obtient un objet courbe de tendance par index, c'est-à-dire par ordre d’insertion dans le tableau des éléments.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Relation_ > line|Représente le format de lignes du graphique. En lecture seule.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriété_ > key|Obtient la clé de la propriété personnalisée. En lecture seule. En lecture seule.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriété_ > type|Obtient le type de valeur de la propriété personnalisée. En lecture seule. En lecture seule. Les valeurs possibles sont : Number, Boolean, Date, String, Float.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriété_ > value|Obtient ou définit la valeur de la propriété personnalisée.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Méthode_ > delete()|Supprime la propriété personnalisée.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Propriété_ > items|Collection d’objets customProperty. En lecture seule.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Méthode_ > add(key: string, value: object)|Crée une nouvelle propriété personnalisée ou en définit une existante.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Méthode_ > deleteAll()|Supprime toutes les propriétés personnalisées de cette collection.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Méthode_ > getCount()|Obtient le nombre des propriétés personnalisées.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Méthode_ > getItem(key: string)|Obtient un objet de propriété personnalisée par sa clé, qui ne tient pas compte de la casse. Ignoré si la propriété personnalisée n’existe pas.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Méthode_ > getItemOrNullObject(key: string)|Obtient un objet de propriété personnalisée par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Propriété_ > items|Collection d’objets dataConnection. En lecture seule.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Méthode_ > refreshAll()|Actualise toutes les dataConnections de la collection.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > author|Obtient ou définit l’auteur du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > category|Obtient ou définit la catégorie du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > comments|Obtient ou définit les commentaires du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > company|Obtient ou définit la compagnie du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > keywords|Obtient ou définit les mots clés du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > lastAuthor|Obtient ou définit le dernier auteur du classeur. En lecture seule. En lecture seule.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > manager|Obtient ou définit le responsable du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > revisionNumber|Obtient le numéro de révision du classeur. En lecture seule.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > subject|Obtient ou définit le sujet du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriété_ > title|Obtient ou définit le titre du classeur.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relation_ > creationDate|Obtient la date de création du classeur. En lecture seule. En lecture seule.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relation_ > custom|Obtient la collection de propriétés personnalisées du classeur. En lecture seule. En lecture seule.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriété_ > formula|Obtient ou définit la formule de l’élément nommé.  La formule commence toujours par un signe '='.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relation_ > arrayValues|Renvoie un objet contenant les valeurs et les types de l’élément nommé. En lecture seule.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriété_ > types|Représente les types de chaque élément dans le tableau élément nommé accessible en lecture seule. Les valeurs possibles sont : Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriété_ > values|Représente les valeurs de chaque élément dans le tableau élément nommé. En lecture seule.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > isEntireColumn|Représente si la plage active est une colonne entière. En lecture seule.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > isEntireRow|Représente si la plage active est une ligne entière. En lecture seule.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > numberFormatLocal|Représente le code de format numérique d’Excel pour la plage donnée en tant que chaîne dans la langue de l’utilisateur.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > style|Représente le style de la plage actuelle. Ceci renvoie soit null, soit une chaîne.|1.7|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > getAbsoluteResizedRange(numRows: number, numColumns: number)|Obtient un objet Plage avec la même cellule supérieure gauche que l’objet de Plage en cours, mais avec un nombre spécifié de lignes et colonnes.|1.7|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > getImage()|Affiche la plage en tant qu’image codée en base 64.|1.7|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > getSurroundingRegion()|Renvoie un objet PLage qui représente la région environnante pour la cellule en haut à gauche de cette plage. Une région environnante est une plage délimitée par une combinaison de lignes et de colonnes vides par rapport à cette plage.|1.7|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > showCard()|Affiche la carte pour une cellule active si son contenu est riche en valeur.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriété_ > textOrientation|Obtient ou définit l’orientation du texte de toutes les cellules dans la plage.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriété_ > useStandardHeight|Détermine si la hauteur de ligne de l’objet de plage est égale à la hauteur standard de la feuille.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriété_ > useStandardWidth|Détermine si la largeur de colonne de l’objet de plage est égale à la largeur standard de la feuille.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriété_ > address|Représente l’url cible du lien hypertexte.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriété_ > document...|Représente le document... cible du lien hypertexte.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriété_ > screenTip|Représente la chaîne affichée lorsque vous pointez sur le lien hypertexte.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriété_ > textToDisplay|Représente la chaîne qui s’affiche dans la cellule en haut à gauche de la plage.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > addIndent|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > autoIndent|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > builtIn|Indique si le style est un style intégré. En lecture seule.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > formulaHidden|Indique si la formule est masquée lorsque la feuille de calcul est protégée.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > horizontalAlignment|Représente l’alignement horizontal pour le style. Les valeurs possibles sont : General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > includeAlignment|Indique si le style inclut les propriétés AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, et TextOrientation.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeBorder|Indique si le style inclut les propriétés dColor, ColorIndex, LineStyle, et Weight border.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > includeFont|Indique si le style inclut les propriétés dBackground, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, et Underline font.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeNumber|Indique si le style inclut la propriété NumberFormat.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includePatterns|Indique si le style inclut les propriétés Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, et PatternColorIndex interior.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > includeProtection|Indique si le style inclut les propriétés FormulaHidden et Locked protection.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > indentLevel|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > locked|Indique si l’objet est verrouillé lorsque la feuille de calcul est protégée.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > name|Nom du style. En lecture seule.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > numberFormat|Le code de format du nombre format pour le style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > numberFormatLocal|Le code de format localisé du nombre format pour le style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > orientation|L’orientation du texte pour le style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > readingOrder|L’ordre de lecture du style. Les valeurs possibles sont : Context, LeftToRight, RightToLeft.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > shrinkToFit|Indique si le texte s’ajuste automatiquement pour tenir dans la largeur de colonne disponible.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > textOrientation|L’orientation du texte pour le style.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > verticalAlignment|Représente l’alignement vertical du style. Les valeurs possibles sont : Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriété_ > wrapText|Indique si Microsoft Excel renvoie le texte à la ligne dans l’objet.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relation_ > borders|Collection de bordures de quatre objets qui représente le style des quatre bordures. En lecture seule.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relation_ > fill|Remplissage du style. En lecture seule.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relation_ > font|Renvoie un objet Police qui représente la police du style. En lecture seule.|1.7|
|[style](/javascript/api/excel/excel.style)|_Méthode_ > delete()|Supprime ce style.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Propriété_ > items|Collection d’objets de style. En lecture seule.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Méthode_ > add(name: string)]|Ajoute un nouveau style à la collection.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Méthode_ > getItem(name: string)|Obtient un style par nom.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriété_ > address|Obtient l’adresse qui représente la zone modifiée d’un tableau figurant dans une feuille de calcul spécifique.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriété_ > changeType|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié. Les valeurs possibles sont : Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriété_ > source|Obtient la source de l’événement. Les valeurs possibles sont : Local, Remote.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriété_ > tableId|Obtient l’id du tableau dans lequel les données sont modifiées.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriété_ > address|Obtient l’adresse de plage qui représente la zone sélectionnée d’un tableau dans une feuille de calcul spécifique.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriété_ > isInsideTable|Indique si la sélection est dans un tableau, l’adresse sera superflue si IsInsideTable est faux.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriété_ > tableId|Obtient l’id du tableau dans lequel la sélection est modifiée.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Propriété_ > name|Obtient le nom du classeur. En lecture seule.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > dataConnections|Actualise toutes les dataConnections du classeur. En lecture seule.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > properties|Obtient les propriétés du classeur. En lecture seule.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > protection|Renvoie un objet de protection de classeur pour un classeur. En lecture seule.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > styles|Représente une collection de styles associés au classeur. En lecture seule.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Méthode_ > getActiveCell()|Obtient la cellule active du classeur.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Propriété_ > protected|Indique si le classeur est protégé. En lecture seule. En lecture seule.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Méthode_ > protect(password: string)|Protège un classeur. Échoue si le classeur est protégé.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Méthode_ > unprotect(password: string)|Annule la protection un classeur.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriété_ > gridlines|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriété_ > headings|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriété_ > showHeadings|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriété_ > standardHeight|Renvoie la hauteur standard (par défaut) de toutes les lignes dans la feuille de calcul, en points. En lecture seule.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriété_ > standardWidth|Renvoie ou définit la largeur standard (par défaut) de toutes les colonnes dans la feuille de calcul.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > tabColor|Obtient ou modifie la couleur d’onglet de la feuille de calcul.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relation_ > freezePanes|Obtient un objet qui peut être utilisé pour manipuler les volets figés sur la feuille de calcul accessible en lecture seule.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|Copie une feuille de calcul et la place à la position spécifiée. Renvoie la feuille de calcul copiée.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|Obtient l’objet plage commençant à un index de ligne et de colonne particulier et couvrant un certain nombre de lignes et de colonnes.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul qui est activée.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriété_ > source|Obtient la source de l’événement. Les valeurs possibles sont : Local, Remote.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul qui est ajoutée au classeur.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriété_ > address|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriété_ > changeType|Obtient le type de modification qui représente la manière dont est déclenché l’événement modifié. Les valeurs possibles sont : Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriété_ > source|Obtient la source de l’événement. Les valeurs possibles sont : Local, Remote.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul qui est desactivée.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriété_ > source|Obtient la source de l’événement. Les valeurs possibles sont : Local, Remote.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul qui est supprimée du classeur.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Méthode_ > freezeAt(frozenRange: Range or string)|Définit les cellules figées dans l’affichage de la feuille de calcul active.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Méthode_ > freezeColumns(count: number)|Figer la/les première(s) colonne(s) de la feuille de calcul en place.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Méthode_ > freezeRows(count: number)|Figer la/les première(s) ligne(s) de la feuille de calcul en place.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Méthode_ > getLocation()|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Méthode_ > getLocationOrNullObject()|Obtient une plage qui définit les cellules figées dans l’affichage de la feuille de calcul active.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Méthode_ > unfreeze()|Supprime tous les volets figés dans la feuille de calcul.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowEditObjects|Représente l’option de protection de feuille de calcul qui autorise la modification d’objets.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowEditScenarios|Représente l’option de protection de feuille de calcul qui autorise la modification de scénarios.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Relation_ > selectionMode|Représente l’option de protection de feuille de calcul qui autorise le mode sélection.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriété_ > address|Obtient l’adresse de plage qui représente la zone sélectionnée dans une feuille de calcul spécifique.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriété_ > type|Obtient le type de l’événement. Les valeurs possibles sont : WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriété_ > worksheetId|Obtient l’id de la feuille de calcul dans laquelle la sélection est modifiée.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Nouveautés de l’API JavaScript 1.6 pour Excel 

### <a name="conditional-formatting"></a>Mise en forme conditionnelle

Présente la mise en forme conditionnelle d’une plage. Autorise les types de mise en forme conditionnelle suivants :

* Échelle de couleurs
* Barre de données
* Jeu d'icônes
* Personnalisé

De plus :

* Renvoie la plage à laquelle s’applique la mise en forme conditionnelle. 
* Supprime la mise en forme conditionnelle. 
* Offre une fonctionnalité de priorité et stopifTrue. 
* Obtient la collection de toutes les mises en forme conditionnelles sur une plage donnée. 
* Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle. 

|Objet| Quelles sont les nouveautés ?| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Méthode_ > suspendApiCalculationUntilNextSync()|Interrompt le calcul jusqu'à ce que la prochaine méthode « context.sync() » soit appelée. Une fois cette option définie, il incombe au développeur de recalculer le classeur afin de garantir que toutes les dépendances sont propagées.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relation_ > format|Renvoie un objet de format, qui comprend les polices, remplissage, bordures des mises en formes conditionnelles, et d’autres propriétés. Lecture seule.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relation_ > rule|Représente l’objet Règle sur cette mise en forme conditionnelle.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Propriété_ > threeColorScale|Si la valeur est True, l’échelle de couleur comporte trois points (minimum, milieu, maximum). Sinon elle en comporte deux (minimum, maximum). Lecture seule.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Relation_ > criteria|Critère de l’échelle de couleur. Le point Milieu est facultatif lorsque vous utilisez une échelle de couleur à deux points.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriété_ > formula1|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriété_ > formula2|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriété_ > operator|L’opérateur de mise en forme conditionnelle du texte. Les valeurs possibles sont : Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relation_ > maximum|Point maximal du critère d’échelle de couleurs.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > midpoint|Point du milieu du critère d’échelle de couleurs, si l’échelle de couleurs est une échelle à 3 couleurs.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relation_ > minimum|Point minimal du critère d’échelle de couleurs.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriété_ > color|Représentation sous forme de code couleur HTML de la couleur de l’échelle de couleurs. Par exemple, #FF0000 représente le rouge.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriété_ > formula|Nombre, formule ou null (si le type est LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriété_ > type|Ce sur quoi la formule conditionnelle icône doit être basée. Les valeurs possibles sont : Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriété_ > borderColor|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriété_ > fillColor|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriété_ > matchPositiveBorderColor|Représentation booléenne indiquant si la barre de données négative a une bordure de la même couleur que la barre de données positive.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > matchPositiveFillColor|Représentation booléenne indiquant si la barre de données négative a un remplissage de la même couleur que la barre de données positive.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriété_ > borderColor|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriété_ > fillColor|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > gradientFill|Représentation booléenne indiquant si la barre de données a un dégradé.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriété_ > formula|Formule, si nécessaire, servant à évaluer la règle de la barre de données.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriété_ > type|Type de règle pour la barre de données. Les valeurs possibles sont : LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriété_ > id|La priorité de la mise en forme conditionnelle dans la ConditionalFormatCollection actuelle. En lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriété_ > priority|Priorité (ou index) dans la collection de mise en forme conditionnelle dans laquelle cette mise en forme conditionnelle existe actuellement. Cette modification également|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriété_ > stopIfTrue|Si les conditions de cette mise en forme conditionnelle sont remplies, aucun format de priorité inférieure ne doit prendre effet sur cette cellule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriété_ > type|Type de mise en forme conditionnelle. Un seul peut être défini à la fois. Lecture seule. Lecture seule. Les valeurs possibles sont : Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > cellValue|Renvoie les propriétés de mise en forme conditionnelle de la valeur de la cellule si la mise en forme conditionnelle actuelle est un type CellValue. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > cellValueOrNullObject|Renvoie les propriétés de mise en forme conditionnelle de la valeur de la cellule si la mise en forme conditionnelle actuelle est un type CellValue. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > colorScale|Renvoie les propriétés de mise en forme condittionnelle ColorScale si la mise en forme conditionnelle actuelle est un type ColorScale. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > colorScaleOrNullObject|Renvoie les propriétés de mise en forme condittionnelle ColorScale si la mise en forme conditionnelle actuelle est un type ColorScale. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > custom|Renvoie les propriétés de mise en forme conditionnelle personnalisée si la mise en forme conditionnelle actuelle est un type personnalisé. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > customOrNullObject|Renvoie les propriétés de mise en forme conditionnelle personnalisée si la mise en forme conditionnelle actuelle est un type personnalisé. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > dataBar|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > dataBarOrNullObject|Renvoie les propriétés de la barre de données si la mise en forme conditionnelle actuelle est une barre de données.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > iconSet|Renvoie les propriétés de mise en forme conditionnelle IconSet si la mise en forme conditionnelle actuelle est un type IconSet. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > iconSetOrNullObject|Renvoie les propriétés de mise en forme conditionnelle IconSet si la mise en forme conditionnelle actuelle est un type IconSet. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > preset|Renvoie la mise en forme conditionnelle des critères prédéfinis comme les propriétés averagebelow averageunique valuescontains blanknonblankerrornoerror ci-dessus. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > presetOrNullObject|Renvoie la mise en forme conditionnelle des critères prédéfinis comme les propriétés averagebelow averageunique valuescontains blanknonblankerrornoerror ci-dessus. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > textComparison|Renvoie les propriétés de mise en forme conditionnelle du texte spécifique si la mise en forme conditionnelle actuelle est un type de texte. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > textComparisonOrNullObject|Renvoie les propriétés de mise en forme conditionnelle du texte spécifique si la mise en forme conditionnelle actuelle est un type de texte. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > topBottom|Renvoie les propriétés de mise en forme conditionnelle TopBottom si la mise en forme conditionnelle actuelle est un type TopBottom. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relation_ > topBottomOrNullObject|Renvoie les propriétés de mise en forme conditionnelle TopBottom si la mise en forme conditionnelle actuelle est un type TopBottom. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Méthode_ > delete()|Supprime cette mise en forme conditionnelle.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Méthode_ > getRange()|Renvoie la plage à laquelle s’applique la mise en forme conditionnelle ou un objet null si la plage est discontinue. Lecture seule.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Méthode_ > getRangeOrNullObject()|Renvoie la plage à laquelle s’applique la mise en forme conditionnelle ou un objet null si la plage est discontinue. Lecture seule.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Propriété_ > items|Collection d’objets conditionalFormat. Lecture seule.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Méthode_ > add(type: string)|Ajoute une nouvelle mise en forme conditionnelle à la collection à la priorité firsttop.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Méthode_ > clearAll()|Efface toutes les mises en forme conditionnelles actives sur la plage spécifiée actuelle.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Méthode_ > getCount()|Renvoie le nombre de mises en formes conditionnelles dans le classeur. Lecture seule.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Méthode_ > getItem(id: string)|Renvoie une mise en forme conditionnelle à un ID donné.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Méthode_ > getItemAt(index: number)|Renvoie une mise en forme conditionnelle à l’index donné.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriété_ > formula|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriété_ > formulaLocal|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la langue de l’utilisateur.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriété_ > formulaR1C1|Formule, si nécessaire, servant à évaluer la règle de mise en forme conditionnelle dans la notation du style R1C1.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriété_ > formula|Un nombre ou une formule en fonction du type.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriété_ > operator|GreaterThan ou GreaterThanOrEqual pour chaque du type de règle pour la mise en forme conditionnelle de l’icône. Les valeurs possibles sont : Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relation_ > customIcon|Icône personnalisée pour le critère en cours si différent de la celui par défaut IconSet. Sinon, null est renvoyé.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relation_ > type|Ce sur quoi la formule conditionnelle de l’icône doit être basée.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Propriété_ > critère|Critère du format conditionnel. Les valeurs possibles sont les suivantes : Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriété_ > color|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriété_ > id|Représente l’identificateur de la bordure. En lecture seule. Les valeurs possibles sont les suivantes : EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriété_ > sideIndex|Valeur constante qui indique un côté spécifique de la bordure. En lecture seule. Les valeurs possibles sont les suivantes : EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriété_ > style|L’une des constantes de style de ligne déterminant le style de ligne de la bordure. Les valeurs possibles sont les suivantes : None (aucune), Continuous (continue), Dash (tirets), DashDot (ligne tiret-point), DashDotDot (ligne tiret-point-point), Dot (points), Double (double), SlantDashDot (ligne tiret-point oblique).|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriété_ > count|Nombre d’objets de bordure de la collection. En lecture seule.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriété_ > items|Collection d’objets conditionalRangeBorder. En lecture seule.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relation_ > bottom|Obtient la bordure supérieure en lecture seule.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relation_ > left|Obtient la bordure supérieure en lecture seule.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relation_ > right|Obtient la bordure supérieure en lecture seule.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relation_ > top|Obtient la bordure supérieure en lecture seule.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Méthode_ > getItem(index: string)|Obtient un objet de bordure à l’aide de son nom.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Méthode_ > getItemAt(index: number)|Obtient un objet de bordure à l’aide de son index.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Propriété_ > color|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Méthode_ > clear()|Réinitialise le remplissage.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriété_ > bold|Représente le format de police Gras.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriété_ > color|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriété_ > italic|Représente le format de police Italique.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriété_ > strikethrough|Représente l’état barré de la police.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriété_ > underline|Type de soulignement appliqué à la police. Les valeurs possibles sont les suivantes : None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Méthode_ > clear()|Réinitialise les formats de police.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Propriété_ > numberFormat|Représente le code de format de nombre d’Excel pour une plage donnée. Ignoré si null est indiqué.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relation_ > borders|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage de mise en forme conditionnelle. Lecture seule.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relation_ > fill|Retourne l’objet de remplissage défini sur l’ensemble de la plage de mise en forme conditionnelle. En lecture seule.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relation_ > font|Retourne l’objet de police défini sur l’ensemble de la plage de mise en forme conditionnelle. En lecture seule.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriété_ > operator|L’opérateur de mise en forme conditionnelle du texte. Les valeurs possibles sont : Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriété_ > text|Valeur de texte de la mise en forme conditionnelle.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriété_ > rank|Rang compris entre 1 et 1000 pour les rangs numériques ou entre 1 et 100 pour les rangs en pourcentage.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriété_ > type|Valeurs de mis en forme basées sur le rang supérieur ou inférieur. Les valeurs possibles sont : Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relation_ > format|Renvoie un objet de format, qui comprend les polices, remplissage, bordures des mises en formes conditionnelles, et d’autres propriétés. Lecture seule.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relation_ > rule|Représente l’objet Règle sur cette mise en forme conditionnelle. En lecture seule.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriété_ > axisColor|Code couleur HTML qui représente la couleur de la ligne Axe, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriété_ > axisFormat|Représentation de comment l’axe est déterminé pour une barre de données Excel. Les valeurs possibles sont : Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriété_ > barDirection|Représente la direction sur laquelle le graphique de barre de données doit être basé. Les valeurs possibles sont : Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriété_ > showDataBarOnly|Si la valeur est True, masque les valeurs des cellules où la barre de données est appliquée.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relation_ > lowerBoundRule|Règle de ce qui constitue la limite inférieure (et comment la calculer, le cas échéant) pour une barre de données.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relation_ > negativeFormat|Représentation de toutes les valeurs à gauche de l’axe dans une barre de données Excel. En lecture seule.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relation_ > positiveFormat|Représentation de toutes les valeurs à droite de l’axe dans une barre de données Excel. En lecture seule.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relation_ > upperBoundRule|Règle de ce qui constitue la limite supérieure (et comment la calculer, le cas échéant) pour une barre de données.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriété_ > reverseIconOrder|Si True, inverse les ordres d’icône pour le IconSet. Notez que ceci ne peut pas être défini si des icônes personnalisées sont utilisés.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriété_ > showIconOnly|Si la valeur est True, masque les valeurs et affiche uniquement les icônes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriété_ > style|Si défini, affiche l’option IconSet pour la mise en forme conditionnelle. Les valeurs possibles sont les suivantes : Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Relation_ > criteria|Tableau de critères et d’IconSets pour les règles et icônes personnalisées potentielles pour les icônes conditionnelles. Notez que pour le premier critère, seule l’icône personnalisée peut être modifiée, tandis que le type, la formule et l’opérateur sont ignorés, si défini.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relation_ > format|Renvoie un objet de format, qui comprend les polices, remplissage, bordures des mises en formes conditionnelles, et d’autres propriétés. Lecture seule.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relation_ > rule|Règle de mise en forme conditionnelle.|1.6|
|[range](/javascript/api/excel/excel.range)|_Relation_ > conditionalFormats|Collection de mises en formes conditionnelles qui ont une intersection avec la plage. En lecture seule.|1.6|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > calculate()|Calcule une plage de cellules dans une feuille de calcul.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relation_ > format|Renvoie un objet de format, qui comprend les polices, remplissage, bordures des mises en formes conditionnelles, et d’autres propriétés. Lecture seule.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relation_ > rule|Règle de mise en forme conditionnelle.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relation_ > format|Renvoie un objet de format, qui comprend les polices, remplissage, bordures des mises en formes conditionnelles, et d’autres propriétés. Lecture seule.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relation_ > rule|Critères de mise en forme conditionnelle TopBottom.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > internalTest|Réservé à un usage interne. En lecture seule.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > calculate(markAllDirty: bool)|Calcule toutes les cellules d’une feuille de calcul.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Nouveautés de l’API JavaScript 1.5 pour Excel

### <a name="custom-xml-part"></a>Partie XML personnalisée

* Ajout d’une collection de parties XML personnalisée à un objet workbook.
* Obtenir la partie XML personnalisée à l’aide de l’ID
* Obtenez une nouvelle collection délimitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.
* Obtenez une chaîne XML associée à une partie.
* Fournissez l’ID et l’espace de noms d’une partie.
* Ajoute une nouvelle partie XML personnalisée au classeur.
* Définissez une partie XML entière.
* Supprimez une partie XML personnalisée.
* Supprimez un attribut avec le nom donné dans l’élément identifié par langage XPath.
* Interrogez le contenu XML par langage XPath.
* Insérez, mettez à jour et supprimez l’attribut.

**Implémentation de référence :** Cliquez [ici](https://github.com/mandren/Excel-CustomXMLPart-Demo) pour obtenir une implémentation de référence qui décrit comment les parties XML personnalisées peuvent être utilisées dans un complément.

### <a name="others"></a>Autres
* `range.getSurroundingRegion()` Renvoie un objet Range qui représente la région environnante pour cette plage. Une région environnante est une plage délimitée par une combinaison de lignes et de colonnes vides par rapport à cette plage.
* `getNextColumn()` et `getPreviousColumn()`, `getLast() sur la colonne du tableau.
* `getActiveWorksheet()` sur le classeur.
* `getRange(address: string)` en dehors du classeur.
* `getBoundingRange(ranges: )` Renvoie le plus petit objet range qui englobe les plages fournies. Par exemple, la plage englobante entre « B2:C5 » et « D10:E15 » est « B2:E15 ».
* `getCount()` sur différentes collections (élément nommé, feuille de calcul, tableau, etc.) pour obtenir le nombre d’éléments dans une collection. `workbook.worksheets.getCount()`
* `getFirst()`et `getLast()` et get last sur différentes collections (feuille de calcul, colonne de tableau, points de graphique, vue de plage).
* `getNext()` et `getPrevious()` sur une collection de feuilles de calcul, colonnes de tableau.
* `getRangeR1C1()` Obtient l’objet plage commençant à un index de ligne et de colonne particulier et couvrant un certain nombre de lignes et de colonnes.

|Objet| Quelles sont les nouveautés ?| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Propriété_ > id|ID de la partie XML personnalisée. En lecture seule.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Propriété_ > namespaceUri|URI de l’espace de noms de la partie XML personnalisée. En lecture seule.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Méthode_ > delete()|Supprime la partie XML personnalisée.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Méthode_ > getNext()|Obtient l’intégralité du contenu XML de la partie XML personnalisée.|1,5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Méthode_ > setXml(xml: string)|Définit l’intégralité du contenu XML de la partie XML personnalisée.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Propriété_ > items|Collection d’objets customXmlPart. En lecture seule.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Méthode_ > add(xml: string)|Ajoute une nouvelle partie XML personnalisée au classeur.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Méthode_ > getByNamespace(namespaceUri: string)|Obtient une nouvelle collection limitée de parties XML personnalisées dont les espaces de noms correspondent à l’espace de noms donné.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Méthode_ > getCount()|Obtient le nombre de parties CustomXml dans la collection.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Méthode_ > getItem(id: string)|Obtient une partie XML personnalisée en fonction de son ID.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Méthode_ > getItemOrNullObject(id: string)|Obtient une partie XML personnalisée en fonction de son ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Propriété_ > items|Collection d’objets customXmlPartScoped. En lecture seule.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Méthode_ > getCount()|Obtient le nombre de parties CustomXML dans cette collection.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Méthode_ > getItem(id: string)|Obtient une partie XML personnalisée en fonction de son ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Méthode_ > getItemOrNullObject(id: string)|Obtient une partie XML personnalisée en fonction de son ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Méthode_ > getOnlyItem()|Si la collection contient exactement un élément, cette méthode le renvoie.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Méthode_ > getOnlyItemOrNullObject()|Si la collection contient exactement un élément, cette méthode le renvoie.|1,5|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > customXmlParts|Représente la collection de parties XML personnalisées contenues dans ce classeur. En lecture seule.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > getNext(visibleOnly: bool)|Obtient la feuille de calcul qui suit celle-ci. Si aucune feuille de calcul ne suit celle-ci, cette méthode génère une erreur.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > getNextOrNullObject(visibleOnly: bool)|Obtient la feuille de calcul qui suit celle-ci. Si aucune feuille de calcul ne suit celle-ci, cette méthode renvoie un objet null.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > getPrevious(visibleOnly: bool)|Obtient la feuille de calcul qui précède celle-ci. Si aucune feuille de calcul ne précède celle-ci, cette méthode génère une erreur.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > getPreviousOrNullObject(visibleOnly: bool)|Obtient la feuille de calcul qui précède celle-ci. Si aucune feuille de calcul ne précède celle-ci, cette méthode renvoie un objet null.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Méthode_ > getFirst(visibleOnly: bool)|Obtient la première feuille de calcul dans la collection.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Méthode_ > getLast(visibleOnly: bool)|Obtient la dernière feuille de calcul dans la collection.|1,5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Nouveautés de l’API JavaScript 1.4 pour Excel
Les ajouts apportés aux API JavaScript pour Excel dans l’ensemble de conditions requises 1.4 sont présentés ci-dessous.

### <a name="named-item-add-and-new-properties"></a>Ajout d’élément nommé et nouvelles propriétés

Nouvelles propriétés :

* `comment`
* `scope` éléments inclus dans la feuille de calcul ou dans le classeur
* `worksheet` renvoie la feuille de calcul dans laquelle est inclus l’élément nommé.

Nouvelles méthodes :

* `add(name: string, reference: Range or string, comment: string)`Ajoute un nouveau nom à la collection de l’étendue donnée.
* `addFormulaLocal(name: string, formula: string, comment: string)`Ajoute un nouveau nom à la collection de l’étendue donnée à l’aide des paramètres régionaux de l’utilisateur pour la formule.

### <a name="settings-api-in-the-excel-namespace"></a>API Settings dans l’espace de noms Excel

L’objet [Setting](/javascript/api/excel/excel.setting) représente une paire clé-valeur d’un paramètre conservé dans le document. La fonctionnalité de `Excel.Setting` équivaut à `Office.Settings`, mais utilise la syntaxe d’API par lots plutôt que le modèle de rappel de l’API commune.

Les API comprennent `getItem()` pour obtenir une entrée de paramètre via la clé, et `add()` pour ajouter la paire de paramètres clé/valeur spécifiée dans le classeur.

### <a name="others"></a>Autres

* Définir le nom de colonne du tableau (la version précédente permettait uniquement un accès en lecture seule).
* Ajouter une colonne à la fin du tableau (la version précédente permettait d’ajouter des colonnes partout sauf à la fin).
* Ajouter plusieurs lignes en même temps à un tableau (la version précédente permettait uniquement d’ajouter 1 ligne à la fois).
* `range.getColumnsAfter(count: number)` et `range.getColumnsBefore(count: number)` pour obtenir un certain nombre de colonnes à droite/gauche de l’objet de plage actuel.
* Fonction pour obtenir l’élément ou l’objet null : Cette fonctionnalité permet d’obtenir un objet à l’aide d’une clé. Si l’objet n’existe pas, la propriété isNullObject renvoyée aura la valeur true. Cette fonctionnalité permet aux développeurs de vérifier si un objet existe ou pas sans avoir à le traiter via la gestion des exceptions. Disponible sur une feuille de calcul, un élément nommé, une liaison, une série de graphiques, etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Objet| Quelles sont les nouveautés ?| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Méthode_ > getCount()|Obtient le nombre de liaisons de la collection.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Méthode_ > getItemOrNullObject(id: string)|Obtient un objet de liaison par ID. Si l’objet de liaison n’existe pas, renvoie un objet null.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Méthode_ > getCount()|Renvoie le nombre de graphiques dans la feuille de calcul.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Méthode_ > getItemOrNullObject(name: string)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Méthode_ > getCount()|Renvoie le nombre de points de graphique dans la série.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Méthode_ > getCount()|Renvoie le nombre de séries de la collection.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriété_ > comment|Représente le commentaire associé à ce nom.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriété_ > scope|Indique si le nom est étendu au classeur ou à une feuille de calcul spécifique. En lecture seule. Les valeurs possibles sont les suivantes : Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relation_ > worksheet|Renvoie la feuille de calcul dans laquelle est inclus l’élément nommé. Génère une erreur si les éléments sont inclus dans le classeur à la place. En lecture seule.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relation_ > worksheetOrNullObject|Renvoie la feuille de calcul dans laquelle est inclus l’élément nommé. Renvoie un objet null si l’élément est inclus dans le classeur à la place. En lecture seule.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Méthode_ > delete()|Supprime le nom donné.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Méthode_ > getRangeOrNullObject()|Renvoie l’objet de plage qui est associé au nom. Renvoie un objet null si le type de l’élément nommé n’est pas une plage.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Méthode_ > add(name: string, reference: Range or string, comment: string)|Ajoute un nouveau nom à la collection de l’étendue donnée.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Méthode_ > addFormulaLocal(name: string, formula: string, comment: string)|Ajoute un nouveau nom à la collection de l’étendue donnée à l’aide des paramètres régionaux de l’utilisateur pour la formule.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Méthode_ > getCount()|Obtient le nombre d’éléments nommés dans la collection.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Méthode_ > getItemOrNullObject(name: string)|Obtient un objet nameditem à l’aide de son nom. Si l’objet nameditem n’existe pas, renvoie un objet null.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Méthode_ > getCount()|Obtient le nombre de tableaux croisés dynamiques de la collection.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Méthode_ > getItemOrNullObject(name: string)|Obtient un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > getIntersectionOrNullObject(anotherRange: Range or string)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données. Si aucune intersection n’est trouvée, renvoie un objet Null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > getUsedRangeOrNullObject(valuesOnly: bool)|Renvoie la plage utilisée d’un objet de plage donné. Si aucune cellule n’est utilisée dans la plage, cette fonction renvoie un objet null.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Méthode_ > getCount()|Obtient le nombre d’objets RangeView dans la collection.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriété_ > key|Renvoie la clé qui représente l’id du paramètre. En lecture seule.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriété_ > value|Représente la valeur stockée pour ce paramètre.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Méthode_ > delete()|Supprime le paramètre.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriété_ > items|Collection d’objets setting. En lecture seule.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Méthode_ > add(key: string, value: (any))|Définit ou ajoute le paramètre spécifié dans le classeur.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Méthode_ > getCount()|Obtient le nombre de paramètres dans la collection.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Méthode_ > getItem(key: string)|Obtient une Entrée de paramètre via la clé.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Méthode_ > getItemOrNullObject(key: string)|Obtient une Entrée de paramètre via la clé. Si le paramètre n’existe pas, renvoie un objet null.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relation_ > settings|Obtient l’objet Setting qui représente la liaison qui a déclenché l’événement SettingsChanged.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Méthode_ > getCount()]|Obtient le nombre de tableaux de la collection.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Méthode_ > getItemOrNullObject(key: number or string)|Obtient un tableau à l’aide de son nom ou de son ID. Si le tableau n’existe pas, renvoie un objet null.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Méthode_ > getCount()|Obtient le nombre de colonnes dans le tableau.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Méthode_ > getItemOrNullObject(key: number or string)|Obtient un objet de colonne par nom ou par ID. Si la colonne n’existe pas, renvoie un objet null.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Méthode_ > getCount()|Obtient le nombre de lignes dans le tableau.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > settings|Représente une collection d’objets Settings associés au classeur. En lecture seule.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relation_ > names|Collection de noms inclus dans l’étendue de la feuille de calcul active. En lecture seule.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Méthode_ > getUsedRangeOrNullObject(valuesOnly: bool)|La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté. Si la feuille de calcul entière est vide, cette fonction renvoie un objet null.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Méthode_ > getCount(visibleOnly: bool)|Obtient le nombre de feuilles de calcul dans la collection.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Méthode_ > getItemOrNullObject(key: string)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID. Si la feuille de calcul n’existe pas, renvoie un objet null.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Nouveautés de l’API JavaScript 1.3 pour Excel

Les ajouts apportés aux API JavaScript pour Excel dans l’ensemble de conditions requises 1.3 sont présentés ci-dessous.

|Objet| Nouveautés| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Méthode_ > delete()|Supprime la liaison.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Méthode_ > add(range: Range or string, bindingType: string, id: string)|Ajoute une nouvelle liaison à une plage spécifique.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Méthode_ > addFromNamedItem(name: string, bindingType: string, id: string)|Ajoute une nouvelle liaison basée sur un élément nommé dans le classeur.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Méthode_ > addFromSelection(bindingType: string, id: string)|Ajoute une nouvelle liaison basée sur la sélection en cours.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Méthode_ > getItemOrNull(id: string)|Obtient un objet de liaison par ID. Si l’objet de liaison n’existe pas, la propriété de l’objet renvoyé est null et aura la valeur true.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Méthode_ > getItemOrNull(name: string)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Méthode_ > getItemOrNull(name: string)|Obtient un objet NamedItem à l’aide de son nom. Si l’objet NamedItem n’existe pas, la propriété de l’objet renvoyé est null et aura la valeur true.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Propriété_ > name|Nom du tableau croisé dynamique.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relation_ > worksheet|Feuille de calcul contenant le tableau croisé dynamique. En lecture seule.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Méthode_ > refresh()|Actualise le tableau croisé dynamique.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Propriété_ > items|Collection d’objets de tableau croisé dynamique. En lecture seule.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Méthode_ > getItem(name: string)|Obtient un tableau croisé dynamique par nom.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Méthode_ > getItemOrNull(name: string)|Obtient un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, la propriété de l’objet renvoyé est null et aura la valeur true.|1.3|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > getIntersectionOrNull(anotherRange: Range or string)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données. Si aucune intersection n’est trouvée, renvoie un objet Null.|1.3|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > getVisibleView()|Représente les lignes visibles de la plage en cours.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > cellAddresses|Représente les adresses de cellule de la RangeView. En lecture seule.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > columnCount|Renvoie le nombre de colonnes visibles. En lecture seule.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > formulas|Représente la formule dans le style de notation A1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > formulasLocal|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur.  Par exemple, la formule « =SUM(A1, présentée dans 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > formulasR1C1|Représente la formule dans le style de notation R1C1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > index|Renvoie une valeur qui représente l’index de l’affichage de plage. En lecture seule.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > numberFormat|Représente le code de format de nombre d’Excel pour une cellule donnée.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > rowCount|Renvoie le nombre de lignes visibles. En lecture seule.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > text|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > valueTypes|Représente le type de données de chaque cellule. En lecture seule. Les valeurs possibles sont les suivantes : Unknown (inconnu), Empty (vide), String (chaîne), Integer (entier), Double (double), Boolean (valeur booléenne), Error (erreur).|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriété_ > values|Représente les valeurs brutes de l’affichage de plage spécifié. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie la chaîne d’erreur.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Relation_ > rows|Représente une collection d’affichages de plage associés à la plage. En lecture seule.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Méthode_ > getRange()|Obtient la plage parent associée à l’affichage de plage actuel.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Propriété_ > items|Collection d’objets rangeView. En lecture seule.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Méthode_ > getItemAt(index: number)|Obtient une ligne d’affichage de plage par l’intermédiaire de son index. Avec index de base zéro.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Propriété_ > key|Renvoie la clé qui représente l’id du paramètre. En lecture seule.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Méthode_ > delete()|Supprime le paramètre.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriété_ > items|Collection d’objets setting. En lecture seule.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Méthode_ > getItem(key: string)|Obtient une Entrée de paramètre via la clé.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Méthode_ > getItemOrNull(key: string)|Obtient une Entrée de paramètre via la clé. Si l’objet Paramètre n’existe pas, la propriété de l’objet renvoyé est null et aura la valeur true.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Méthode_ > set(key: string, value: string)|Définit ou ajoute le paramètre spécifié dans le classeur.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relation_ > settingCollection|Obtient l’objet Setting qui représente la liaison qui a déclenché l’événement SettingsChanged.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriété_ > highlightFirstColumn|Indique si la première colonne contient une mise en forme spéciale.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriété_ > highlightLastColumn|Indique si la dernière colonne contient une mise en forme spéciale.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriété_ > showBandedColumns|Indique si les colonnes affichent une mise en forme à bandes dans laquelle la mise en évidence des colonnes impaires diffère de celle des colonnes paires pour faciliter la lecture du tableau.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriété_ > showBandedRows|Indique si les lignes affichent une mise en forme à bandes dans laquelle la mise en évidence des lignes impaires diffère de celle des lignes paires pour faciliter la lecture du tableau.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriété_ > showFilterButton|Indique si les boutons de filtre sont visibles dans la partie supérieure de chaque en-tête de colonne. Ce paramètre est autorisé uniquement si le tableau contient une ligne d’en-tête.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Méthode_ > getItemOrNull(key: number or string)|Obtient un tableau à l’aide de son nom ou de son ID. Si le tableau n’existe pas, la propriété de l’objet renvoyé est null et aura la valeur true.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Méthode_ > getItemOrNull(key: number or string)|Obtient un objet de colonne par son nom ou son ID. Si la colonne n’existe pas, la propriété de l’objet renvoyé est null et aura la valeur true.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > pivotTables|Représente une collection de tableaux croisés dynamiques associés au classeur. En lecture seule.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > settings|Représente une collection d’objets Settings associés au classeur. En lecture seule.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relation_ > pivotTables|Collection de tableaux croisés dynamiques qui font partie de la feuille de calcul. En lecture seule.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Nouveautés de l’API JavaScript 1.2 pour Excel

Les ajouts apportés aux API JavaScript pour Excel dans l’ensemble de conditions requises 1.2 sont présentés ci-dessous.

|Objet| Nouveautés| Description|Ensemble de conditions requises|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriété_ > id|Extrait un graphique en fonction de sa position dans la collection. En lecture seule.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Relation_ > worksheet|Feuille de calcul contenant le graphique actuel. En lecture seule.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Méthode_ > getImage(height: number, width: number, fittingMode: string)|Affiche le graphique sous forme d’image codée en Base64 ajustée aux dimensions spécifiées.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Relation_ > criteria|Le filtre actuellement appliqué à la colonne donnée. En lecture seule.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > apply(criteria: FilterCriteria)|Appliquer les critères de filtre donnés à la colonne indiquée.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyBottomItemsFilter(count: number)|Appliquer un filtre « Élément inférieur » à la colonne pour le nombre d’éléments donné.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ >applyBottomPercentFilter(percent: number)|Appliquer un filtre « Pourcentage inférieur » à la colonne pour le pourcentage d’éléments donné.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyCellColorFilter(color: string)|Appliquer un filtre « Couleur de cellule » à la colonne pour la couleur donnée.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyCustomFilter(criteria1: string, criteria2: string, oper: string)|Appliquer un filtre « Icône » à la colonne pour les chaînes de critères données.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyDynamicFilter(criteria: string)|Appliquer un filtre « Dynamique » à la colonne.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyFontColorFilter(color: string)|Appliquer un filtre « Couleur de police » à la colonne pour la couleur donnée.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyIconFilter(icon: Icon)|Appliquer un filtre « Icône » à la colonne pour l’icône donnée.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyTopItemsFilter(count: number)|Appliquer un filtre « Élément supérieur » à la colonne pour le nombre d’éléments donné.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyTopPercentFilter(percent: number)|Appliquer un filtre « Pourcentage supérieur » à la colonne pour le pourcentage d’éléments donné.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > applyValuesFilter (values : ())|Appliquer un filtre « Valeurs » à la colonne pour les valeurs données.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Méthode_ > clear()|Effacer le filtre sur la colonne donnée.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriété_ > color|Chaîne de couleur HTML utilisée pour filtrer des cellules. Utilisée avec le filtrage « cellColor » et « fontColor ».|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriété_ > criterion1|Premier critère utilisé pour filtrer des données. Utilisé comme opérateur dans le cas d’un filtrage « Custom ».|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriété_ > criterion2|Second critère utilisé pour filtrer des données. Utilisé uniquement comme opérateur dans le cas d’un filtrage « Custom ».|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriété_ > dynamicCriteria|Critères dynamiques de l’ensemble Excel.DynamicFilterCriteria à appliquer à cette colonne. Utilisé avec un filtrage « Dynamic ». Les valeurs possibles sont les suivantes : Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriété_ > filterOn|Propriété utilisée par le filtre pour déterminer si les valeurs doivent rester visibles. Les valeurs possibles sont les suivantes : BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriété_ > operator|Opérateur utilisé pour combiner les critères 1 et 2 lorsque vous utilisez le filtrage « Custom ». Les valeurs possibles sont les suivantes : And, Or.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriété_ > values|Valeurs à utiliser pour le filtrage « Values ».|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Relation_ > icon|Icône utilisée pour filtrer des cellules. Utilisé avec le filtrage « Icon ».|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriété_ > date|Date au format ISO8601 utilisée pour filtrer des données.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriété_ > specificity|Utilisation de la date pour conserver des données. Par exemple, si la date est 2005-04-02 et la spécificité est définie sur « mois », le filtre conservera toutes les lignes dont la date correspond au mois d’avril 2009. Les valeurs possibles sont les suivantes : Year (année), Monday (lundi), Day (jour), Hour (heure), Minute (minute), Second (seconde).|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriété_ > formulaHidden|Indique si Excel masque la formule des cellules dans la plage. Une valeur null indique que les paramètres de formule masquée ne sont pas les mêmes sur l’ensemble de la plage.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriété_ > locked|Indique si Excel verrouille les cellules dans l’objet. Une valeur null indique que les paramètres de verrouillage ne sont pas les mêmes sur l’ensemble de la plage.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriété_ > index|Représente l’index de l’icône dans l’ensemble donné.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriété_ > set|Représente l’ensemble dont fait partie l’icône. Les valeurs possibles sont les suivantes : Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > columnHidden|Indique si toutes les colonnes de la plage active sont masquées.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > formulasR1C1|Représente la formule dans le style de notation R1C1.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > hidden|Indique si toutes les cellules de la plage active sont masquées. En lecture seule.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriété_ > rowHidden|Indique si toutes les lignes de la plage active sont masquées.|1.2|
|[range](/javascript/api/excel/excel.range)|_Relation_ > sort|Représente le tri de plage de la plage actuelle. En lecture seule.|1.2|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > merge(across: bool)|Fusionne la plage de cellules dans une zone de la feuille de calcul.|1.2|
|[range](/javascript/api/excel/excel.range)|_Méthode_ > unmerge()|Annule la fusion de la plage de cellules et les sépare dans des cellules distinctes.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriété_ > columnWidth|Obtient ou définit la largeur de toutes les colonnes de la plage. Si les largeurs de colonne ne sont pas uniformes, la valeur « null » est renvoyée.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriété_ > rowHeight|Obtient ou définit la hauteur de toutes les lignes de la plage. Si les hauteurs de lignes ne sont pas uniformes, la valeur « null » est renvoyée.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Relation_ > protection|Renvoie l’objet de protection du format pour une plage. En lecture seule.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Méthode_ > autofitColumns()|Modifie la largeur des colonnes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Méthode_ > autoFitRows()|Modifie la hauteur des lignes de la plage active pour obtenir le meilleur ajustement, en fonction des données présentes dans les colonnes.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Propriété_ > address|Représente les lignes visibles de la plage en cours.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Méthode_ > apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|Effectue une opération de tri.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriété_ > ascending|Indique si le tri s’effectue dans l’ordre croissant.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriété_ > color|Couleur ciblée par la condition si le tri est appliqué à la couleur ou à la police de la cellule.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriété_ > dataOption|Options de tri supplémentaires pour ce champ. Les valeurs possibles sont les suivantes : Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriété_ > key|Colonne (ou ligne, selon l’orientation du tri) ciblée par la condition. Représentée sous forme d’un décalage par rapport à la première colonne (ou ligne).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriété_ > sortOn|Type de tri de cette condition. Les valeurs possibles sont les suivantes : Value, CellColor, FontColor, Icon.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Relation_ > icon|Représente l’icône ciblée par la condition si le tri est appliqué à l’icône de la cellule.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relation_ > sort|Représente le tri du tableau. En lecture seule.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relation_ > worksheet|Feuille de calcul contenant le tableau actif. En lecture seule.|1.2|
|[table](/javascript/api/excel/excel.table)|_Méthode_ > clearFilters()|Supprime tous les filtres appliqués actuellement sur le tableau.|1.2|
|[table](/javascript/api/excel/excel.table)|_Méthode_ > convertToRange()|Convertit le tableau en plage normale de cellules. Toutes les données sont conservées.|1.2|
|[table](/javascript/api/excel/excel.table)|_Méthode_ > reapplyFilters()|Applique de nouveau tous les filtres actuellement appliqués sur le tableau.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Relation_ > filter|Extrait le filtre appliqué à la colonne. En lecture seule.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriété_ > matchCase|Indique si la casse a influé sur le dernier tri du tableau. En lecture seule.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriété_ > method|Dernière méthode de classement des caractères chinois utilisée pour trier le tableau. En lecture seule. Les valeurs possibles sont les suivantes : PinYin, StrokeCount|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Relation_ > fields|Dernières conditions utilisées pour trier le tableau. En lecture seule.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Méthode_ > apply(fields: SortField, matchCase: bool, method: string)|Effectue une opération de tri.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Méthode_ > clear()|Efface le tri actuellement appliqué au tableau. Même si le classement du tableau n’est pas modifié, l’état des boutons d’en-tête est rétabli.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Méthode_ > reapply()|Applique à nouveau les paramètres actuels de tri au tableau.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Relation_ > functions|Représente l’instance de l’application Excel contenant ce classeur. En lecture seule.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relation_ > protection|Renvoie un objet de protection de feuille pour une feuille de calcul. En lecture seule.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Propriété_ > protected|Indique si la feuille de calcul est protégée. En lecture seule. En lecture seule.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Relation_ > options|Options de protection de feuille. En lecture seule.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Méthode_ > protect(options: WorksheetProtectionOptions)|Protège une feuille de calcul. Échoue si la feuille de calcul est protégée.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Méthode_ > unprotect()|Annule la protection d’une feuille de calcul.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowAutoFilter|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Filtre automatique.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowDeleteColumns|Représente l’option de protection de feuille de calcul qui autorise la suppression des colonnes.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowDeleteRows|Représente l’option de protection de feuille de calcul qui autorise la suppression des lignes.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowFormatCells|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des cellules.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowFormatColumns|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des colonnes.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowFormatRows|Représente l’option de protection de feuille de calcul qui autorise la mise en forme des lignes.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowInsertColumns|Représente l’option de protection de feuille de calcul qui autorise l’insertion des colonnes.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowInsertHyperlinks|Représente l’option de protection de feuille de calcul qui autorise l’insertion des liens hypertexte.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowInsertRows|Représente l’option de protection de feuille de calcul qui autorise l’insertion des lignes.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowPivotTables|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tableau croisé dynamique.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriété_ > allowSort|Représente l’option de protection de feuille de calcul qui autorise l’utilisation de la fonctionnalité Tri.|1.2|

## <a name="excel-javascript-api-11"></a>API JavaScript 1.1 pour Excel

L’API JavaScript 1.1 pour Excel est la première version de l’API. Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[API JavaScript pour Excel](/javascript/api/excel).

## <a name="see-also"></a>Voir aussi

- [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Spécification des exigences en matière d’hôtes Office et d’API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifeste XML des compléments Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
