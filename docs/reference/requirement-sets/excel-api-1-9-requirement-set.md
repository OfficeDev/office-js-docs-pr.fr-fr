---
title: Excel Ensemble de conditions requises de l’API JavaScript 1.9
description: Détails sur l’ensemble de conditions requises ExcelApi 1.9.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: dde36db799a7f0612439e934d50af4f3ab04077e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152243"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Nouveautés de l Excel API JavaScript 1.9

Plus de 500 nouvelles API Excel ont été ajoutés avec l’ensemble de conditions requises 1.9. Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | Insertion, la position et format images, formes géométriques et zones de texte. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Filtre automatique](../../excel/excel-add-ins-worksheets.md#filter-data) | Ajouter des filtres à des plages. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Zones](../../excel/excel-add-ins-multiple-ranges.md) | Prise en charge de plages discontinues. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Cellules spéciales](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Obtenez les cellules contenant des dates, des commentaires ou des formules dans une plage. | [Plage](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Chercher](../../excel/excel-add-ins-ranges-string-match.md) | Recherchez des valeurs ou des formules dans une plage ou une feuille de calcul. | [Plage](/javascript/api/excel/excel.range#find-text--criteria-)[feuille de calcul](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copier et coller](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Copier des formules, formats et valeurs d’une plage à l’autre. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calcul](../../excel/performance.md#suspend-calculation-temporarily) | Contrôle plus étroit sur le moteur de calcul Excel. | [Application](/javascript/api/excel/excel.application) |
| Nouveaux graphiques | Explorez nos nouveaux types de graphiques pris en charge : cartes, zone et valeur, en cascade, en rayons de soleil, pareto. et entonnoir. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | Nouvelles fonctionnalités avec les formats de plage. | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de Excel l’ensemble de conditions requises de l’API JavaScript 1.9. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.9 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationEngineVersion)|Renvoie la version du moteur de calcul Excel utilisée pour le dernier recalcul complet.|
||[calculationState](/javascript/api/excel/excel.application#calculationState)|Renvoie l’état de calcul de l’application.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativeCalculation)|Renvoie les paramètres de calcul itératifs.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendScreenUpdatingUntilNextSync__)|Suspend la mise à jour de l’écran jusqu’à ce que le `context.sync()` suivant soit appelé.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[Appliquer (plage : plage \| chaîne, columnIndex ? : nombre, critères ? : Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply_range__columnIndex__criteria_)|Applique le filtre automatique à une plage.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearCriteria__)|Efface les critères de filtre du filtre automatique.|
||[getRange()](/javascript/api/excel/excel.autofilter#getRange__)|Renvoie `Range` l’objet qui représente la plage à laquelle le filtre automatique s’applique.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getRangeOrNullObject__)|Renvoie `Range` l’objet qui représente la plage à laquelle le filtre automatique s’applique.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Tableau qui conserve tous les critères de filtre dans une plage filtrée.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Spécifie si le filtre automatique est activé.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isDataFiltered)|Spécifie si le filtre automatique a des critères de filtre.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply__)|Applique l’objet Autofilter spécifié actuellement sur la plage.|
||[remove()](/javascript/api/excel/excel.autofilter#remove__)|Supprime le filtre automatique pour la plage.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Représente la `color` propriété d’une bordure simple.|
||[style](/javascript/api/excel/excel.cellborder#style)|Représente la `style` propriété d’une bordure simple.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintAndShade)|Représente la `tintAndShade` propriété d’une bordure simple.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Représente la `weight` propriété d’une bordure simple.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bas](/javascript/api/excel/excel.cellbordercollection#bottom)|Représente la `format.borders.bottom` propriété.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonalDown)|Représente la `format.borders.diagonalDown` propriété.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalUp)|Représente la `format.borders.diagonalUp` propriété.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Représente la `format.borders.horizontal` propriété.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Représente la `format.borders.left` propriété.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Représente la `format.borders.right` propriété.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Représente la `format.borders.top` propriété.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Représente la `format.borders.vertical` propriété.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[adresse](/javascript/api/excel/excel.cellproperties#address)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addressLocal)|Représente la `addressLocal` propriété.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Représente la `hidden` propriété.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Représente la `format.fill.color` propriété.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Représente la `format.fill.pattern` propriété.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patternColor)|Représente la `format.fill.patternColor` propriété.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patternTintAndShade)|Représente la `format.fill.patternTintAndShade` propriété.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintAndShade)|Représente la `format.fill.tintAndShade` propriété.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Représente la `format.font.bold` propriété.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Représente la `format.font.color` propriété.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Représente la `format.font.italic` propriété.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Représente la `format.font.name` propriété.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Représente la `format.font.size` propriété.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Représente la `format.font.strikethrough` propriété.|
||[Subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Représente la `format.font.subscript` propriété.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Représente la `format.font.superscript` propriété.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintAndShade)|Représente la `format.font.tintAndShade` propriété.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Représente la `format.font.underline` propriété.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoIndent)|Représente la `autoIndent` propriété.|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Représente la `borders` propriété.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Représente la `fill` propriété.|
||[police](/javascript/api/excel/excel.cellpropertiesformat#font)|Représente la `font` propriété.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalAlignment)|Représente la `horizontalAlignment` propriété.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentLevel)|Représente la `indentLevel` propriété.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Représente la `protection` propriété.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingOrder)|Représente la `readingOrder` propriété.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinkToFit)|Représente la `shrinkToFit` propriété.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textOrientation)|Représente la `textOrientation` propriété.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)|Représente la `useStandardHeight` propriété.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth)|Représente la `useStandardWidth` propriété.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalAlignment)|Représente la `verticalAlignment` propriété.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wrapText)|Représente la `wrapText` propriété.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulaHidden)|Représente la `format.protection.formulaHidden` propriété.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Représente la `format.protection.locked` propriété.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueAfter)|Représente la valeur après la modification.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valueBefore)|Représente la valeur avant la modification.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valueTypeAfter)|Représente le type de valeur après la modification.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valueTypeBefore)|Représente le type de valeur avant la modification.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate__)|Active le graphique dans l’interface utilisateur Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotOptions)|Encapsule les options pour le graphique croisé dynamique.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorScheme)|Spécifie le modèle de couleurs du graphique.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedCorners)|Spécifie si la zone de graphique du graphique possède des coins arrondis.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linkNumberFormat)|Spécifie si le format numérique est lié aux cellules.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowOverflow)|Spécifie si le débordement bin est activé dans un histogramme ou un graphique de pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowUnderflow)|Spécifie si le sous-flux bin est activé dans un histogramme ou un graphique de pareto.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Spécifie le nombre de bacs d’un histogramme ou d’un graphique de pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowValue)|Spécifie la valeur de débordement bin d’un histogramme ou d’un graphique de pareto.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Spécifie le type du bac pour un histogramme ou un graphique de pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowValue)|Spécifie la valeur de sous-flux bin d’un histogramme ou d’un graphique de pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Spécifie la valeur de largeur bin d’un histogramme ou d’un graphique de pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartileCalculation)|Spécifie si le type de calcul du quartile d’un graphique de zone et de zone.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showInnerPoints)|Spécifie si les points internes sont affichés dans une zone et un graphique de zone.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanLine)|Spécifie si la ligne moyenne est affichée dans une zone et un graphique de zone.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanMarker)|Spécifie si la marque moyenne est affichée dans une zone et un graphique de zone.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showOutlierPoints)|Spécifie si les points aberrants sont affichés dans un graphique zone et valeur.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linkNumberFormat)|Spécifie si le format numérique est lié aux cellules (de sorte que le format numérique change dans les étiquettes lorsqu’il change dans les cellules).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linkNumberFormat)|Spécifie si le format numérique est lié aux cellules.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endStyleCap)|Spécifie si les barres d’erreur ont un style de fin.|
||[inclure](/javascript/api/excel/excel.charterrorbars#include)|Spécifie les parties de barres d’erreur à inclure.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Spécifie le type de mise en forme de barres d’erreur.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Le type de plage marqué par des barres d’erreur.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Spécifie si les barres d’erreur sont affichées.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Représente le format des lignes du graphique.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelStrategy)|Spécifie la stratégie d’étiquettes de carte de série d’un graphique de carte région.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Spécifie le niveau de mappage des séries d’un graphique de carte région.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectionType)|Spécifie le type de projection de série d’un graphique région carte.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showAxisFieldButtons)|Spécifie s’il faut afficher les boutons de champ d’axe sur une PivotChart.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showLegendFieldButtons)|Spécifie s’il faut afficher les boutons de champ de légende sur un PivotChart.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showReportFilterFieldButtons)|Spécifie s’il faut afficher les boutons de champ de filtre de rapport sur une PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showValueFieldButtons)|Spécifie s’il faut afficher les boutons de champ afficher la valeur sur une PivotChart.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubbleScale)|Peut être une valeur d’entier entre 0 (zéro) et 300 correspondant à un pourcentage de la taille par défaut.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientMaximumColor)|Spécifie la couleur de la valeur maximale d’une série de graphique région carte.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientMaximumType)|Spécifie le type de valeur maximale d’une série de graphique région carte.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientMaximumValue)|Spécifie la valeur maximale d’une série de graphique région carte.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientMidpointColor)|Spécifie la couleur de la valeur du milieu d’une série de graphique région carte.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientMidpointType)|Spécifie le type de la valeur de milieu d’une série de graphique région carte.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientMidpointValue)|Spécifie la valeur du milieu d’une série de graphique région carte.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientMinimumColor)|Spécifie la couleur de la valeur minimale d’une série de graphique région carte.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientMinimumType)|Spécifie le type de la valeur minimale d’une série de graphique région carte.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientMinimumValue)|Spécifie la valeur minimale d’une série de graphique région carte.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientStyle)|Spécifie le style de dégradé de série d’un graphique région carte.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertColor)|Spécifie la couleur de remplissage des points de données négatifs d’une série.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentLabelStrategy)|Spécifie la zone de stratégie des étiquettes parentes de série pour un graphique en arborescence.|
||[binOptions](/javascript/api/excel/excel.chartseries#binOptions)|Encapsule les options bin uniquement pour les histogrammes et graphiques de pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskerOptions)|Résume les options pour le graphique de zone et valeur.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapOptions)|Encapsule les options pour le graphique carte de région.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xErrorBars)|Représente l’objet de la barre d’erreur pour une série de graphique.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yErrorBars)|Représente l’objet de la barre d’erreur pour une série de graphique.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showConnectorLines)|Spécifie si les lignes de connecteur sont affichées dans les graphiques en cascade.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showLeaderLines)|Spécifie si les lignes d’étiquettes sont affichées pour chaque étiquette de données de la série.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitValue)|Spécifie la valeur de seuil qui sépare deux sections d’un graphique en secteurs de secteur ou d’un graphique en barres de secteur.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linkNumberFormat)|Spécifie si le format numérique est lié aux cellules (de sorte que le format numérique change dans les étiquettes lorsqu’il change dans les cellules).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[adresse](/javascript/api/excel/excel.columnproperties#address)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addressLocal)|Représente la `addressLocal` propriété.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnIndex)|Représente la `columnIndex` propriété.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getRanges__)|Renvoie la , comprenant une ou plusieurs plages rectangulaires, à laquelle le format conditionnel `RangeAreas` est appliqué.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getInvalidCells__)|Renvoie un `RangeAreas` objet, comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valides.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getInvalidCellsOrNullObject__)|Renvoie un `RangeAreas` objet, comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valides.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subField)|Propriété utilisée par le filtre pour faire un filtre enrichi sur les valeurs enrichies.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Représente l’identificateur de forme.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Renvoie `Shape` l’objet de la forme géométrique.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getCount__)|Renvoie le nombre de formes dans le groupe de la forme.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getItem_key_)|Obtient une forme à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getItemAt_index_)|Obtient une forme en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerFooter)|Pied de la feuille de calcul.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerHeader)|En-tête central de la feuille de calcul.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftFooter)|Pied de la feuille de calcul gauche.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftHeader)|En-tête gauche de la feuille de calcul.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightFooter)|Pied de la feuille de calcul à droite.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightHeader)|En-tête droit de la feuille de calcul.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultForAllPages)|L’en-tête/pied de page, utilisé pour toutes les pages, sauf si la première page ou page impaire/paire est spécifiée.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenPages)|L’en-tête/le pied de page à utiliser pour les pages paires, en-tête/pied de page impaire doit être spécifié pour les pages impaires.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstPage)|La première en-tête/le premier pied de page, pour toutes les autres pages générales ou impair/pair est utilisé.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddPages)|L’en-tête/le pied de page à utiliser pour les pages paires, l’en-tête/pied de page paire doit être spécifié pour les pages paires.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|État selon lequel les en-têtes/pieds de groupe sont définies.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#useSheetMargins)|Obtient ou définit un indicateur indiquant si les en-têtes/pieds de page sont alignés avec les marges de page définis dans les options de mise en page pour la feuille de calcul.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#useSheetScale)|Obtient ou définit un indicateur indiquant si les en-têtes/pieds de page sont alignés avec les marges de page définis dans les options de mise en page pour la feuille de calcul.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Renvoie le format de l’image.|
||[id](/javascript/api/excel/excel.image#id)|Spécifie l’identificateur de forme de l’objet image.|
||[shape](/javascript/api/excel/excel.image#shape)|Renvoie `Shape` l’objet associé à l’image.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Cette propriété a la valeur True si Microsoft Excel utilise l'itération pour résoudre des références circulaires.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxChange)|Spécifie la quantité maximale de modification entre chaque itération à mesure Excel des références circulaires.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxIteration)|Spécifie le nombre maximal d’itérations que Excel pouvez utiliser pour résoudre une référence circulaire.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginArrowheadLength)|Renvoie ou définit la longueur de la pointe de la flèche au début de la ligne spécifiée.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginArrowheadStyle)|Représente le style de la pointe de la flèche au début de la ligne spécifiée.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginArrowheadWidth)|Représente la largeur de la pointe de la flèche au début de la ligne spécifiée.|
||[connectBeginShape (forme : Excel.Shape, connectionSite : nombre)](/javascript/api/excel/excel.line#connectBeginShape_shape__connectionSite_)|Détache le début du connecteur spécifié de la forme à laquelle il est attaché.|
||[connectBeginShape (forme : Excel.Shape, connectionSite : nombre)](/javascript/api/excel/excel.line#connectEndShape_shape__connectionSite_)|Joint la fin du connecteur spécifié à une forme spécifiée.|
||[connectorType](/javascript/api/excel/excel.line#connectorType)|Représente le type de connecteur pour la ligne.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectBeginShape__)|Détache le début du connecteur spécifié de la forme à laquelle il est attaché.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectEndShape__)|Détache la fin du connecteur spécifié de la forme à laquelle il est attaché.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endArrowheadLength)|Représente la longueur de la pointe de la flèche à la fin de la ligne spécifiée.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endArrowheadStyle)|Représente le style de la pointe de la flèche à la fin de ligne spécifée.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endArrowheadWidth)|Représente la largeur de la pointe de la flèche à la fin de la ligne spécifiée.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginConnectedShape)|Représente la forme de la pointe de la flèche au début de la ligne spécifiée.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginConnectedSite)|Représente le site de connexion indiquant le point de connexion auquel le début d'un connecteur est relié.|
||[endConnectedShape](/javascript/api/excel/excel.line#endConnectedShape)|Représente la forme de la pointe de la flèche à la fin de la ligne spécifiée.|
||[endConnectedSite](/javascript/api/excel/excel.line#endConnectedSite)|Représente le site de connexion indiquant le point de connexion auquel la fin d'un connecteur est relié.|
||[id](/javascript/api/excel/excel.line#id)|Spécifie l’identificateur de forme.|
||[isBeginConnected](/javascript/api/excel/excel.line#isBeginConnected)|Spécifie si le début de la ligne spécifiée est connecté à une forme.|
||[isEndConnected](/javascript/api/excel/excel.line#isEndConnected)|Spécifie si la fin de la ligne spécifiée est connectée à une forme.|
||[shape](/javascript/api/excel/excel.line#shape)|Renvoie `Shape` l’objet associé à la ligne.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete__)|Supprime un objet de saut de page.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getCellAfterBreak__)|Obtient la première cellule après le saut de page.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnIndex)|Spécifie l’index de colonne pour le pause de page.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowIndex)|Spécifie l’index de ligne pour le pause de page.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[Ajouter (pageBreakRange : plage \| chaîne)](/javascript/api/excel/excel.pagebreakcollection#add_pageBreakRange_)|Ajoute un saut de page avant la cellule en haut à gauche de la plage spécifiée.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getCount__)|Obtient le nombre de pages de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.pagebreakcollection#getItem_index_)|Obtient un objet de saut de page via l’index.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removePageBreaks__)|Redéfinit tous les sauts de page de la collection.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackAndWhite)|Option d’impression noir et blanc de la feuille de calcul.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottomMargin)|Marge de page inférieure de la feuille de calcul à utiliser pour l’impression en points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerHorizontally)|Indicateur horizontal du centre de la feuille de calcul.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centerVertically)|Indicateur vertical du centre de la feuille de calcul.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftMode)|Option de mode brouillon de la feuille de calcul.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstPageNumber)|Premier numéro de page de la feuille de calcul à imprimer.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footerMargin)|Marge de pied de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getPrintArea__)|Obtient l’objet, comprenant une ou plusieurs plages rectangulaires, qui représente la zone `RangeAreas` d’impression de la feuille de calcul.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintAreaOrNullObject__)|Obtient l’objet, comprenant une ou plusieurs plages rectangulaires, qui représente la zone `RangeAreas` d’impression de la feuille de calcul.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumns__)|Obtient l’objet plage représentant les colonnes de titre.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumnsOrNullObject__)|Obtient l’objet plage représentant les colonnes de titre.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getPrintTitleRows__)|Obtient l’objet plage représentant les rangées de titre.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleRowsOrNullObject__)|Obtient l’objet plage représentant les rangées de titre.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headerMargin)|Marge d’en-tête de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftMargin)|Marge gauche de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[Orientation](/javascript/api/excel/excel.pagelayout#orientation)|Orientation de la feuille de calcul de la page.|
||[paperSize](/javascript/api/excel/excel.pagelayout#paperSize)|Taille de la feuille de calcul de la page.|
||[printComments](/javascript/api/excel/excel.pagelayout#printComments)|Spécifie si les commentaires de la feuille de calcul doivent être affichés lors de l’impression.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printErrors)|Option d’erreurs d’impression de la feuille de calcul.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printGridlines)|Spécifie si le quadrillage de la feuille de calcul sera imprimé.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printHeadings)|Spécifie si les en-tête de la feuille de calcul seront imprimés.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printOrder)|Option d’ordre d’impression de page de la feuille de calcul.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersFooters)|Configuration de l’en-tête et pied de page de la feuille de calcul.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightMargin)|Marge droite de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[setPrintArea (printArea : plage \| RangeAreas \| chaîne)](/javascript/api/excel/excel.pagelayout#setPrintArea_printArea_)|Définit la zone d’impression de la feuille de calcul.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setPrintMargins_unit__marginOptions_)|Définit les marges de page de la feuille de calcul avec des unités.|
||[setPrintTitleColumns (printTitleColumns : plage \| chaîne)](/javascript/api/excel/excel.pagelayout#setPrintTitleColumns_printTitleColumns_)|Définit les colonnes qui contiennent des cellules répétées à gauche de chaque page de la feuille de calcul pour l’impression.|
||[setPrintTitleRows (printTitleRows : plage \| chaîne)](/javascript/api/excel/excel.pagelayout#setPrintTitleRows_printTitleRows_)|Définit les rangées qui contiennent des cellules répétées en haut de chaque page de la feuille de calcul pour l’impression.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topMargin)|Marge supérieure de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Options de zoom avant impression de la feuille de calcul.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bas](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Spécifie la marge inférieure de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Spécifie la marge de pied de page de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Spécifie la marge d’en-tête de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Spécifie la marge gauche de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Spécifie la marge droite de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Spécifie la marge supérieure de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalFitToPages)|Nombre de pages pour l’ajuster horizontalement.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|La valeur d’échelle de page d’impression peut être comprise entre 10 et 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalFitToPages)|Nombre de pages pour l’ajuster verticalement.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortByValues_sortBy__valuesHierarchy__pivotItemScope_)|Trie le PivotField par valeurs spécifiées dans une étendue donnée.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoFormat)|Spécifie si la mise en forme sera automatiquement mise en forme lorsqu’elle est actualisée ou lorsque les champs sont déplacés.|
||[getDataHierarchy (cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#getDataHierarchy_cell_)|Obtient DataHierarchy servant à calculer la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[getPivotItems (axe : Excel.PivotAxis, cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#getPivotItems_axis__cell_)|Obtient le PivotItems à partir d’un axe qui composent la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveFormatting)|Spécifie si la mise en forme est conservée lorsque le rapport est actualisé ou recalculé par des opérations telles que la pivotation, le tri ou la modification d’éléments de champ de page.|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setAutoSortOnCell_cell__sortBy_)|Définit le tableau croisé dynamique pour trier automatiquement à l’aide de la cellule spécifiée pour sélectionner automatiquement tous les critères et contexte nécessaires.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enableDataValueEditing)|Spécifie si le tableau croisé dynamique permet à l’utilisateur de modifier des valeurs dans le corps des données.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#useCustomSortLists)|Spécifie si le tableau croisé dynamique utilise des listes personnalisées lors du tri.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel. AutoFillType)](/javascript/api/excel/excel.range#autoFill_destinationRange__autoFillType_)|Remplit les plages de la plage actuelle à la plage de destination à l’aide de la logique de remplissage automatique spécifiée.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertDataTypeToText__)|Convertit en texte les cellules de plage avec des types de données.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#convertToLinkedDataType_serviceID__languageCulture_)|Convertit les cellules de la plage en types de données liés dans la feuille de calcul.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Copie les données de cellule ou la mise en forme de la plage source ou `RangeAreas` de la plage actuelle.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find_text__criteria_)|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findOrNullObject_text__criteria_)|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[flashFill()](/javascript/api/excel/excel.range#flashFill__)|Fait un remplissage flash à la plage actuelle.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getCellProperties_cellPropertiesLoadOptions_)|Renvoie une plage en 2D, qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getColumnProperties_columnPropertiesLoadOptions_)|Renvoie une plage à dimension unique, qui comprend les données de char colonne de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getRowProperties_rowPropertiesLoadOptions_)|Renvoie une plage à dimension unique , qui comprend les données de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCells_cellType__cellValueType_)|Obtient l’objet, comprenant une ou plusieurs plages rectangulaires, qui représente toutes les cellules qui correspondent au type et à la valeur `RangeAreas` spécifiés.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCellsOrNullObject_cellType__cellValueType_)|Obtient l’objet, comprenant une ou plusieurs plages, qui représente toutes les cellules qui correspondent au type et à la valeur `RangeAreas` spécifiés.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getTables_fullyContained_)|Obtient une collection de tableaux qui se chevauchent avec la plage dans l’étendue.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkedDataTypeState)|Représente l’état du type de données de chaque cellule.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_)|Supprime les valeurs dupliquées de la plage spécifiée par les colonnes.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceAll_text__replacement__criteria_)|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setCellProperties_cellPropertiesData_)|Met à jour la plage basée sur un tableau 2D de propriétés de cellule, en encapsulant des éléments tels que la police, le remplissage, les bordures et l’alignement.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setColumnProperties_columnPropertiesData_)|Met à jour la plage basée sur un tableau à une dimension des propriétés de colonne, en encapsulant des éléments tels que la police, le remplissage, les bordures et l’alignement.|
||[setDirty()](/javascript/api/excel/excel.range#setDirty__)|Cette méthode désigne une plage qui doit être recalculée lorsque le recalcul suivant se produit.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setRowProperties_rowPropertiesData_)|Met à jour la plage basée sur un tableau à une dimension de propriétés de ligne, en encapsulant des éléments tels que la police, le remplissage, les bordures et l’alignement.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate__)|Calcule toutes les cellules dans `RangeAreas` le .|
||[Effacer (applyTo ? : Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear_applyTo_)|Cette propriété permet d’effacer les valeurs, le format, le remplissage, la bordure et d’autres propriétés de chacune des zones qui composent cet `RangeAreas` objet.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertDataTypeToText__)|Convertit toutes les cellules des `RangeAreas` types de données en texte.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#convertToLinkedDataType_serviceID__languageCulture_)|Convertit toutes les cellules de l’ensemble `RangeAreas` en types de données liés.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Copie les données de cellule ou la mise en forme de la plage source ou `RangeAreas` de la plage `RangeAreas` actuelle.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getEntireColumn__)|Renvoie un objet qui représente les colonnes entières de l'(par exemple, si le courant représente les cellules `RangeAreas` « B4:E11, H2 », il renvoie un qui représente les `RangeAreas` `RangeAreas` `RangeAreas` colonnes « B:E, H:H »).|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getEntireRow__)|Renvoie un objet qui représente les lignes entières de l'(par exemple, si le courant représente les cellules « B4:E11 », il renvoie un qui représente les lignes `RangeAreas` `RangeAreas` « `RangeAreas` `RangeAreas` 4:11 »).|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersection_anotherRange_)|Renvoie `RangeAreas` l’objet qui représente l’intersection des plages données ou `RangeAreas` .|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersectionOrNullObject_anotherRange_)|Renvoie `RangeAreas` l’objet qui représente l’intersection des plages données ou `RangeAreas` .|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getOffsetRangeAreas_rowOffset__columnOffset_)|Renvoie un objet décalé par le décalage de ligne `RangeAreas` et de colonne spécifique.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCells_cellType__cellValueType_)|Renvoie un objet qui représente toutes les cellules qui correspondent au type et à la valeur `RangeAreas` spécifiés.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCellsOrNullObject_cellType__cellValueType_)|Renvoie un objet qui représente toutes les cellules qui correspondent au type et à la valeur `RangeAreas` spécifiés.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#getTables_fullyContained_)|Renvoie une collection étendue de tableaux qui chevauchent n’importe quelle plage de cet `RangeAreas` objet.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreas_valuesOnly_)|Renvoie l’objet utilisé qui comprend toutes les zones utilisées `RangeAreas` de plages rectangulaires individuelles dans l’objet. `RangeAreas`|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreasOrNullObject_valuesOnly_)|Renvoie l’objet utilisé qui comprend toutes les zones utilisées `RangeAreas` de plages rectangulaires individuelles dans l’objet. `RangeAreas`|
||[adresse](/javascript/api/excel/excel.rangeareas#address)|Renvoie la `RangeAreas` référence en style A1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addressLocal)|Renvoie la `RangeAreas` référence dans les paramètres régionaux de l’utilisateur.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areaCount)|Renvoie le nombre de plages rectangulaires qui composent cet `RangeAreas` objet.|
||[Zones](/javascript/api/excel/excel.rangeareas#areas)|Renvoie une collection de plages rectangulaires qui composent cet `RangeAreas` objet.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellCount)|Renvoie le nombre de cellules dans l’objet, récapitulant le nombre de cellules de toutes les `RangeAreas` plages rectangulaires individuelles.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalFormats)|Renvoie une collection de formats conditionnels qui se coupent avec les cellules de cet `RangeAreas` objet.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#dataValidation)|Renvoie un objet de validation de données pour toutes les plages dans `RangeAreas` le .|
||[format](/javascript/api/excel/excel.rangeareas#format)|Renvoie un objet, qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de toutes les `RangeFormat` plages de `RangeAreas` l’objet.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isEntireColumn)|Spécifie si toutes les plages de cet objet représentent des colonnes entières `RangeAreas` (par exemple, « A:C, Q:Z »).|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isEntireRow)|Spécifie si toutes les plages de cet objet représentent des lignes `RangeAreas` entières (par exemple, « 1:3, 5:7 »).|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Renvoie la feuille de calcul pour l’actuel `RangeAreas` .|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setDirty__)|Définit les `RangeAreas` données à recalculer lors du recalcul suivant.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Représente le style de toutes les plages de cet `RangeAreas` objet.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintAndShade)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour la bordure de plage, la valeur est entre -1 (plus sombre) et 1 (plus clair), avec 0 pour la couleur d’origine.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintAndShade)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour les bordures de plage.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getCount__)|Renvoie le nombre de plages dans `RangeCollection` le .|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getItemAt_index_)|Renvoie l’objet de plage en fonction de sa position dans `RangeCollection` le .|
||[items](/javascript/api/excel/excel.rangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Motif d’une plage.|
||[patternColor](/javascript/api/excel/excel.rangefill#patternColor)|Code couleur HTML représentant la couleur du modèle de plage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patternTintAndShade)|Spécifie un double qui s’éclaircit ou assombrit une couleur de motif pour le remplissage de la plage.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintAndShade)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour le remplissage de la plage.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Spécifie l’état de la police de type strikethrough.|
||[Subscript](/javascript/api/excel/excel.rangefont#subscript)|Spécifie l’état d’indice de la police.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Spécifie l’état d’exposant de la police.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintAndShade)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour la police de plage.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoIndent)|Spécifie si le texte est automatiquement mis en retrait lorsque l’alignement du texte est égal à la distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentLevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingOrder)|L’ordre de lecture de la plage.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinkToFit)|Spécifie si le texte est automatiquement réduit pour tenir dans la largeur de colonne disponible.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Nombre de lignes dupliquées supprimées par l’opération.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueRemaining)|Nombre de lignes uniques restantes présents dans la plage qui en résulte.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completeMatch)|Spécifie si la correspondance doit être complète ou partielle.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchCase)|Spécifie si la correspondance est sensible à la cas.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[adresse](/javascript/api/excel/excel.rowproperties#address)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addressLocal)|Représente la `addressLocal` propriété.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowIndex)|Représente la `rowIndex` propriété.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completeMatch)|Spécifie si la correspondance doit être complète ou partielle.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchCase)|Spécifie si la correspondance est sensible à la cas.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchDirection)|Détermine le sens de la recherche.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Représente la `format` propriété.|
||[lien hypertexte](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Représente la `hyperlink` propriété.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Représente la `style` propriété.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnHidden)|Représente la `columnHidden` propriété.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnWidth)||
||[format : Excel. CellPropertiesFormat & {
            columnWidth?] (/javascript/api/excel/excel.settablecolumnproperties#format)|Représente la `format` propriété.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format : Excel. CellPropertiesFormat & {
            rowHeight?] (/javascript/api/excel/excel.settablerowproperties#format)|Représente la `format` propriété.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowHeight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowHidden)|Représente la `rowHidden` propriété.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#altTextDescription)|Spécifie l’autre texte de description d’un `Shape` objet.|
||[altTextTitle](/javascript/api/excel/excel.shape#altTextTitle)|Spécifie le texte de titre de remplacement `Shape` d’un objet.|
||[delete()](/javascript/api/excel/excel.shape#delete__)|Supprime la forme à partir de la feuille de calcul.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricShapeType)|Spécifie le type de forme géométrique de cette forme géométrique.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getAsImage_format_)|Convertit la forme à une image et renvoie l’image comme une chaîne codée en base 64.|
||[height](/javascript/api/excel/excel.shape#height)|Spécifie la hauteur, en points, de la forme.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementLeft_increment_)|Déplace horizontalement la forme spécifiée selon le nombre de points indiqué.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementRotation_increment_)|Fait pivoter la forme spécifiée dans le sens des aiguilles d’une montre, selon le nombre de degrés spécifié, autour de l'axe z.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementTop_increment_)|Décale vers le haut la forme spécifiée selon le nombre de points spécifié.|
||[left](/javascript/api/excel/excel.shape#left)|La distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockAspectRatio)|Spécifie si les proportions de cette forme sont verrouillées.|
||[name](/javascript/api/excel/excel.shape#name)|Spécifie le nom de la forme.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionSiteCount)|Renvoie le nombre de sites de connexion sur la forme spécifiée.|
||[fill](/javascript/api/excel/excel.shape#fill)|Renvoie la mise en forme de remplissage de cette forme.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricShape)|Renvoie la Forme géométrique associée à la forme.|
||[groupe](/javascript/api/excel/excel.shape#group)|Renvoie le groupe de la Forme associée à la forme.|
||[id](/javascript/api/excel/excel.shape#id)|Spécifie l’identificateur de forme.|
||[image](/javascript/api/excel/excel.shape#image)|Renvoie l’image associé à la forme.|
||[level](/javascript/api/excel/excel.shape#level)|Spécifie le niveau de la forme spécifiée.|
||[line](/javascript/api/excel/excel.shape#line)|Renvoie l’image associée à la forme.|
||[lineFormat](/javascript/api/excel/excel.shape#lineFormat)|Renvoie la mise en forme de ligne de cette forme.|
||[onActivated](/javascript/api/excel/excel.shape#onActivated)|Se produit lorsque la forme est activée.|
||[onDeactivated](/javascript/api/excel/excel.shape#onDeactivated)|Se produit lorsque la forme est désactivée.|
||[parentGroup](/javascript/api/excel/excel.shape#parentGroup)|Spécifie le groupe parent de cette forme.|
||[textFrame](/javascript/api/excel/excel.shape#textFrame)|Renvoie l’objet textFrame d’une forme.|
||[type](/javascript/api/excel/excel.shape#type)|Renvoie le type de cette forme.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zOrderPosition)|Renvoie la position de la forme spécifiée dans l’ordre z, valeur z de commande de la forme tout en bas est égal à 0.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Spécifie la rotation, en degrés, de la forme.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleHeight_scaleFactor__scaleType__scaleFrom_)|Met la hauteur de la forme à l’échelle en utilisant un facteur spécifié.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleWidth_scaleFactor__scaleType__scaleFrom_)|Met la largeur de la forme à l’échelle en utilisant un facteur spécifié.|
||[setZOrder(value: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setZOrder_position_)|Déplace la forme spécifiée vers le haut ou vers le bas z de commande de la collection qui décale devant ou derrière les autres formes.|
||[top](/javascript/api/excel/excel.shape#top)|La distance, en points, du bord supérieur de l’objet au bord supérieur de la feuille de calcul.|
||[visible](/javascript/api/excel/excel.shape#visible)|Spécifie si la forme est visible.|
||[width](/javascript/api/excel/excel.shape#width)|Spécifie la largeur, en points, de la forme.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeId)|Obtient l’ID de la forme activée.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle la forme est activée.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_)|Ajoute une forme géométrique à la feuille de calcul.|
||[addGroup (valeurs : matrice < chaîne \| forme >)](/javascript/api/excel/excel.shapecollection#addGroup_values_)|Groupes un sous-ensemble de formes dans la feuille de calcul de cette collection de sites.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_)|Crée une image à partir d’une chaîne en base 64 et il est ajouté à la feuille de calcul.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_)|Ajoute une ligne à la feuille de calcul.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addTextBox_text_)|Ajoute une zone de texte à la feuille de calcul avec le texte fourni en tant que le contenu.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getCount__)|Renvoie le nombre de graphiques dans la feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getItem_key_)|Obtient une forme à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getItemAt_index_)|Obtient une forme en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeId)|Obtient l’ID de la forme désactivée.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle la forme est désactivée.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear__)|Renvoie la mise en forme de remplissage de cette forme.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundColor)|Représente la couleur de premier plan de remplissage de la forme au format HTML, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[type](/javascript/api/excel/excel.shapefill#type)|Renvoie le type de remplissage de la forme.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setSolidColor_color_)|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|
||[Transparency](/javascript/api/excel/excel.shapefill#transparency)|Spécifie le pourcentage de transparence du remplissage sous la forme d’une valeur entre 0.0 (opaque) et 1.0 (clair).|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.shapefont#color)|Représentation de code couleur HTML de la couleur du texte (par exemple, « #FF0000 » représente le rouge).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.shapefont#name)|Représente le nom de la police (par exemple, « Calibri »).|
||[size](/javascript/api/excel/excel.shapefont#size)|Représente la taille de police en points (par exemple, 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Type de soulignement appliqué à la police.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Spécifie l’identificateur de forme.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Renvoie `Shape` l’objet associé au groupe.|
||[Formes](/javascript/api/excel/excel.shapegroup#shapes)|Renvoie la collection `Shape` d’objets.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup__)|Dissocie toutes les formes groupées dans la forme spécifiée.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Représente la couleur de trait au format HTML, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashStyle)|Représente le style de trait de la forme.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Représente le style de trait de la forme.|
||[Transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Renvoie ou définit le degré de transparence du remplissage spécifié sous la forme d’une valeur comprise entre 0,0 (opaque) et 1,0 (transparent).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Spécifie si la mise en forme de trait d’un élément de forme est visible.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Représente l’épaisseur de ligne, en points.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subField)|Spécifie le sous-champ qui est le nom de propriété cible d’une valeur enrichie à trier.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getCount__)|Obtient le nombre de tableaux de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getItemAt_index_)|Obtient une forme en fonction de sa position dans la collection.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autoFilter)|Représente `AutoFilter` l’objet du tableau.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtient la source de l’événement.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableId)|Obtient l’ID de la table qui est ajoutée.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le tableau est ajouté.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[Détails](/javascript/api/excel/excel.tablechangedeventargs#details)|Obtient les informations sur les détails des changements.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onAdded)|Se produit lorsqu’une nouvelle table est ajoutée dans un workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#onDeleted)|Se produit lorsque le tableau spécifié est supprimé dans un classeur.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Obtient la source de l’événement.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableId)|Obtient l’ID de la table qui est supprimée.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tableName)|Obtient le nom de la table qui est supprimée.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le tableau est supprimé.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getCount__)|Obtient le nombre de tableaux de la collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getFirst__)|Obtient le premier tableau de cette collection.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItem_key_)|Obtient un tableau à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autoSizeSetting)|Paramètres de resserrage automatique pour le cadre de texte.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottomMargin)|Représente la marge bas, en points du cadre du texte.|
||[deleteText()](/javascript/api/excel/excel.textframe#deleteText__)|Supprime tout le texte dans la textframe.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalAlignment)|Représente l’alignement horizontal pour le style.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontalOverflow)|Représente le type de débordement horizontal du cadre du texte.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftMargin)|Représente la marge gauche, en points du cadre du texte.|
||[Orientation](/javascript/api/excel/excel.textframe#orientation)|Représente l’angle vers lequel le texte est orienté pour le cadre de texte.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingOrder)|Représente l’ordre de lecture du cadre texte gauche à droite ou de droite à gauche.|
||[hasText](/javascript/api/excel/excel.textframe#hasText)|Spécifie si le cadre de texte contient du texte.|
||[textRange](/javascript/api/excel/excel.textframe#textRange)|Représente le texte lié à une forme, en plus des propriétés et des méthodes de manipulation du texte.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightMargin)|Représente la marge droite, en points du cadre du texte.|
||[topMargin](/javascript/api/excel/excel.textframe#topMargin)|Représente la marge du haut, en points du cadre du texte.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalAlignment)|Représente l’alignement vertical pour le style.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticalOverflow)|Représente le type de débordement vertical du cadre du texte.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getSubstring_start__length_)|Renvoie un objet TextRange pour les caractères dans la plage de donnée.|
||[police](/javascript/api/excel/excel.textrange#font)|Renvoie un `ShapeFont` objet qui représente les attributs de police de la plage de texte.|
||[text](/javascript/api/excel/excel.textrange#text)|Représente le contenu de texte brut de la plage de texte.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartDataPointTrack)|True si tous les graphiques dans le classeur suivent les points de données réelles auquel qu’il sont joints.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getActiveChart__)|Obtient la feuille de calcul active du classeur.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getActiveChartOrNullObject__)|Obtient la feuille de calcul active du classeur.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getIsActiveCollabSession__)|Renvoie si le workbook est modifié par plusieurs `true` utilisateurs (via la co-édition).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getSelectedRanges__)|Obtient la ou les plage(s) sélectionnée(s) actuelle(s) dans le classeur.|
||[isDirty](/javascript/api/excel/excel.workbook#isDirty)|Spécifie si des modifications ont été apportées depuis le dernier enregistré du manuel.|
||[autoSave](/javascript/api/excel/excel.workbook#autoSave)|Spécifie si le workbook est en mode d’auto-ave.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationEngineVersion)|Renvoie un nombre sur la version de moteur de calcul Excel.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged)|Se produit lorsque le paramètre AutoSave est modifié sur le workbook.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslySaved)|Spécifie si le manuel a déjà été enregistré localement ou en ligne.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#usePrecisionAsDisplayed)|True si les calculs réalisés dans ce classeur utiliseront uniquement la précision des nombres tels qu’ils sont affichés. |
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Obtient le type de l’événement.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enableCalculation)|Détermine si Excel recalculer la feuille de calcul si nécessaire.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAll_text__criteria_)|Recherche toutes les occurrences de la chaîne donnée en fonction des critères spécifiés et les renvoie en tant qu’objet, comprenant une ou `RangeAreas` plusieurs plages rectangulaires.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAllOrNullObject_text__criteria_)|Recherche toutes les occurrences de la chaîne donnée en fonction des critères spécifiés et les renvoie en tant qu’objet, comprenant une ou `RangeAreas` plusieurs plages rectangulaires.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getRanges_address_)|Obtient `RangeAreas` l’objet, qui représente un ou plusieurs blocs de plages rectangulaires, spécifiés par l’adresse ou le nom.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autoFilter)|Représente `AutoFilter` l’objet de la feuille de calcul.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalPageBreaks)|Obtient la collection de saut de page horizontal pour la feuille de calcul.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onFormatChanged)|Se produit lorsque le filtre est modifié sur un tableau spécifique.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pageLayout)|Obtient `PageLayout` l’objet de la feuille de calcul.|
||[Formes](/javascript/api/excel/excel.worksheet#shapes)|Renvoie une collection de tous les objets Forme sur la feuille de calcul.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalPageBreaks)|Obtient la collection de saut de page vertical pour la feuille de calcul.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceAll_text__replacement__criteria_)|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[détails](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Représente les informations sur les détails des changements.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onChanged)|Se produit lorsqu’une feuille de calcul dans le classeur est modifiée.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onFormatChanged)|Se produit lorsqu’une feuille de calcul du manuel a un format modifié.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged)|Se produit lorsque la sélection change sur n’importe quelle feuille de calcul.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRange_ctx_)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRangeOrNullObject_ctx_)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle les données ont été modifiées.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completeMatch)|Spécifie si la correspondance doit être complète ou partielle.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchCase)|Spécifie si la correspondance est sensible à la cas.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
