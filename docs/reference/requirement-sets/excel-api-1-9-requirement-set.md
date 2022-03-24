---
title: Excel conditions requises de l’API JavaScript 1.9
description: Détails sur l’ensemble de conditions requises ExcelApi 1.9.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f34b109f95f013cf27f0abfca9c2a8c6b1e4e7c9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746696"
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

Le tableau suivant répertorie les API Excel l’ensemble de conditions requises de l’API JavaScript 1.9. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par Excel l’ensemble de conditions requises de l’API JavaScript 1.9 ou une version antérieure, voir les API Excel dans l’ensemble de conditions requises [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true) ou une version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#excel-excel-application-calculationengineversion-member)|Renvoie la version du moteur de calcul Excel utilisée pour le dernier recalcul complet.|
||[calculationState](/javascript/api/excel/excel.application#excel-excel-application-calculationstate-member)|Renvoie l’état de calcul de l’application.|
||[iterativeCalculation](/javascript/api/excel/excel.application#excel-excel-application-iterativecalculation-member)|Renvoie les paramètres de calcul itératifs.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendscreenupdatinguntilnextsync-member(1))|Suspend la mise à jour de l’écran jusqu’à ce que le `context.sync()` suivant soit appelé.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[Appliquer (plage : plage \| chaîne, columnIndex ? : nombre, critères ? : Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-apply-member(1))|Applique le filtre automatique à une plage.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcriteria-member(1))|Cette fonction permet d’effacer les critères de filtre et l’état de tri du filtre automatique.|
||[criteria](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-criteria-member)|Tableau qui conserve tous les critères de filtre dans une plage filtrée.|
||[enabled](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-enabled-member)|Spécifie si le filtre automatique est activé.|
||[getRange()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrange-member(1))|Renvoie l’objet `Range` qui représente la plage à laquelle le filtre automatique s’applique.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrangeornullobject-member(1))|Renvoie l’objet `Range` qui représente la plage à laquelle le filtre automatique s’applique.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-isdatafiltered-member)|Spécifie si le filtre automatique a des critères de filtre.|
||[reapply()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-reapply-member(1))|Applique l’objet Autofilter spécifié actuellement sur la plage.|
||[remove()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-remove-member(1))|Supprime le filtre automatique pour la plage.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-color-member)|Représente la `color` propriété d’une bordure simple.|
||[style](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-style-member)|Représente la `style` propriété d’une bordure simple.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-tintandshade-member)|Représente la `tintAndShade` propriété d’une bordure simple.|
||[weight](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-weight-member)|Représente la `weight` propriété d’une bordure simple.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bas](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-bottom-member)|Représente la `format.borders.bottom` propriété.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonaldown-member)|Représente la `format.borders.diagonalDown` propriété.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonalup-member)|Représente la `format.borders.diagonalUp` propriété.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-horizontal-member)|Représente la `format.borders.horizontal` propriété.|
||[left](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-left-member)|Représente la `format.borders.left` propriété.|
||[right](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-right-member)|Représente la `format.borders.right` propriété.|
||[top](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-top-member)|Représente la `format.borders.top` propriété.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-vertical-member)|Représente la `format.borders.vertical` propriété.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[adresse](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-address-member)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-addresslocal-member)|Représente la `addressLocal` propriété.|
||[hidden](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-hidden-member)|Représente la `hidden` propriété.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-color-member)|Représente la `format.fill.color` propriété.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-pattern-member)|Représente la `format.fill.pattern` propriété.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterncolor-member)|Représente la `format.fill.patternColor` propriété.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterntintandshade-member)|Représente la `format.fill.patternTintAndShade` propriété.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-tintandshade-member)|Représente la `format.fill.tintAndShade` propriété.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-bold-member)|Représente la `format.font.bold` propriété.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-color-member)|Représente la `format.font.color` propriété.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-italic-member)|Représente la `format.font.italic` propriété.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-name-member)|Représente la `format.font.name` propriété.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-size-member)|Représente la `format.font.size` propriété.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-strikethrough-member)|Représente la `format.font.strikethrough` propriété.|
||[Subscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-subscript-member)|Représente la `format.font.subscript` propriété.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-superscript-member)|Représente la `format.font.superscript` propriété.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-tintandshade-member)|Représente la `format.font.tintAndShade` propriété.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-underline-member)|Représente la `format.font.underline` propriété.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-autoindent-member)|Représente la `autoIndent` propriété.|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-borders-member)|Représente la `borders` propriété.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-fill-member)|Représente la `fill` propriété.|
||[police](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-font-member)|Représente la `font` propriété.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-horizontalalignment-member)|Représente la `horizontalAlignment` propriété.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-indentlevel-member)|Représente la `indentLevel` propriété.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-protection-member)|Représente la `protection` propriété.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-readingorder-member)|Représente la `readingOrder` propriété.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-shrinktofit-member)|Représente la `shrinkToFit` propriété.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-textorientation-member)|Représente la `textOrientation` propriété.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member)|Représente la `useStandardHeight` propriété.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member)|Représente la `useStandardWidth` propriété.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-verticalalignment-member)|Représente la `verticalAlignment` propriété.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-wraptext-member)|Représente la `wrapText` propriété.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-formulahidden-member)|Représente la `format.protection.formulaHidden` propriété.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-locked-member)|Représente la `format.protection.locked` propriété.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valueafter-member)|Représente la valeur après la modification.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuebefore-member)|Représente la valeur avant la modification.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypeafter-member)|Représente le type de valeur après la modification.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypebefore-member)|Représente le type de valeur avant la modification.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#excel-excel-chart-activate-member(1))|Active le graphique dans l’interface utilisateur Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#excel-excel-chart-pivotoptions-member)|Encapsule les options pour le graphique croisé dynamique.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-colorscheme-member)|Spécifie le modèle de couleurs du graphique.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-roundedcorners-member)|Spécifie si la zone de graphique du graphique possède des coins arrondis.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-linknumberformat-member)|Spécifie si le format numérique est lié aux cellules.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowoverflow-member)|Spécifie si le débordement bin est activé dans un histogramme ou un graphique de pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowunderflow-member)|Spécifie si le sous-flux bin est activé dans un histogramme ou un graphique de pareto.|
||[count](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-count-member)|Spécifie le nombre de bacs d’un histogramme ou d’un graphique de pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-overflowvalue-member)|Spécifie la valeur de débordement bin d’un histogramme ou d’un graphique de pareto.|
||[type](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-type-member)|Spécifie le type du bac pour un histogramme ou un graphique de pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-underflowvalue-member)|Spécifie la valeur de sous-flux bin d’un histogramme ou d’un graphique de pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-width-member)|Spécifie la valeur de largeur bin d’un histogramme ou d’un graphique de pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-quartilecalculation-member)|Spécifie si le type de calcul du quartile d’un graphique de zone et de zone.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showinnerpoints-member)|Spécifie si les points internes sont affichés dans une zone et un graphique de zone.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanline-member)|Spécifie si la ligne moyenne est affichée dans un graphique de zone et de zone.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanmarker-member)|Spécifie si la marque moyenne est affichée dans un graphique zone et de zone.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showoutlierpoints-member)|Spécifie si les points aberrants sont affichés dans un graphique zone et valeur.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-linknumberformat-member)|Spécifie si le format numérique est lié aux cellules (de sorte que le format numérique change dans les étiquettes lorsqu’il change dans les cellules).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-linknumberformat-member)|Spécifie si le format numérique est lié aux cellules.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-endstylecap-member)|Spécifie si les barres d’erreur ont un style de fin.|
||[format](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-format-member)|Spécifie le type de mise en forme de barres d’erreur.|
||[inclure](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-include-member)|Spécifie les parties de barres d’erreur à inclure.|
||[type](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-type-member)|Le type de plage marqué par des barres d’erreur.|
||[visible](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-visible-member)|Spécifie si les barres d’erreur sont affichées.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#excel-excel-charterrorbarsformat-line-member)|Représente le format des lignes du graphique.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-labelstrategy-member)|Spécifie la stratégie d’étiquettes de carte de série d’un graphique de carte région.|
||[level](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-level-member)|Spécifie le niveau de mappage des séries d’un graphique de carte région.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-projectiontype-member)|Spécifie le type de projection de série d’un graphique région carte.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showaxisfieldbuttons-member)|Spécifie s’il faut afficher les boutons de champ d’axe sur une PivotChart.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showlegendfieldbuttons-member)|Spécifie s’il faut afficher les boutons de champ de légende sur un PivotChart.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showreportfilterfieldbuttons-member)|Spécifie s’il faut afficher les boutons de champ de filtre de rapport sur une PivotChart.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showvaluefieldbuttons-member)|Spécifie s’il faut afficher les boutons de champ afficher la valeur sur une PivotChart.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[binOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-binoptions-member)|Encapsule les options bin uniquement pour les histogrammes et graphiques de pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-boxwhiskeroptions-member)|Résume les options pour le graphique de zone et valeur.|
||[bubbleScale](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-bubblescale-member)|Peut être une valeur d’entier entre 0 (zéro) et 300 correspondant à un pourcentage de la taille par défaut.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumcolor-member)|Spécifie la couleur de la valeur maximale d’une série de graphique région carte.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumtype-member)|Spécifie le type de valeur maximale d’une série de graphique région carte.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumvalue-member)|Spécifie la valeur maximale d’une série de graphique région carte.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointcolor-member)|Spécifie la couleur de la valeur du milieu d’une série de graphique région carte.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointtype-member)|Spécifie le type de la valeur de milieu d’une série de graphique région carte.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointvalue-member)|Spécifie la valeur du milieu d’une série de graphique région carte.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumcolor-member)|Spécifie la couleur de la valeur minimale d’une série de graphique région carte.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumtype-member)|Spécifie le type de la valeur minimale d’une série de graphique région carte.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumvalue-member)|Spécifie la valeur minimale d’une série de graphique région carte.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientstyle-member)|Spécifie le style de dégradé de série d’un graphique région carte.|
||[invertColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertcolor-member)|Spécifie la couleur de remplissage des points de données négatifs d’une série.|
||[mapOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-mapoptions-member)|Encapsule les options pour le graphique carte de région.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-parentlabelstrategy-member)|Spécifie la zone de stratégie d’étiquette parent de série pour un graphique en arborescence.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showconnectorlines-member)|Spécifie si les lignes de connecteur sont affichées dans les graphiques en cascade.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showleaderlines-member)|Spécifie si les lignes d’étiquettes sont affichées pour chaque étiquette de données de la série.|
||[splitValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splitvalue-member)|Spécifie la valeur de seuil qui sépare deux sections d’un graphique en secteurs de secteur ou d’un graphique en barres de secteur.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-xerrorbars-member)|Représente l’objet de la barre d’erreur pour une série de graphique.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-yerrorbars-member)|Représente l’objet de la barre d’erreur pour une série de graphique.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-linknumberformat-member)|Spécifie si le format numérique est lié aux cellules (de sorte que le format numérique change dans les étiquettes lorsqu’il change dans les cellules).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[adresse](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-address-member)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-addresslocal-member)|Représente la `addressLocal` propriété.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-columnindex-member)|Représente la `columnIndex` propriété.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getranges-member(1))|Renvoie la `RangeAreas`, comprenant une ou plusieurs plages rectangulaires, à laquelle le format conditionnel est appliqué.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcells-member(1))|Renvoie un `RangeAreas` objet, comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valides.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcellsornullobject-member(1))|Renvoie un `RangeAreas` objet, comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valides.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#excel-excel-filtercriteria-subfield-member)|Propriété utilisée par le filtre pour faire un filtre enrichi sur des valeurs enrichies.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-id-member)|Représente l’identificateur de forme.|
||[shape](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-shape-member)|Renvoie l’objet `Shape` de la forme géométrique.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getcount-member(1))|Renvoie le nombre de formes dans le groupe de la forme.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitem-member(1))|Obtient une forme à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemat-member(1))|Obtient une forme en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooter-member)|Pied de la feuille de calcul.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheader-member)|En-tête central de la feuille de calcul.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooter-member)|Pied de la feuille de calcul gauche.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheader-member)|En-tête gauche de la feuille de calcul.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooter-member)|Pied de la feuille de calcul à droite.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheader-member)|En-tête droit de la feuille de calcul.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-defaultforallpages-member)|L’en-tête/pied de page, utilisé pour toutes les pages, sauf si la première page ou page impaire/paire est spécifiée.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-evenpages-member)|L’en-tête/le pied de page à utiliser pour les pages paires, en-tête/pied de page impaire doit être spécifié pour les pages impaires.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-firstpage-member)|La première en-tête/le premier pied de page, pour toutes les autres pages générales ou impair/pair est utilisé.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-oddpages-member)|L’en-tête/le pied de page à utiliser pour les pages paires, l’en-tête/pied de page paire doit être spécifié pour les pages paires.|
||[state](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-state-member)|État selon lequel les en-têtes/pieds de pied de groupe sont définies.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetmargins-member)|Obtient ou définit un indicateur indiquant si les en-têtes/pieds de page sont alignés avec les marges de page définis dans les options de mise en page pour la feuille de calcul.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetscale-member)|Obtient ou définit un indicateur indiquant si les en-têtes/pieds de page sont alignés avec les marges de page définis dans les options de mise en page pour la feuille de calcul.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#excel-excel-image-format-member)|Renvoie le format de l’image.|
||[id](/javascript/api/excel/excel.image#excel-excel-image-id-member)|Spécifie l’identificateur de forme de l’objet image.|
||[shape](/javascript/api/excel/excel.image#excel-excel-image-shape-member)|Renvoie l’objet `Shape` associé à l’image.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-enabled-member)|Cette propriété a la valeur True si Microsoft Excel utilise l'itération pour résoudre des références circulaires.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxchange-member)|Spécifie la quantité maximale de modification entre chaque itération à mesure Excel des références circulaires.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxiteration-member)|Spécifie le nombre maximal d’itérations que Excel pouvez utiliser pour résoudre une référence circulaire.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadlength-member)|Renvoie ou définit la longueur de la pointe de la flèche au début de la ligne spécifiée.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadstyle-member)|Représente le style de la pointe de la flèche au début de la ligne spécifiée.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadwidth-member)|Représente la largeur de la pointe de la flèche au début de la ligne spécifiée.|
||[beginConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedshape-member)|Représente la forme de la pointe de la flèche au début de la ligne spécifiée.|
||[beginConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedsite-member)|Représente le site de connexion indiquant le point de connexion auquel le début d'un connecteur est relié.|
||[connectBeginShape (forme : Excel.Shape, connectionSite : nombre)](/javascript/api/excel/excel.line#excel-excel-line-connectbeginshape-member(1))|Détache le début du connecteur spécifié de la forme à laquelle il est attaché.|
||[connectBeginShape (forme : Excel.Shape, connectionSite : nombre)](/javascript/api/excel/excel.line#excel-excel-line-connectendshape-member(1))|Joint la fin du connecteur spécifié à une forme spécifiée.|
||[connectorType](/javascript/api/excel/excel.line#excel-excel-line-connectortype-member)|Représente le type de connecteur pour la ligne.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectbeginshape-member(1))|Détache le début du connecteur spécifié de la forme à laquelle il est attaché.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectendshape-member(1))|Détache la fin du connecteur spécifié de la forme à laquelle il est attaché.|
||[endArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadlength-member)|Représente la longueur de la pointe de la flèche à la fin de la ligne spécifiée.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadstyle-member)|Représente le style de la pointe de la flèche à la fin de ligne spécifée.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadwidth-member)|Représente la largeur de la pointe de la flèche à la fin de la ligne spécifiée.|
||[endConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-endconnectedshape-member)|Représente la forme de la pointe de la flèche à la fin de la ligne spécifiée.|
||[endConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-endconnectedsite-member)|Représente le site de connexion indiquant le point de connexion auquel la fin d'un connecteur est relié.|
||[id](/javascript/api/excel/excel.line#excel-excel-line-id-member)|Spécifie l’identificateur de forme.|
||[isBeginConnected](/javascript/api/excel/excel.line#excel-excel-line-isbeginconnected-member)|Spécifie si le début de la ligne spécifiée est connecté à une forme.|
||[isEndConnected](/javascript/api/excel/excel.line#excel-excel-line-isendconnected-member)|Spécifie si la fin de la ligne spécifiée est connectée à une forme.|
||[shape](/javascript/api/excel/excel.line#excel-excel-line-shape-member)|Renvoie l’objet `Shape` associé à la ligne.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[columnIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-columnindex-member)|Spécifie l’index de colonne pour le pause de page.|
||[delete()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-delete-member(1))|Supprime un objet de saut de page.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-getcellafterbreak-member(1))|Obtient la première cellule après le saut de page.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-rowindex-member)|Spécifie l’index de ligne pour le pause de page.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[Ajouter (pageBreakRange : plage \| chaîne)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-add-member(1))|Ajoute un saut de page avant la cellule en haut à gauche de la plage spécifiée.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getcount-member(1))|Obtient le nombre de pages de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getitem-member(1))|Obtient un objet de saut de page via l’index.|
||[items](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-removepagebreaks-member(1))|Redéfinit tous les sauts de page de la collection.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-blackandwhite-member)|Option d’impression noir et blanc de la feuille de calcul.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-bottommargin-member)|Marge de page inférieure de la feuille de calcul à utiliser pour l’impression en points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centerhorizontally-member)|Indicateur horizontal du centre de la feuille de calcul.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centervertically-member)|Indicateur vertical du centre de la feuille de calcul.|
||[draftMode](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-draftmode-member)|Option de mode brouillon de la feuille de calcul.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-firstpagenumber-member)|Premier numéro de page de la feuille de calcul à imprimer.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-footermargin-member)|Marge de pied de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintarea-member(1))|Obtient l’objet `RangeAreas` , comprenant une ou plusieurs plages rectangulaires, qui représente la zone d’impression de la feuille de calcul.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintareaornullobject-member(1))|Obtient l’objet `RangeAreas` , comprenant une ou plusieurs plages rectangulaires, qui représente la zone d’impression de la feuille de calcul.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumns-member(1))|Obtient l’objet plage représentant les colonnes de titre.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumnsornullobject-member(1))|Obtient l’objet plage représentant les colonnes de titre.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerows-member(1))|Obtient l’objet plage représentant les rangées de titre.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerowsornullobject-member(1))|Obtient l’objet plage représentant les rangées de titre.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headermargin-member)|Marge d’en-tête de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headersfooters-member)|Configuration de l’en-tête et pied de page de la feuille de calcul.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-leftmargin-member)|Marge gauche de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[Orientation](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-orientation-member)|Orientation de la feuille de calcul de la page.|
||[paperSize](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-papersize-member)|Format de papier de la feuille de calcul de la page.|
||[printComments](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printcomments-member)|Spécifie si les commentaires de la feuille de calcul doivent être affichés lors de l’impression.|
||[printErrors](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printerrors-member)|Option d’erreurs d’impression de la feuille de calcul.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printgridlines-member)|Spécifie si le quadrillage de la feuille de calcul sera imprimé.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printheadings-member)|Spécifie si les en-tête de la feuille de calcul seront imprimés.|
||[printOrder](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printorder-member)|Option d’ordre d’impression de page de la feuille de calcul.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-rightmargin-member)|Marge droite de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[setPrintArea (printArea : plage \| RangeAreas \| chaîne)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintarea-member(1))|Définit la zone d’impression de la feuille de calcul.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintmargins-member(1))|Définit les marges de page de la feuille de calcul avec des unités.|
||[setPrintTitleColumns (printTitleColumns : plage \| chaîne)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlecolumns-member(1))|Définit les colonnes qui contiennent des cellules répétées à gauche de chaque page de la feuille de calcul pour l’impression.|
||[setPrintTitleRows (printTitleRows : plage \| chaîne)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlerows-member(1))|Définit les rangées qui contiennent des cellules répétées en haut de chaque page de la feuille de calcul pour l’impression.|
||[topMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-topmargin-member)|Marge supérieure de la feuille de calcul, en points, à utiliser lors de l’impression.|
||[zoom](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-zoom-member)|Options de zoom avant impression de la feuille de calcul.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bas](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-bottom-member)|Spécifie la marge inférieure de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-footer-member)|Spécifie la marge de pied de page de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-header-member)|Spécifie la marge d’en-tête de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-left-member)|Spécifie la marge gauche de la mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-right-member)|Spécifie la marge droite de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-top-member)|Spécifie la marge supérieure de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-horizontalfittopages-member)|Nombre de pages pour l’ajuster horizontalement.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-scale-member)|La valeur d’échelle de page d’impression peut être comprise entre 10 et 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-verticalfittopages-member)|Nombre de pages pour l’ajuster verticalement.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbyvalues-member(1))|Trie le PivotField par valeurs spécifiées dans une étendue donnée.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-autoformat-member)|Spécifie si la mise en forme sera automatiquement mise en forme lorsqu’elle est actualisée ou lorsque les champs sont déplacés.|
||[getDataHierarchy (cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatahierarchy-member(1))|Obtient DataHierarchy servant à calculer la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[getPivotItems (axe : Excel.PivotAxis, cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getpivotitems-member(1))|Obtient le PivotItems à partir d’un axe qui composent la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-preserveformatting-member)|Spécifie si la mise en forme est conservée lorsque le rapport est actualisé ou recalculé par des opérations telles que le pivotage, le tri ou la modification d’éléments de champ de page.|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setautosortoncell-member(1))|Définit le tableau croisé dynamique pour trier automatiquement à l’aide de la cellule spécifiée pour sélectionner automatiquement tous les critères et contexte nécessaires.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-enabledatavalueediting-member)|Spécifie si le tableau croisé dynamique permet à l’utilisateur de modifier des valeurs dans le corps des données.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-usecustomsortlists-member)|Spécifie si le tableau croisé dynamique utilise des listes personnalisées lors du tri.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Range \| string, autoFillType?: Excel. AutoFillType)](/javascript/api/excel/excel.range#excel-excel-range-autofill-member(1))|Remplit la plage actuelle à la plage de destination à l’aide de la logique de remplissage automatique spécifiée.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#excel-excel-range-convertdatatypetotext-member(1))|Convertit en texte les cellules de plage avec des types de données.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#excel-excel-range-converttolinkeddatatype-member(1))|Convertit les cellules de la plage en types de données liés dans la feuille de calcul.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1))|Copie les données de cellule ou la mise en forme de la plage source ou `RangeAreas` de la plage actuelle.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-find-member(1))|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-findornullobject-member(1))|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[flashFill()](/javascript/api/excel/excel.range#excel-excel-range-flashfill-member(1))|Fait un remplissage flash à la plage actuelle.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcellproperties-member(1))|Renvoie une plage en 2D, qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcolumnproperties-member(1))|Renvoie une plage à dimension unique, qui comprend les données de char colonne de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getrowproperties-member(1))|Renvoie une plage à dimension unique , qui comprend les données de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1))|Obtient l’objet `RangeAreas` , comprenant une ou plusieurs plages rectangulaires, qui représente toutes les cellules qui correspondent au type et à la valeur spécifiés.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1))|Obtient `RangeAreas` l’objet, comprenant une ou plusieurs plages, qui représente toutes les cellules qui correspondent au type et à la valeur spécifiés.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-gettables-member(1))|Obtient une collection de tableaux qui se chevauchent avec la plage dans l’étendue.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#excel-excel-range-linkeddatatypestate-member)|Représente l’état du type de données de chaque cellule.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))|Supprime les valeurs dupliquées de la plage spécifiée par les colonnes.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#excel-excel-range-replaceall-member(1))|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#excel-excel-range-setcellproperties-member(1))|Met à jour la plage en fonction d’un tableau 2D de propriétés de cellule, en encapsulant des éléments tels que la police, le remplissage, les bordures et l’alignement.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setcolumnproperties-member(1))|Met à jour la plage basée sur un tableau à une dimension de propriétés de colonne, en encapsulant des éléments tels que la police, le remplissage, les bordures et l’alignement.|
||[setDirty()](/javascript/api/excel/excel.range#excel-excel-range-setdirty-member(1))|Cette méthode désigne une plage qui doit être recalculée lorsque le recalcul suivant se produit.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setrowproperties-member(1))|Met à jour la plage basée sur un tableau à une dimension de propriétés de ligne, en encapsulant des éléments tels que la police, le remplissage, les bordures et l’alignement.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[adresse](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-address-member)|Renvoie la `RangeAreas` référence en style A1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-addresslocal-member)|Renvoie la référence `RangeAreas` dans les paramètres régionaux de l’utilisateur.|
||[areaCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areacount-member)|Renvoie le nombre de plages rectangulaires qui composent cet `RangeAreas` objet.|
||[Zones](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areas-member)|Renvoie une collection de plages rectangulaires qui composent cet `RangeAreas` objet.|
||[calculate()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-calculate-member(1))|Calcule toutes les cellules dans le `RangeAreas`.|
||[cellCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-cellcount-member)|Renvoie le nombre de cellules dans l’objet `RangeAreas` , récapitulant le nombre de cellules de toutes les plages rectangulaires individuelles.|
||[Effacer (applyTo ? : Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clear-member(1))|Cette propriété permet d’effacer les valeurs, le format, le remplissage, la bordure et d’autres propriétés de chacune des zones qui composent cet `RangeAreas` objet.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-conditionalformats-member)|Renvoie une collection de formats conditionnels qui se coupent avec les cellules de cet `RangeAreas` objet.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-convertdatatypetotext-member(1))|Convertit toutes les cellules des types de `RangeAreas` données en texte.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-converttolinkeddatatype-member(1))|Convertit toutes les cellules de l’ensemble en `RangeAreas` types de données liés.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-copyfrom-member(1))|Copie les données de cellule ou la mise en forme de la plage source ou `RangeAreas` de la plage actuelle `RangeAreas`.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-datavalidation-member)|Renvoie un objet de validation de données pour toutes les plages dans le `RangeAreas`.|
||[format](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-format-member)|Renvoie un `RangeFormat` objet, qui encapsule la police, le remplissage, les bordures, l’alignement et d’autres propriétés de toutes les plages de l’objet `RangeAreas` .|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirecolumn-member(1))|`RangeAreas` Renvoie un objet qui représente les colonnes entières `RangeAreas` de l'(par exemple, `RangeAreas` si le courant représente les cellules « B4:E11, H2 », il renvoie un `RangeAreas` qui représente les colonnes « B:E, H:H »).|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirerow-member(1))|Renvoie un `RangeAreas` objet qui représente les lignes entières `RangeAreas` de l'(par exemple, `RangeAreas` si le courant représente les cellules « B4:E11 `RangeAreas` », il renvoie un qui représente les lignes « 4:11 »).|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersection-member(1))|Renvoie l’objet `RangeAreas` qui représente l’intersection des plages données ou `RangeAreas`.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersectionornullobject-member(1))|Renvoie l’objet `RangeAreas` qui représente l’intersection des plages données ou `RangeAreas`.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getoffsetrangeareas-member(1))|Renvoie un `RangeAreas` objet décalé par le décalage de ligne et de colonne spécifique.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcells-member(1))|Renvoie un `RangeAreas` objet qui représente toutes les cellules qui correspondent au type et à la valeur spécifiés.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcellsornullobject-member(1))|Renvoie un `RangeAreas` objet qui représente toutes les cellules qui correspondent au type et à la valeur spécifiés.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-gettables-member(1))|Renvoie une collection étendue de tableaux qui chevauchent n’importe quelle plage de cet `RangeAreas` objet.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareas-member(1))|Renvoie l’objet utilisé `RangeAreas` qui comprend toutes les zones utilisées de plages rectangulaires individuelles dans l’objet `RangeAreas` .|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareasornullobject-member(1))|Renvoie l’objet utilisé `RangeAreas` qui comprend toutes les zones utilisées de plages rectangulaires individuelles dans l’objet `RangeAreas` .|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirecolumn-member)|Spécifie si toutes les plages de cet objet représentent des colonnes entières `RangeAreas` (par exemple, « A:C, Q:Z »).|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirerow-member)|Spécifie si toutes les plages `RangeAreas` de cet objet représentent des lignes entières (par exemple, « 1:3, 5:7 »).|
||[setDirty()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-setdirty-member(1))|Définit les `RangeAreas` données à recalculer lors du recalcul suivant.|
||[style](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-style-member)|Représente le style de toutes les plages de cet `RangeAreas` objet.|
||[worksheet](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-worksheet-member)|Renvoie la feuille de calcul pour l’actuel `RangeAreas`.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-tintandshade-member)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour la bordure de plage, la valeur est entre -1 (plus sombre) et 1 (plus clair), avec 0 pour la couleur d’origine.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-tintandshade-member)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour les bordures de plage.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getcount-member(1))|Renvoie le nombre de plages dans le `RangeCollection`.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getitemat-member(1))|Renvoie l’objet de plage en fonction de sa position dans le `RangeCollection`.|
||[items](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-pattern-member)|Motif d’une plage.|
||[patternColor](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterncolor-member)|Code couleur HTML représentant la couleur du modèle de plage, sous la forme #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterntintandshade-member)|Spécifie un double qui s’éclaircit ou assombrit une couleur de motif pour le remplissage de la plage.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-tintandshade-member)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour le remplissage de la plage.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-strikethrough-member)|Spécifie l’état de la police de type strikethrough.|
||[Subscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-subscript-member)|Spécifie l’état d’indice de la police.|
||[superscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-superscript-member)|Spécifie l’état d’exposant de la police.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-tintandshade-member)|Spécifie un double qui s’éclaircit ou assombrit une couleur pour la police de plage.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autoindent-member)|Spécifie si le texte est automatiquement mis en retrait lorsque l’alignement du texte est égal à la distribution.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-indentlevel-member)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-readingorder-member)|L’ordre de lecture de la plage.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-shrinktofit-member)|Spécifie si le texte est automatiquement réduit pour tenir dans la largeur de colonne disponible.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-removed-member)|Nombre de lignes dupliquées supprimées par l’opération.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-uniqueremaining-member)|Nombre de lignes uniques restantes présents dans la plage qui en résulte.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-completematch-member)|Spécifie si la correspondance doit être complète ou partielle.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-matchcase-member)|Spécifie si la correspondance est sensible à la cas.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[adresse](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-address-member)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-addresslocal-member)|Représente la `addressLocal` propriété.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-rowindex-member)|Représente la `rowIndex` propriété.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-completematch-member)|Spécifie si la correspondance doit être complète ou partielle.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-matchcase-member)|Spécifie si la correspondance est sensible à la cas.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-searchdirection-member)|Détermine le sens de la recherche.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-format-member)|Représente la `format` propriété.|
||[lien hypertexte](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-hyperlink-member)|Représente la `hyperlink` propriété.|
||[style](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-style-member)|Représente la `style` propriété.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnhidden-member)|Représente la `columnHidden` propriété.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnwidth-member)||
||[format : Excel. CellPropertiesFormat & { columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-format-member)|Représente la `format` propriété.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format : Excel. CellPropertiesFormat & { rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-format-member)|Représente la `format` propriété.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowheight-member)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowhidden-member)|Représente la `rowHidden` propriété.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#excel-excel-shape-alttextdescription-member)|Spécifie l’autre texte de description d’un `Shape` objet.|
||[altTextTitle](/javascript/api/excel/excel.shape#excel-excel-shape-alttexttitle-member)|Spécifie le texte de titre de remplacement d’un `Shape` objet.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#excel-excel-shape-connectionsitecount-member)|Renvoie le nombre de sites de connexion sur la forme spécifiée.|
||[delete()](/javascript/api/excel/excel.shape#excel-excel-shape-delete-member(1))|Supprime la forme à partir de la feuille de calcul.|
||[fill](/javascript/api/excel/excel.shape#excel-excel-shape-fill-member)|Renvoie la mise en forme de remplissage de cette forme.|
||[geometricShape](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshape-member)|Renvoie la Forme géométrique associée à la forme.|
||[geometricShapeType](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshapetype-member)|Spécifie le type de forme géométrique de cette forme géométrique.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1))|Convertit la forme à une image et renvoie l’image comme une chaîne codée en base 64.|
||[groupe](/javascript/api/excel/excel.shape#excel-excel-shape-group-member)|Renvoie le groupe de la Forme associée à la forme.|
||[height](/javascript/api/excel/excel.shape#excel-excel-shape-height-member)|Spécifie la hauteur, en points, de la forme.|
||[id](/javascript/api/excel/excel.shape#excel-excel-shape-id-member)|Spécifie l’identificateur de forme.|
||[image](/javascript/api/excel/excel.shape#excel-excel-shape-image-member)|Renvoie l’image associé à la forme.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementleft-member(1))|Déplace horizontalement la forme spécifiée selon le nombre de points indiqué.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementrotation-member(1))|Fait pivoter la forme spécifiée dans le sens des aiguilles d’une montre, selon le nombre de degrés spécifié, autour de l'axe z.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementtop-member(1))|Décale vers le haut la forme spécifiée selon le nombre de points spécifié.|
||[left](/javascript/api/excel/excel.shape#excel-excel-shape-left-member)|La distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[level](/javascript/api/excel/excel.shape#excel-excel-shape-level-member)|Spécifie le niveau de la forme spécifiée.|
||[line](/javascript/api/excel/excel.shape#excel-excel-shape-line-member)|Renvoie l’image associée à la forme.|
||[lineFormat](/javascript/api/excel/excel.shape#excel-excel-shape-lineformat-member)|Renvoie la mise en forme de ligne de cette forme.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#excel-excel-shape-lockaspectratio-member)|Spécifie si les proportions de cette forme sont verrouillées.|
||[name](/javascript/api/excel/excel.shape#excel-excel-shape-name-member)|Spécifie le nom de la forme.|
||[onActivated](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)|Se produit lorsque la forme est activée.|
||[onDeactivated](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)|Se produit lorsque la forme est désactivée.|
||[parentGroup](/javascript/api/excel/excel.shape#excel-excel-shape-parentgroup-member)|Spécifie le groupe parent de cette forme.|
||[rotation](/javascript/api/excel/excel.shape#excel-excel-shape-rotation-member)|Spécifie la rotation, en degrés, de la forme.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scaleheight-member(1))|Met la hauteur de la forme à l’échelle en utilisant un facteur spécifié.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scalewidth-member(1))|Met la largeur de la forme à l’échelle en utilisant un facteur spécifié.|
||[setZOrder(value: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#excel-excel-shape-setzorder-member(1))|Déplace la forme spécifiée vers le haut ou vers le bas z de commande de la collection qui décale devant ou derrière les autres formes.|
||[textFrame](/javascript/api/excel/excel.shape#excel-excel-shape-textframe-member)|Renvoie l’objet textFrame d’une forme.|
||[top](/javascript/api/excel/excel.shape#excel-excel-shape-top-member)|La distance, en points, du bord supérieur de l’objet au bord supérieur de la feuille de calcul.|
||[type](/javascript/api/excel/excel.shape#excel-excel-shape-type-member)|Renvoie le type de cette forme.|
||[visible](/javascript/api/excel/excel.shape#excel-excel-shape-visible-member)|Spécifie si la forme est visible.|
||[width](/javascript/api/excel/excel.shape#excel-excel-shape-width-member)|Spécifie la largeur, en points, de la forme.|
||[zOrderPosition](/javascript/api/excel/excel.shape#excel-excel-shape-zorderposition-member)|Renvoie la position de la forme spécifiée dans l’ordre z, valeur z de commande de la forme tout en bas est égal à 0.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-shapeid-member)|Obtient l’ID de la forme activée.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle la forme est activée.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1))|Ajoute une forme géométrique à la feuille de calcul.|
||[addGroup (valeurs : matrice < chaîne \| forme >)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgroup-member(1))|Groupes un sous-ensemble de formes dans la feuille de calcul de cette collection de sites.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1))|Crée une image à partir d’une chaîne en base 64 et il est ajouté à la feuille de calcul.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1))|Ajoute une ligne à la feuille de calcul.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1))|Ajoute une zone de texte à la feuille de calcul avec le texte fourni en tant que le contenu.|
||[getCount()](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getcount-member(1))|Renvoie le nombre de graphiques dans la feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitem-member(1))|Obtient une forme à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemat-member(1))|Obtient une forme en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-shapeid-member)|Obtient l’ID de la forme désactivée.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle la forme est désactivée.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-clear-member(1))|Renvoie la mise en forme de remplissage de cette forme.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-foregroundcolor-member)|Représente la couleur de premier plan de remplissage de la forme au format HTML, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-setsolidcolor-member(1))|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|
||[Transparency](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-transparency-member)|Spécifie le pourcentage de transparence du remplissage sous la forme d’une valeur entre 0.0 (opaque) et 1.0 (clair).|
||[type](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-type-member)|Renvoie le type de remplissage de la forme.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-bold-member)|Représente le format de police Gras.|
||[color](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-color-member)|Représentation de code couleur HTML de la couleur du texte (par exemple, « #FF0000 » représente le rouge).|
||[italic](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-italic-member)|Représente le format de police Italique.|
||[name](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-name-member)|Représente le nom de la police (par exemple, « Calibri »).|
||[taille](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-size-member)|Représente la taille de police en points (par exemple, 11).|
||[underline](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-underline-member)|Type de soulignement appliqué à la police.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-id-member)|Spécifie l’identificateur de forme.|
||[shape](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shape-member)|Renvoie l’objet `Shape` associé au groupe.|
||[Formes](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shapes-member)|Renvoie la collection d’objets `Shape` .|
||[ungroup()](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-ungroup-member(1))|Dissocie toutes les formes groupées dans la forme spécifiée.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-color-member)|Représente la couleur de trait au format HTML, au format #RRGGBB (par exemple, « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-dashstyle-member)|Représente le style de trait de la forme.|
||[style](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-style-member)|Représente le style de trait de la forme.|
||[Transparency](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-transparency-member)|Renvoie ou définit le degré de transparence du remplissage spécifié sous la forme d’une valeur comprise entre 0,0 (opaque) et 1,0 (transparent).|
||[visible](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-visible-member)|Spécifie si la mise en forme de trait d’un élément de forme est visible.|
||[weight](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-weight-member)|Représente l’épaisseur de ligne, en points.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#excel-excel-sortfield-subfield-member)|Spécifie le sous-champ qui est le nom de propriété cible d’une valeur enrichie à trier.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getcount-member(1))|Obtient le nombre de tableaux de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemat-member(1))|Obtient une forme en fonction de sa position dans la collection.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#excel-excel-table-autofilter-member)|Représente l’objet `AutoFilter` du tableau.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-source-member)|Obtient la source de l’événement.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-tableid-member)|Obtient l’ID du tableau qui est ajouté.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle le tableau est ajouté.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[Détails](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-details-member)|Obtient les informations sur les détails des changements.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member)|Se produit lorsqu’une nouvelle table est ajoutée dans un workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member)|Se produit lorsque le tableau spécifié est supprimé dans un classeur.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-source-member)|Obtient la source de l’événement.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tableid-member)|Obtient l’ID de la table qui est supprimée.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tablename-member)|Obtient le nom de la table qui est supprimée.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle le tableau est supprimé.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getcount-member(1))|Obtient le nombre de tableaux de la collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getfirst-member(1))|Obtient le premier tableau de cette collection.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitem-member(1))|Obtient un tableau à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-items-member)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#excel-excel-textframe-autosizesetting-member)|Paramètres de resserrage automatique pour le cadre de texte.|
||[bottomMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-bottommargin-member)|Représente la marge bas, en points du cadre du texte.|
||[deleteText()](/javascript/api/excel/excel.textframe#excel-excel-textframe-deletetext-member(1))|Supprime tout le texte dans la textframe.|
||[hasText](/javascript/api/excel/excel.textframe#excel-excel-textframe-hastext-member)|Spécifie si le cadre de texte contient du texte.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontalalignment-member)|Représente l’alignement horizontal pour le style.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontaloverflow-member)|Représente le type de débordement horizontal du cadre du texte.|
||[leftMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-leftmargin-member)|Représente la marge gauche, en points du cadre du texte.|
||[Orientation](/javascript/api/excel/excel.textframe#excel-excel-textframe-orientation-member)|Représente l’angle vers lequel le texte est orienté pour le cadre de texte.|
||[readingOrder](/javascript/api/excel/excel.textframe#excel-excel-textframe-readingorder-member)|Représente l’ordre de lecture du cadre texte gauche à droite ou de droite à gauche.|
||[rightMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-rightmargin-member)|Représente la marge droite, en points du cadre du texte.|
||[textRange](/javascript/api/excel/excel.textframe#excel-excel-textframe-textrange-member)|Représente le texte lié à une forme, en plus des propriétés et des méthodes de manipulation du texte.|
||[topMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-topmargin-member)|Représente la marge du haut, en points du cadre du texte.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticalalignment-member)|Représente l’alignement vertical pour le style.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticaloverflow-member)|Représente le type de débordement vertical du cadre du texte.|
|[TextRange](/javascript/api/excel/excel.textrange)|[police](/javascript/api/excel/excel.textrange#excel-excel-textrange-font-member)|Renvoie un `ShapeFont` objet qui représente les attributs de police de la plage de texte.|
||[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#excel-excel-textrange-getsubstring-member(1))|Renvoie un objet TextRange pour les caractères dans la plage de donnée.|
||[text](/javascript/api/excel/excel.textrange#excel-excel-textrange-text-member)|Représente le contenu de texte brut de la plage de texte.|
|[Workbook](/javascript/api/excel/excel.workbook)|[autoSave](/javascript/api/excel/excel.workbook#excel-excel-workbook-autosave-member)|Spécifie si le workbook est en mode d’auto-ave.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#excel-excel-workbook-calculationengineversion-member)|Renvoie un nombre sur la version de moteur de calcul Excel.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbook#excel-excel-workbook-chartdatapointtrack-member)|True si tous les graphiques dans le classeur suivent les points de données réelles auquel qu’il sont joints.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechart-member(1))|Obtient la feuille de calcul active du classeur.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechartornullobject-member(1))|Obtient la feuille de calcul active du classeur.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getisactivecollabsession-member(1))|Renvoie `true` si le manuel est modifié par plusieurs utilisateurs (par le biais de la co-édition).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedranges-member(1))|Obtient la ou les plage(s) sélectionnée(s) actuelle(s) dans le classeur.|
||[isDirty](/javascript/api/excel/excel.workbook#excel-excel-workbook-isdirty-member)|Indique si des modifications ont été apportées depuis le dernier enregistré du manuel.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member)|Se produit lorsque le paramètre AutoSave est modifié sur le workbook.|
||[previouslySaved](/javascript/api/excel/excel.workbook#excel-excel-workbook-previouslysaved-member)|Spécifie si le manuel a déjà été enregistré localement ou en ligne.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#excel-excel-workbook-useprecisionasdisplayed-member)|True si les calculs réalisés dans ce classeur utiliseront uniquement la précision des nombres tels qu’ils sont affichés. |
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#excel-excel-workbookautosavesettingchangedeventargs-type-member)|Obtient le type de l’événement.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[autoFilter](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-autofilter-member)|Représente l’objet `AutoFilter` de la feuille de calcul.|
||[enableCalculation](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-enablecalculation-member)|Détermine si Excel devez recalculer la feuille de calcul si nécessaire.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1))|Recherche toutes les occurrences `RangeAreas` de la chaîne donnée en fonction des critères spécifiés et les renvoie en tant qu’objet, comprenant une ou plusieurs plages rectangulaires.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findallornullobject-member(1))|Recherche toutes les occurrences `RangeAreas` de la chaîne donnée en fonction des critères spécifiés et les renvoie en tant qu’objet, comprenant une ou plusieurs plages rectangulaires.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1))|Obtient l’objet `RangeAreas` , qui représente un ou plusieurs blocs de plages rectangulaires, spécifiés par l’adresse ou le nom.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-horizontalpagebreaks-member)|Obtient la collection de saut de page horizontal pour la feuille de calcul.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|Se produit lorsque le filtre est modifié sur un tableau spécifique.|
||[pageLayout](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pagelayout-member)|Obtient l’objet `PageLayout` de la feuille de calcul.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-replaceall-member(1))|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
||[Formes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-shapes-member)|Renvoie une collection de tous les objets Forme sur la feuille de calcul.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-verticalpagebreaks-member)|Obtient la collection de saut de page vertical pour la feuille de calcul.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[détails](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-details-member)|Représente les informations sur les détails des changements.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member)|Se produit lorsqu’une feuille de calcul dans le classeur est modifiée.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member)|Se produit lorsqu’une feuille de calcul du manuel a un format modifié.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member)|Se produit lorsque la sélection change sur n’importe quelle feuille de calcul.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-address-member)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrange-member(1))|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrangeornullobject-member(1))|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-source-member)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-type-member)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-worksheetid-member)|Obtient l’ID de la feuille de calcul dans laquelle les données ont été modifiées.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-completematch-member)|Spécifie si la correspondance doit être complète ou partielle.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-matchcase-member)|Spécifie si la correspondance est sensible à la cas.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
