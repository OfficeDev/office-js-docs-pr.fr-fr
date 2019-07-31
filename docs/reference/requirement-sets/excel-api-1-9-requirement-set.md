---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,9
description: Détails sur l’ensemble de conditions requises ExcelApi 1,9
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1d7e16a6e0aca202798016c136dfc7e2188c44f0
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940849"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Nouveautés de l’API JavaScript pour Excel 1,9

Plus de 500 nouvelles API Excel ont été ajoutés avec l’ensemble de conditions requises 1.9. Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | Insertion, la position et format images, formes géométriques et zones de texte. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Filtre automatique](../../excel/excel-add-ins-worksheets.md#filter-data) | Ajouter des filtres à des plages. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Zones](../../excel/excel-add-ins-multiple-ranges.md) | Prise en charge de plages discontinues. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Cellules spéciales](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Obtenez les cellules contenant des dates, des commentaires ou des formules dans une plage. | [Plage](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Chercher](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | Recherchez des valeurs ou des formules dans une plage ou une feuille de calcul. | [Plage](/javascript/api/excel/excel.range#find-text--criteria-)[feuille de calcul](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copier et coller](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | Copier des formules, formats et valeurs d’une plage à l’autre. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calcul](../../excel/performance.md#suspend-calculation-temporarily) | Contrôle plus étroit sur le moteur de calcul Excel. | [Application](/javascript/api/excel/excel.application) |
| Nouveaux graphiques | Explorez nos nouveaux types de graphiques pris en charge : cartes, zone et valeur, en cascade, en rayons de soleil, pareto. et entonnoir. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | Nouvelles fonctionnalités avec les formats de plage. | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>Liste des API

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Renvoie la version du moteur de calcul Excel utilisée pour le dernier recalcul complet. En lecture seule.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Renvoie l’état de calcul de l’application. Pour plus d’informations, voir Excel.CalculationState. En lecture seule.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Capture d’écran des paramètres de calcul itératif.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Interrompt le calcul jusqu'à ce que la prochaine méthode « context.sync() » soit appelée.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[Appliquer (plage : plage \| chaîne, columnIndex ? : nombre, critères ? : Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Applique le filtre automatique à une plage. Ceci permet de filtrer la colonne si les critères de filtre de colonne et index sont spécifiés.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Efface les critères de filtre du filtre automatique.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Cette propriété renvoie un objet Range qui représente la plage sur laquelle s'applique le filtre automatique spécifié.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Cette propriété renvoie un objet Plage qui représente la plage sur laquelle s'applique le filtre automatique spécifié.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Tableau qui conserve tous les critères de filtre dans une plage filtrée. En lecture seule.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Indique si le filtre automatique est activé ou non. En lecture seule.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Indique si le filtre automatique comporte des critères de filtre. En lecture seule.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Applique l’objet Autofilter spécifié actuellement sur la plage.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Supprime le filtre automatique pour la plage.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Représente la `color` propriété d’une bordure simple.|
||[style](/javascript/api/excel/excel.cellborder#style)|Représente la `style` propriété d’une bordure simple.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|Représente la `tintAndShade` propriété d’une bordure simple.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Représente la `weight` propriété d’une bordure simple.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bas](/javascript/api/excel/excel.cellbordercollection#bottom)|Représente la `format.borders.bottom` propriété.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|Représente la `format.borders.diagonalDown` propriété.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|Représente la `format.borders.diagonalUp` propriété.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Représente la `format.borders.horizontal` propriété.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Représente la `format.borders.left` propriété.|
||[right](/javascript/api/excel/excel.cellbordercollection#right)|Représente la `format.borders.right` propriété.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Représente la `format.borders.top` propriété.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Représente la `format.borders.vertical` propriété.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[adresse](/javascript/api/excel/excel.cellproperties#address)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|Représente la `addressLocal` propriété.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Représente la `hidden` propriété.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Représente la `format.fill.color` propriété.|
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Représente la `format.fill.pattern` propriété.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Représente la `format.fill.patternColor` propriété.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Représente la `format.fill.patternTintAndShade` propriété.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Représente la `format.fill.tintAndShade` propriété.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Représente la `format.font.bold` propriété.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Représente la `format.font.color` propriété.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Représente la `format.font.italic` propriété.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Représente la `format.font.name` propriété.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Représente la `format.font.size` propriété.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Représente la `format.font.strikethrough` propriété.|
||[Subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Représente la `format.font.subscript` propriété.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Représente la `format.font.superscript` propriété.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|Représente la `format.font.tintAndShade` propriété.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Représente la `format.font.underline` propriété.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|Représente la `autoIndent` propriété.|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Représente la `borders` propriété.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Représente la `fill` propriété.|
||[police](/javascript/api/excel/excel.cellpropertiesformat#font)|Représente la `font` propriété.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|Représente la `horizontalAlignment` propriété.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|Représente la `indentLevel` propriété.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Représente la `protection` propriété.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|Représente la `readingOrder` propriété.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|Représente la `shrinkToFit` propriété.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|Représente la `textOrientation` propriété.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|Représente la `useStandardHeight` propriété.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|Représente la `useStandardWidth` propriété.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|Représente la `verticalAlignment` propriété.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Représente la `wrapText` propriété.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|Représente la `format.protection.formulaHidden` propriété.|
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Représente la `format.protection.locked` propriété.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Représente la valeur une fois modifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Représente la valeur avant qu’elle soit modifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Les cellules contenant une erreur renvoie la chaîne d’erreur.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Représente le type de valeur après avoir été modifié|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Représente le type de valeur avant d’avoir été modifié|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Active le graphique dans l’interface utilisateur Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Encapsule les options pour le graphique croisé dynamique. En lecture seule.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Renvoie ou définit le jeu de couleurs du graphique. Lecture/écriture.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Spécifie si la zone graphique du graphique possède des coins arrondis ou non. Lecture/écriture.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Représente si le format de nombre est lié aux cellules ou non. Si true, le format de numérotation est affiché dans les étiquettes lorsqu’il se transforme dans les cellules.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Spécifie si le débordement bin est activé dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Spécifie si le soupassement bin est activé dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Renvoie ou définit si le nombre de bin d’un histogramme ou un graphique de pareto. Lecture/écriture.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Renvoie ou définit la valeur du débordement bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Renvoie ou définit le type de bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Renvoie ou définit la valeur du soupassement bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Renvoie ou définit la valeur de largeur de bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Renvoie ou définit le type de calcul quartile d’un graphique zone et valeur. Lecture/écriture.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Spécifie si les points internes sont affichés dans un graphique zone et valeur ou non. Lecture/écriture.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Spécifie si la ligne moyenne est affichée dans un graphique zone et valeur ou non. Lecture/écriture.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Spécifie si le marqueur moyen est affiché dans un graphique zone et valeur ou non. Lecture/écriture.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Spécifie si les points externes sont affichés dans un graphique zone et valeur ou non. Lecture/écriture.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Valeur booléenne si le format numérique est lié aux cellules (de sorte que le format de nombre est modifié dans les étiquettes lorsqu'il est modifié dans les cellules).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Représente si le format de nombre est lié aux cellules ou non. Si true, le format de numérotation sera changé dans les étiquettes lorsqu’il se transforme dans les cellules.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Spécifie si les barres d’erreur ont une lettrine de style de fin ou non.|
||[inclure](/javascript/api/excel/excel.charterrorbars#include)|Spécifie les parties de barres d’erreur à inclure.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Spécifie le type de mise en forme de barres d’erreur.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Le type de plage marqué par des barres d’erreur.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Spécifie si les barres d’erreur sont affichées ou non.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Représente le format des lignes du graphique.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Renvoie ou de définition de stratégie d’étiquettes de carte série d’un graphique de carte région. Lecture/écriture.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Renvoie ou définit le niveau d’une série de graphique région carte. Lecture/écriture.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Renvoie ou définit le type de projection de séries d’un graphique de carte région. Lecture/écriture.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Spécifie si les boutons de champ d’axe sont affichés sur un graphique croisé dynamique ou non. La propriété ShowAxisFieldButtons correspond aux commandes « Afficher les boutons du champ Axe » sur la liste « Boutons de champ » de l’onglet « Analyse » qui est disponible quand un graphique croisé dynamique est sélectionné.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Spécifie si les boutons de champ de légende sont affichés sur un graphique croisé dynamique ou non.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Spécifie si les boutons de champ de filtre de rapport sont affichés sur un graphique croisé dynamique ou non.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Spécifie si les boutons de champ d’affichage de valeur sont affichés sur un graphique croisé dynamique ou non.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Peut être une valeur d’entier entre 0 (zéro) et 300 correspondant à un pourcentage de la taille par défaut. Cette propriété s'applique uniquement aux graphiques en bulles. Lecture/écriture.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Renvoie ou définit la couleur pour la valeur maximale d’une série de graphique région carte. Lecture/écriture.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Renvoie ou définit le type pour la valeur maximale d’une série de graphique région carte. Lecture/écriture.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Renvoie ou définit la valeur maximale d’une série de graphique région carte. Lecture/écriture.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Renvoie ou définit la couleur pour la valeur de point médian d’une série de graphique région carte. Lecture/écriture.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Renvoie ou définit le type pour la valeur de point médian d’une série de graphique région carte. Lecture/écriture.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Renvoie ou définit la couleur pour la valeur du milieu d’une série de graphique région carte. Lecture/écriture.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Renvoie ou définit la couleur pour la valeur minimale d’une série de graphique région carte. Lecture/écriture.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Renvoie ou définit le type de la valeur minimale d’une série de graphique région carte. Lecture/écriture.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Renvoie ou définit la valeur minimale d’une série de graphique région carte. Lecture/écriture.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Renvoie ou définit le style de gradient de carte série d’un graphique de carte région. Lecture/écriture.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Renvoie ou définit la couleur de remplissage de point de données négative dans une série. Lecture/écriture.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Renvoie ou définit la zone de stratégie d’étiquettes de séries parents d’un graphique de compartimentage. Lecture/écriture.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Encapsule les options bin uniquement pour les histogrammes et graphiques de pareto. En lecture seule.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Résume les options pour le graphique de zone et valeur. En lecture seule.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Encapsule les options pour le graphique carte de région. En lecture seule.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Représente l’objet de la barre d’erreur pour une série de graphique.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Représente l’objet de la barre d’erreur pour une série de graphique.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Spécifie si les lignes de connexion apparaissent dans les graphiques en cascade ou non. Lecture/écriture.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Spécifie si des lignes d’étiquettes sont affichées à chaque étiquette de données de la série ou non. Lecture/écriture.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Cette propriété renvoie ou définit le seuil de la valeur séparant les deux sections d'un graphique en secteurs ou d'un graphique en barres de secteur. Lecture/écriture.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Valeur booléenne si le format numérique est lié aux cellules (de sorte que le format de nombre est modifié dans les étiquettes lorsqu'il est modifié dans les cellules).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[adresse](/javascript/api/excel/excel.columnproperties#address)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Représente la `addressLocal` propriété.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Représente la `columnIndex` propriété.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Renvoie le RangeAreas comprenant une ou plusieurs plages rectangulaires, le format conditionnel est appliqué. En lecture seule.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Renvoie un RangeAreas comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valide. Si toutes les valeurs de cellule sont valides, cette fonction génère une erreur ItemNotFound.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Renvoie un RangeAreas comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valide. Si toutes les valeurs de cellule sont valides, cette fonction renverra une valeur null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|La propriété utilisée par le filtre pour faire filtre enrichi sur richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Représente l’identificateur de forme. En lecture seule.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Renvoie l’objet de la forme de la forme géométrique. En lecture seule.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Renvoie le nombre de formes dans le groupe de la forme. En lecture seule.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Extrait un graphique à l’aide de son Nom ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Obtient ou définit le pied de page du centre de la feuille de calcul.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Obtient ou définit l’en-tête de page du centre de la feuille de calcul.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Obtient ou définit le pied de page gauche du centre de la feuille de calcul.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Obtient ou définit l’en-tête gauche du centre de la feuille de calcul.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Obtient ou définit le pied de page droit du centre de la feuille de calcul.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Obtient ou définit l’en-tête droit du centre de la feuille de calcul.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|L’en-tête/pied de page, utilisé pour toutes les pages, sauf si la première page ou page impaire/paire est spécifiée.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|L’en-tête/le pied de page à utiliser pour les pages paires, en-tête/pied de page impaire doit être spécifié pour les pages impaires.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|La première en-tête/le premier pied de page, pour toutes les autres pages générales ou impair/pair est utilisé.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|L’en-tête/le pied de page à utiliser pour les pages paires, l’en-tête/pied de page paire doit être spécifié pour les pages paires.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Obtient ou définit l’état des en-têtes/pieds de page qui sont définis. Pour plus d’informations, voir Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Obtient ou définit un indicateur indiquant si les en-têtes/pieds de page sont alignés avec les marges de page définis dans les options de mise en page pour la feuille de calcul.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Obtient ou définit un indicateur indiquant si les en-têtes/pieds de page sont alignés avec les marges de page définis dans les options de mise en page pour la feuille de calcul.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Renvoie le format de l’image. En lecture seule.|
||[id](/javascript/api/excel/excel.image#id)|Représente l’identificateur de forme pour l’objet d’image. En lecture seule.|
||[shape](/javascript/api/excel/excel.image#shape)|Renvoie l’objet de la Forme associé à la l’image. En lecture seule.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Cette propriété a la valeur True si Microsoft Excel utilise l'itération pour résoudre des références circulaires.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Cette propriété renvoie ou définit l'écart maximal utilisé pour chaque itération pendant que Microsoft Excel résout des références circulaires. |
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Cette propriété renvoie ou définit le nombre maximal d'itérations que Microsoft Excel peut utiliser pour résoudre une référence circulaire. |
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|Renvoie ou définit la longueur de la pointe de la flèche au début de la ligne spécifiée.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|Représente le style de la pointe de la flèche au début de la ligne spécifiée.|
||[beginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|Représente la largeur de la pointe de la flèche au début de la ligne spécifiée.|
||[connectBeginShape (forme : Excel.Shape, connectionSite : nombre)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|Détache le début du connecteur spécifié de la forme à laquelle il est attaché.|
||[connectBeginShape (forme : Excel.Shape, connectionSite : nombre)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|Joint la fin du connecteur spécifié à une forme spécifiée.|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|Représente le type de connecteur pour la ligne.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|Détache le début du connecteur spécifié de la forme à laquelle il est attaché.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|Détache la fin du connecteur spécifié de la forme à laquelle il est attaché.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|Représente la longueur de la pointe de la flèche à la fin de la ligne spécifiée.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|Représente le style de la pointe de la flèche à la fin de ligne spécifée.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|Représente la largeur de la pointe de la flèche à la fin de la ligne spécifiée.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Représente la forme de la pointe de la flèche au début de la ligne spécifiée. En lecture seule.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Représente le site de connexion indiquant le point de connexion auquel le début d'un connecteur est relié. En lecture seule. Renvoie la valeur null lorsque le début de la ligne n’est pas attaché à n’importe quelle forme.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Représente la forme de la pointe de la flèche à la fin de la ligne spécifiée. En lecture seule.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Représente le site de connexion indiquant le point de connexion auquel la fin d'un connecteur est relié. En lecture seule. Renvoie la valeur null lorsque la fin de la ligne n’est pas attaché à n’importe quelle forme.|
||[id](/javascript/api/excel/excel.line#id)|Représente l’identificateur de forme. En lecture seule.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Détermine si le début du connecteur spécifié est connecté à une forme. En lecture seule.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Spécifie si la fin du connecteur spécifié est connecté à une forme ou non. En lecture seule.|
||[shape](/javascript/api/excel/excel.line#shape)|Renvoie l’objet de la Forme associée à la ligne. En lecture seule.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Supprime un objet de saut de page.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Obtient la première cellule après le saut de page.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Représente l’index de colonne pour le saut de page|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Représente l’index de la rangée pour le saut de page|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[Ajouter (pageBreakRange : plage \| chaîne)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Ajoute un saut de page avant la cellule en haut à gauche de la plage spécifiée.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Obtient le nombre de pages de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Obtient un objet de saut de page via l’index.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Redéfinit tous les sauts de page de la collection.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Obtient ou définit l’option d’impression noir et blanc de la feuille de calcul.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Obtient ou définit la marge de page en bas de la feuille de calcul à utiliser pour l’impression en points.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Obtient ou définit l’indicateur horizontal du centre de la feuille de calcul. Cet indicateur détermine si la feuille de calcul est centrée horizontalement lorsqu’elle est imprimée.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Obtient ou définit l’indicateur vertical du centre de la feuille de calcul. Cet indicateur détermine si la feuille de calcul est centrée verticalement lorsqu’elle est imprimée.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul. Si true, la feuille sera imprimée sans graphismes.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|Obtient ou définit le premier numéro de page de la feuille de calcul à imprimer. La valeur NULL représente la numérotation des pages « automatique ».|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Obtient ou définit la marge de pied de page de la feuille de calcul, en points usage lors de l’impression.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages rectangulaires qui représente la zone d’impression pour la feuille de calcul. S’il n’existe aucune zone d’impression, une erreur ItemNotFound sera levée.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages rectangulaires qui représente la zone d’impression pour la feuille de calcul. S’il n’existe aucune zone d’impression, un objet null est renvoyé.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Obtient l’objet plage représentant les colonnes de titre.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Obtient l’objet plage représentant les colonnes de titre. Si ce n’est pas ensemble, un objet null est renvoyé.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Obtient l’objet plage représentant les rangées de titre.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Obtient l’objet plage représentant les rangées de titre. Si ce n’est pas ensemble, un objet null est renvoyé.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Obtient ou définit la marge de l’en-tête de la feuille de calcul, en points usage lors de l’impression.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Obtient ou définit la marge gauche de la feuille de calcul, en points usage lors de l’impression.|
||[Orientation](/javascript/api/excel/excel.pagelayout#orientation)|Obtient ou définit l’orientation de la feuille de calcul de la page.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Obtient ou définit l’orientation de la feuille de calcul de la page.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Obtient ou définit si les commentaires de la feuille de calcul doivent être affichées lors de l’impression.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Obtient ou définit l’option d’erreurs d’impression de la feuille de calcul.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Obtient ou définit l’indicateur de quadrillage de la feuille de calcul. Cet indicateur détermine si le quadrillage est imprimé ou non.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Obtient ou définit l’indicateur d’en-tête de la feuille de calcul. Cet indicateur détermine si les titres seront imprimés.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Obtient ou définit l’option de commande d’impression de la feuille de calcul. Cela indique l’ordre à utiliser pour traiter le numéro de page imprimé.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Configuration de l’en-tête et pied de page de la feuille de calcul.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Obtient ou définit la marge droite de la feuille de calcul, en points pour un usage lors de l’impression.|
||[setPrintArea (printArea : plage \| RangeAreas \| chaîne)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Définit la zone d’impression de la feuille de calcul.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Définit les marges de page de la feuille de calcul avec des unités.|
||[setPrintTitleColumns (printTitleColumns : plage \| chaîne)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Définit les colonnes qui contiennent des cellules répétées à gauche de chaque page de la feuille de calcul pour l’impression.|
||[setPrintTitleRows (printTitleRows : plage \| chaîne)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Définit les rangées qui contiennent des cellules répétées en haut de chaque page de la feuille de calcul pour l’impression.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Obtient ou définit la marge de pied de page de la feuille de calcul, en points pour un usage lors de l’impression.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Obtient ou définit les options de zoom de la feuille de calcul.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bas](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Représente la marge de bas de mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Représente le pied de page de la mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Représente l’en-tête de la mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Représente la marge gauche de la mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[right](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Représente la marge droite de la mise en page dans l’unité spécifiée à utiliser pour l’impression.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Représente la marge du haut de la mise en page dans l’unité spécifiée à utiliser pour l’impression.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Nombre de pages pour l’ajuster horizontalement. Cette valeur peut être null si l’échelle de pourcentage est utilisée.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|La valeur d’échelle de page d’impression peut être comprise entre 10 et 400. Cette valeur peut être null si la conformité à la page en hauteur ou largeur est spécifiée.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Nombre de pages pour l’ajuster verticalement. Cette valeur peut être null si l’échelle de pourcentage est utilisée.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Trie le PivotField par valeurs spécifiées dans une étendue donnée. L’étendue définit les valeurs spécifiques permettant de trier quand|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Spécifie si la mise en forme sera automatiquement formatés lorsqu’elle est actualisée ou lorsque les champs sont déplacés.|
||[getDataHierarchy (cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Obtient DataHierarchy servant à calculer la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[getPivotItems (axe : Excel.PivotAxis, cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtient le PivotItems à partir d’un axe qui composent la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Spécifie si la mise en forme est conservée lorsque le rapport est actualisé ou recalculé par des opérations telles que par glissement, tri ou en modifiant des éléments de champ de page.|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Définit le tableau croisé dynamique pour trier automatiquement à l’aide de la cellule spécifiée pour sélectionner automatiquement tous les critères et contexte nécessaires. Cela se comporte de la même manière que d’appliquer un tri automatique de l’interface utilisateur.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Spécifie si le tableau croisé dynamique autorise les valeurs dans le corps de données modifié par l’utilisateur.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Spécifie si le tableau croisé dynamique utilise des listes personnalisées lors du tri.|
|[Range](/javascript/api/excel/excel.range)|[recopie incrémentée (destinationRange : plage \| chaîne, autoFillType ? : Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Remplit la plage de la plage active à la plage de destination.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Convertit la plage de cellules avec des types de données en texte.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Convertit la plage de cellules en type de données liée dans la feuille de calcul.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copie les cellules de données ou de mise en forme à partir de la plage source ou RangeAreas à la plage active.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Apporte un remplissage instantané étendue en cours. Le remplissage instantané renseignera automatiquement les données lorsqu’il détectera un modèle, la plage doit donc être la seule plage de la colonne et avoir des données autour afin de trouver le modèle.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Renvoie une plage en 2D, qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Renvoie une plage à dimension unique, qui comprend les données de char colonne de police, de remplissage, de bordures, d’alignement, etc. de la plage.  Pour les propriétés ne sont pas cohérentes au sein de chaque cellule dans une colonne donnée, null est renvoyé.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Renvoie une plage à dimension unique , qui comprend les données de police, de remplissage, de bordures, d’alignement, etc. de la plage.  Pour les propriétés ne sont pas cohérentes au sein de chaque cellule dans une rangée donnée, la valeur null est renvoyée.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages rectangulaires qui représente toutes les cellules qui correspondent au type et la valeur spécifiés.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages qui représente les cellules qui correspondent au type et à la valeur spécifiés.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Obtient une collection de tableaux qui se chevauchent avec la plage dans l’étendue.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Représente l’état du type de données de chaque cellule. En lecture seule.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Supprime les valeurs dupliquées de la plage spécifiée par les colonnes.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Met à jour la plage basée sur une matrice 2D des propriétés de la cellule, résumant les éléments tels que la police, remplissage, bordures, alignement et ainsi de suite.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Met à jour la plage basée sur une matrice à une dimension des propriétés de la cellule, résumant les éléments tels que la police, remplissage, bordures, alignement et ainsi de suite.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Cette méthode désigne une plage qui doit être recalculée lorsque le recalcul suivant se produit.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Met à jour la plage basée sur une matrice à une dimension des propriétés de la cellule, résumant les éléments tels que la police, remplissage, bordures, alignement et ainsi de suite.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Calcule toutes les cellules de la RangeAreas.|
||[Effacer (applyTo ? : Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Efface les valeurs, format, remplissage, bordure, etc. sur chacune des zones qui composent cet objet RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Convertit toutes les cellules de RangeAreas avec des types de données en texte.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Convertit toutes les cellules de RangeAreas avec des types de données en texte.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copie les cellules de données ou de mise en forme à partir de la plage source ou RangeAreas à la plage active.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Renvoie un objet qui représente la colonne entière de la RangeAreas (par exemple, si la RangeAreas actuelle représente les cellules «B4:E11, H2 », elle renvoie une plage RangeAreas qui représente les colonnes « B:E, H:H»).|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Renvoie un objet RangeAreas qui représente la colonne entière de la RangeAreas (par exemple, si la RangeAreas actuelle représente les cellules «B4:E11 », elle renvoie une RangeAreas qui représente les rangées « 4:11»).|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Obtient l’objet de plage qui représente l’intersection des plages données ou RangeAreas. Si aucune intersection n’est trouvée, une erreur ItemNotFound sera levée.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Obtient l’objet de plage qui représente l’intersection des plages données ou RangeAreas. Si aucune intersection n’est trouvée, renvoie un objet null.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Renvoie un objet RangeAreas est décalé vers le décalage de lignes et des colonnes spécifiques. Les dimensions de la plage renvoyée correspondent à cette plage. Si la plage obtenue se retrouve en dehors des limites de grille de la feuille de calcul, une erreur est déclenchée.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Renvoie un objet RangeAreas qui représente toutes les cellules correspondant au type et à la valeur spécifiés. Lève une erreur si aucune cellule spéciale n’est trouvée qui corresponde au critère.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Renvoie un objet RangeAreas qui représente toutes les cellules correspondant au type et à la valeur spécifiés. Lève un objet null si aucune cellule spéciale n’est trouvée qui corresponde au critère.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Renvoie une collection de tableaux qui se chevauchent avec n’importe quelle plage dans cet objet RangeAreas dans l’étendue.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Renvoie les RangeAreas utilisés comprenant tous les domaines utilisés du individuelles et des plages dans l’objet RangeAreas rectangulaires.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Renvoie les RangeAreas utilisés comprenant tous les domaines utilisés du individuelles et des plages dans l’objet RangeAreas rectangulaires.|
||[adresse](/javascript/api/excel/excel.rangeareas#address)|Renvoie la référence RageAreas dans le style A1. La valeur de l’adresse contient le nom de feuille de calcul pour chaque bloc rectangulaire de cellules (par exemple, « Feuil1 ! A1 : B4, Feuil1 ! D1:D4 »). En lecture seule.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Renvoie la référence RageAreas dans l’utilisateur local. En lecture seule.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Renvoie le nombre de plages rectangulaires qui composent cet objet RangeArea.|
||[Zones](/javascript/api/excel/excel.rangeareas#areas)|Renvoie une collection de plages rectangulaires qui composent cet objet RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Renvoie le nombre de cellules dans l’objet RangeAreas récapitulatif du nombre de cellule de toutes les plages individuelles rectangulaires. Renvoie -1 si le nombre de cellules est supérieure à 2 ^ 31-1 (2 147 483 647). En lecture seule.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Renvoie un ensemble de ConditionalFormats qui se coupent pas avec cet objet RangeAreas toutes les cellules. En lecture seule.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Renvoie un objet dataValidation pour toutes les plages dans la RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Renvoie un objet de format rangeFormat, qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, et autres propriétés de toutes les plages dans l’objet RangeAreas. En lecture seule.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Indique si toutes les plages cet objet RangeAreas représentent des colonnes entières (par exemple, « A:C, Q:Z »). En lecture seule.|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Indique si toutes les plages cet objet RangeAreas représentent des colonnes entières (par exemple, « 1:3, 5:7 »). En lecture seule.|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Renvoie la feuille de calcul RangeAreas actuelle. En lecture seule.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Cette méthode désigne une plage RangeAreas qui doit être recalculée lorsque le recalcul suivant se produit.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Représente le style pour toutes les plages dans cet objet RangeAreas.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Renvoie ou définit un double qui s’éclaire ou assombrit une couleur de bordure de la plage, la valeur est comprise entre -1 (zones les plus sombres) et 1 (plus clair) avec 0 pour la couleur d’origine.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Renvoie ou définit un double qui s’éclaire ou assombrit une couleur de bordure de la plage, la valeur est comprise entre -1 (zones les plus sombres) et 1 (plus clair) avec 0 pour la couleur d’origine.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Renvoie le nombre de pages de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Renvoie la plage d’objet selon sa position dans la RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[pattern](/javascript/api/excel/excel.rangefill#pattern)|Obtient ou définit le modèle d’une plage. Pour plus d’informations, voir Excel.FillPattern. LinearGradient et RectangularGradient ne sont pas pris en charge.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Code couleur HTML qui représente la couleur de la ligne Axe, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Renvoie ou définit un double qui s’éclaire ou assombrit une couleur de bordure de la plage, la valeur est comprise entre -1 (zones les plus sombres) et 1 (plus clair) avec 0 pour la couleur d’origine.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Renvoie ou définit un double qui s’éclaire ou assombrit une couleur de bordure de la plage, la valeur est comprise entre -1 (zones les plus sombres) et 1 (plus clair) avec 0 pour la couleur d’origine.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Représente l’état barré de la police. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|
||[Subscript](/javascript/api/excel/excel.rangefont#subscript)|Représente le format de police indice.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Représente le format de police exposant.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Renvoie ou définit un double qui s’éclaire ou assombrit une couleur de bordure de la plage, la valeur est comprise entre -1 (zones les plus sombres) et 1 (plus clair) avec 0 pour la couleur d’origine.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Indique si le texte est automatiquement mis en retrait lorsque l’alignement du texte dans une cellule est défini sur distribution égale.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Entier compris entre 0 à 250 qui indique le niveau de retrait du style.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|L’ordre de lecture de la plage.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Indique si le texte s’ajuste automatiquement pour tenir dans la largeur de colonne disponible.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Nombre de lignes dupliquées supprimées par l’opération.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Nombre de lignes uniques restantes présents dans la plage qui en résulte.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Spécifie si la correspondance doit être complète ou partielle. La valeur par défaut est False (partielle).|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Spécifie si la correspondance respecte ou non la casse. Par défaut est false (ou minuscules).|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[adresse](/javascript/api/excel/excel.rowproperties#address)|Représente la `address` propriété.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Représente la `addressLocal` propriété.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Représente la `rowIndex` propriété.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Spécifie si la correspondance doit être complète ou partielle. Une correspondance complète correspond à la totalité du contenu de la cellule. La valeur par défaut est False (partielle).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Spécifie si la correspondance respecte ou non la casse. Par défaut est false (ou minuscules).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Détermine le sens de la recherche. Par défaut est transférer. Voir Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Représente la `format` propriété.|
||[lien hypertexte](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Représente la `hyperlink` propriété.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Représente la `style` propriété.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Représente la `columnHidden` propriété.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[format: Excel. CellPropertiesFormat & {
            columnWidth?] (format/JavaScript/API/Excel/Excel.settablecolumnproperties #)|Représente la `format` propriété.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel. CellPropertiesFormat & {
            rowHeight?] (format/JavaScript/API/Excel/Excel.settablerowproperties #)|Représente la `format` propriété.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Représente la `rowHidden` propriété.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Renvoie ou définit le texte de description de remplacement d’un objet de forme.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Renvoie ou définit le texte de titre de remplacement pour un objet de Forme.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Supprime la forme à partir de la feuille de calcul.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Représente le type de forme géométrique de cette forme géométrique. Voir Excel.GeometricShapeType pour les détails. Renvoie la valeur null si le type de forme n’est pas "GeometricShape".|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Convertit la forme à une image et renvoie l’image comme une chaîne codée en base 64. La résolution est 96. Les formats pris en charge uniquement sont `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, et `Excel.PictureFormat.GIF`.|
||[height](/javascript/api/excel/excel.shape#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Déplace horizontalement la forme spécifiée selon le nombre de points indiqué.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Fait pivoter la forme spécifiée dans le sens des aiguilles d’une montre, selon le nombre de degrés spécifié, autour de l'axe z.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Décale vers le haut la forme spécifiée selon le nombre de points spécifié.|
||[left](/javascript/api/excel/excel.shape#left)|La distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Spécifie ou non si la proportion d’aspect de cette forme est verrouillée.|
||[name](/javascript/api/excel/excel.shape#name)|Représente le nom de la forme.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Renvoie le nombre de sites de connexion sur la forme spécifiée. En lecture seule.|
||[fill](/javascript/api/excel/excel.shape#fill)|Renvoie la mise en forme de remplissage de cette forme. En lecture seule.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Renvoie la Forme géométrique associée à la forme. Une erreur sera lancée si le type de forme n’est pas "GeometricShape".|
||[groupe](/javascript/api/excel/excel.shape#group)|Renvoie le groupe de la Forme associée à la forme. Une erreur sera lancée si le type de forme n’est pas "GroupShape".|
||[id](/javascript/api/excel/excel.shape#id)|Représente l’identificateur de forme. En lecture seule.|
||[image](/javascript/api/excel/excel.shape#image)|Renvoie l’image associé à la forme. Une erreur sera lancée si le type de forme n’est pas "Image".|
||[level](/javascript/api/excel/excel.shape#level)|Représente le titre de la forme spécifiée. Par exemple, un niveau de 0 signifie que la forme ne fait pas partie d’un groupe, un niveau de la forme 1 signifie fait partie d’un groupe de niveau supérieur et un niveau de 2, la forme fait partie d’un groupe sous-blocs de niveau supérieur.|
||[line](/javascript/api/excel/excel.shape#line)|Renvoie l’image associée à la forme. Une erreur sera lancée si le type de forme n’est pas "Ligne".|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Renvoie la mise en forme de ligne de cette forme. En lecture seule.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Se produit lorsque la forme est activée.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Se produit lorsque la forme est désactivée.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Représente le groupe parent de la forme spécifiée.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Renvoie l’objet textFrame d’une forme. En lecture seule.|
||[type](/javascript/api/excel/excel.shape#type)|Renvoie le type de cette forme. Voir Excel.GeometricShapeType des détails. En lecture seule.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Renvoie la position de la forme spécifiée dans l’ordre z, valeur z de commande de la forme tout en bas est égal à 0. En lecture seule.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Représente la rotation en degrés, de la forme.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Met la hauteur de la forme à l’échelle en utilisant un facteur spécifié. Pour des images, vous pouvez indiquer si vous souhaitez mettre la forme à l’échelle par rapport à la taille d’origine ou la taille actuelle. Les formes autres que des images sont toujours mis à l’échelle par rapport à leur hauteur actuelle.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Met la largeur de la forme à l’échelle en utilisant un facteur spécifié. Pour des images, vous pouvez indiquer si vous souhaitez mettre la forme à l’échelle par rapport à la taille d’origine ou la taille actuelle. Les formes autres que des images sont toujours mis à l’échelle par rapport à leur largeur actuelle.|
||[setZOrder(value: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Déplace la forme spécifiée vers le haut ou vers le bas z de commande de la collection qui décale devant ou derrière les autres formes.|
||[top](/javascript/api/excel/excel.shape#top)|La distance, en points, du bord supérieur de l’objet au bord supérieur de la feuille de calcul.|
||[visible](/javascript/api/excel/excel.shape#visible)|Représente la visibilité de cette forme.|
||[width](/javascript/api/excel/excel.shape#width)|Représente la largeur, en points, de la forme.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Obtient l’id de la forme activée.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la forme est activée.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Ajoute une forme géométrique à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addGroup (valeurs : matrice < chaîne \| forme >)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Groupes un sous-ensemble de formes dans la feuille de calcul de cette collection de sites. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Crée une image à partir d’une chaîne en base 64 et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Ajoute une ligne à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Ajoute une zone de texte à la feuille de calcul avec le texte fourni en tant que le contenu. Elle renvoie un objet Shape qui représente la nouvelle zone de texte.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Renvoie le nombre de graphiques dans la feuille de calcul. En lecture seule.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Extrait un graphique à l’aide de son Nom ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Obtient l’id de la forme qui est désactivée.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la forme est désactivée.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Renvoie la mise en forme de remplissage de cette forme.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[type](/javascript/api/excel/excel.shapefill#type)|Renvoie le type de remplissage de la forme. En lecture seule. Voir Excel.GeometricShapeType des détails.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Définit le format de remplissage d’un élément de graphique sur une couleur unie. Cette opération modifie le type de remplissage à « Unie ».|
||[Transparency](/javascript/api/excel/excel.shapefill#transparency)|Renvoie ou définit le degré de transparence du remplissage spécifié sous la forme d’une valeur comprise entre 0,0 (opaque) et 1,0 (transparent). Renvoie null si le type de forme ne prend pas en charge transparence ou le remplissage de forme présente incohérente transparence, par exemple, avec un type de remplissage dégradé.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Représente le format de police Gras. Retourne la valeur null le TextRange inclut les deux fragments de texte en gras et en non.|
||[color](/javascript/api/excel/excel.shapefont#color)|Représentation sous forme de code couleur HTML de la couleur du texte(par exemple, #FF0000 représente le rouge). Renvoie la valeur null si le TextRange inclut les fragments de texte avec les différentes couleurs.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Représente le format de police Italique. Renvoie null si le TextRange inclut les deux fragments de texte en italique et non italique.|
||[name](/javascript/api/excel/excel.shapefont#name)|Représente le nom de la police (par exemple « Calibri ») Si le texte est un langage de Script complexe ou Asie de l’est, représente le nom de la police correspondante ; dans le cas contraire représente nom de police de caractères latins.|
||[size](/javascript/api/excel/excel.shapefont#size)|Représente la taille de police en points (par exemple, 11). Renvoie la valeur null si le TextRange inclut les fragments de texte avec les différentes couleurs.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Type de soulignement appliqué à la police. Renvoie la valeur null si le TextRange inclut les fragments de texte avec les différentes styles. Pour plus d’informations, voir Excel.ShapeFontUnderlineStyle.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Représente l’identificateur de forme. En lecture seule.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Renvoie l’objet de la Forme associée au groupe. En lecture seule.|
||[Formes](/javascript/api/excel/excel.shapegroup#shapes)|Renvoie la collection d’objets de forme. En lecture seule.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Dissocie toutes les formes groupées dans la forme spécifiée.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Représente le style de trait de la forme. Renvoie null lors de la ligne n’est pas visible ou il existe des stylets tiret incohérents. Pour plus d’informations, voir Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Représente le style de trait de la forme. Renvoie null lors de la ligne n’est pas visible ou il existe des stylets tiret incohérents. Pour plus d’informations, voir Excel.ShapeFontUnderlineStyle.|
||[Transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Renvoie ou définit le degré de transparence du remplissage spécifié sous la forme d’une valeur comprise entre 0,0 (opaque) et 1,0 (transparent). Renvoie la valeur null lorsque la forme a des transparences incohérentes.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Représente si la mise en forme de la ligne d’un élément de forme est visible. Renvoie la valeur null lorsque la forme a des visibilités incohérentes.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Représente l’épaisseur de ligne, en points. Renvoie null lors de la ligne n’est pas visible ou il existe des poids de ligne incohérents.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Représente les sous-champs est le nom de la propriété cible d’une valeur enrichi effectuer le tri.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Obtient le nombre de tableaux de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Représente l’objet de filtre automatique de la table. En lecture seule.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Obtient l’ID du tableau. En lecture seule.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est ajouté.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[Détails](/javascript/api/excel/excel.tablechangedeventargs#details)|Représente des informations sur les détails de modification|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Se produit lorsque la nouvelle table est ajoutée dans un classeur.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Se produit lorsque le tableau spécifié est supprimé dans un classeur.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Spécifie la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Obtient l’ID du tableau. En lecture seule.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Spécifie le nom du tableau qui est supprimé.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Spécifie le type du champ. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est ajouté.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Obtient le nombre de tableaux de la collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Obtient le premier tableau de cette collection. Les tables dans la collection de sont triées de haut en bas et gauche vers la droite par ce tableau supérieure gauche afin que le premier tableau soit dans la collection de sites.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Obtient un tableau à l’aide de son nom ou de son ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|Obtient ou définit le paramètres de texte de dimensionnement automatique. Un bloc de texte peut être configuré pour ajuster automatiquement le texte pour le cadre du texte, pour ajuster automatiquement le bloc de texte au texte ou de ne pas effectuer tout problème de dimensionnement automatique.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Représente la marge bas, en points du cadre du texte.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Supprime tout le texte dans la textframe.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Représente l’alignement horizontal pour le style. Pour plus d’informations, voir Excel.ShapeTextHorizontalAlignment.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Représente le type de débordement horizontal du cadre du texte. Pour plus d’informations, voir Excel.ShapeTextHorizontalOverflow. |
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Représente la marge gauche, en points du cadre du texte.|
||[Orientation](/javascript/api/excel/excel.textframe#orientation)|Représente l’orientation du texte de l’encadrement de texte. Pour plus d’informations, voir Excel.ShapeTextOrientation.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Représente l’ordre de lecture du cadre texte gauche à droite ou de droite à gauche. Pour plus d’informations, voir Excel.ShapeTextReadingOrder.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Spécifie si la TextFrame contient du texte.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Représente le texte lié à une forme, en plus des propriétés et des méthodes de manipulation du texte. Pour plus d’informations, voir Excel.TextRange.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Représente la marge droite, en points du cadre du texte.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Représente la marge du haut, en points du cadre du texte.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Représente l’alignement vertical pour le style. Pour plus d’informations, voir Excel.ShapeTextVerticalAlignment.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Représente le type de débordement vertical du cadre du texte. Pour plus d’informations, voir Excel.ShapeTextVerticalOverflow.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Renvoie un objet TextRange pour les caractères dans la plage de donnée.|
||[police](/javascript/api/excel/excel.textrange#font)|Renvoie un objet ShapeFont qui représente les attributs de police pour la plage de texte. En lecture seule.|
||[text](/javascript/api/excel/excel.textrange#text)|Représente le contenu de texte brut de la plage de texte.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True si tous les graphiques dans le classeur suivent les points de données réelles auquel qu’il sont joints.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Obtient la feuille de calcul active du classeur. S’il n’existe aucun graphique actif, génère des exceptions lorsque appeler cette déclaration|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Obtient la feuille de calcul active du classeur. S’il n’existe aucun graphique actif, renverra l’objet null|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True si le classeur est modifié par plusieurs utilisateurs (co-création).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Obtient la ou les plage(s) sélectionnée(s) actuelle(s) dans le classeur. Contrairement aux getSelectedRange(), cette méthode renvoie un objet RangeAreas qui représente toutes les plages sélectionnées.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Spécifie ou non les modifications ont été apportées étant donné que le classeur a été enregistré.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Spécifie si le classeur est en mode enregistrement automatique. En lecture seule.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Renvoie un nombre sur la version de moteur de calcul Excel. En lecture seule.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Se produit lorsque le paramètre de l’enregistrement automatique est modifié dans le classeur.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Spécifie ou non si le classeur a jamais été enregistré localement ou en ligne. En lecture seule.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True si les calculs réalisés dans ce classeur utiliseront uniquement la précision des nombres tels qu’ils sont affichés. |
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[type](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Obtient ou définit EnableCalculation, propriété de la feuille de calcul.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Trouve toutes les occurrences de la chaîne donnée en fonction des critères spécifiées et renvoie un objet RangeAreas comprenant une ou plusieurs plages rectangulaires.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Trouve toutes les occurrences de la chaîne donnée en fonction des critères spécifiées et renvoie un objet RangeAreas comprenant une ou plusieurs plages rectangulaires.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Obtient l’objet RangeAreas représentant un ou plusieurs blocs de plages rectangulaires, spécifiés par nom ou l’adresse.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Représente l’objet AutoFilter de filtre automatique de la feuille de calcul. En lecture seule.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Obtient la collection de saut de page horizontal pour la feuille de calcul. Cette collection contient uniquement les sauts de page manuels.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Se produit lorsque le filtre est modifié sur un tableau spécifique.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Obtient l’objet PageLayout de la feuille de calcul.|
||[Formes](/javascript/api/excel/excel.worksheet#shapes)|Renvoie une collection de tous les objets Forme sur la feuille de calcul. En lecture seule.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Obtient la collection de saut de page vertical pour la feuille de calcul. Cette collection contient uniquement les sauts de page manuels.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[détails](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Représente des informations sur les détails de modification|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Se produit lorsqu’une feuille de calcul dans le classeur est modifiée.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Se produit lorsqu’une feuille de calcul dans le classeur a un format modifié.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Se produit lorsque la sélection change sur n’importe quelle feuille de calcul.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique. Il peut renvoyer un objet null.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Spécifie si la correspondance doit être complète ou partielle. Une correspondance complète correspond à la totalité du contenu de la cellule. La valeur par défaut est False (partielle).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Spécifie si la correspondance respecte ou non la casse. Par défaut est false (ou minuscules).|

## <a name="see-also"></a>Voir aussi

- [Documentation de référence de l’API JavaScript pour Excel](/javascript/api/excel)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
