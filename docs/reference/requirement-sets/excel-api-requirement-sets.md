---
title: Ensembles de conditions requises de l’API JavaScript pour Excel
description: ''
ms.date: 02/15/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 9985fabdf0c5e9e6c09cf490b55fffd7f87a195a
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199626"
---
# <a name="excel-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour Excel

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Versions d’Office et ensembles de conditions requises](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Les compléments Excel peuvent être exécutés dans différentes versions d’Office, notamment Office 2016 pour Windows, Office pour iPad, Office pour Mac et Office Online. Le tableau suivant répertorie les ensembles de conditions requises pour Excel, les applications hôtes Office qui prennent en charge chaque ensemble de conditions et les versions ou numéro de build de ces applications.

> [!NOTE]
> Pour utiliser l’API dans un des jeux exigence numérotée, vous devez référencer la **production** de la bibliothèque sur le CDN : https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Pour plus d’informations sur l’utilisation aperçu API, voir la section[JavaScript d’Excel preview API](#excel-javascript-preview-apis) dans cet article.

|  Ensemble de conditions requises  |  Office 365 pour Windows  |  Office 365 pour iPad  |  Office 365 pour Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Aperçu  | Veuillez utiliser la dernière version d’Office pour tester la préversion API (vous devrez peut-être rejoindre la [programme Office Insider](https://products.office.com/office-insider)) |
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

## <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis. Le tableau suivant répertorie les API actuellement disponibles en mode aperçu. Pour formuler des commentaires sur une version d’évaluation API, utilisez le mécanisme de commentaires à la fin de la page web où l’API est documenté.

> [!NOTE]
> L’aperçu API peut être modifiés et n’est pas destinés à utiliser dans un environnement de production. Nous vous recommandons de les tester uniquement dans les environnements de test et de développement. N’utilisez pas un aperçu d’API dans un environnement de production ou dans les documents commerciaux importants.
>
> Pour utiliser l’aperçu API, vous devez référencer la bibliothèque**bêta**sur le CDN : https://appsforoffice.microsoft.com/lib/beta/hosted/office.js et vous devrez également participer au programme Office Insider pour obtenir un build Office suffisamment récent.

Plus de 400 nouvelles API Excel sont actuellement dans l’aperçu. Le premier tableau fournit un résumé concis de l’API, tandis que le tableau suivant qui fournit une liste détaillée. Essayez les nouvelles fonctionnalités et partagez vos commentaires avec nous.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Segment | Insérer et configurez segments aux tableaux et tableaux croisés dynamiques. | [Segment](/javascript/api/excel/excel.slicer) |
| Commentaires | Ajouter, modifier et supprimer des listes. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Formes | Insertion, la position et format images, formes géométriques et zones de texte. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| Nouveaux graphiques | Explorez nos nouveaux types de graphiques pris en charge : cartes, zone et valeur, en cascade, en rayons de soleil, pareto. et entonnoir. | [Chart](/javascript/api/excel/excel.charttype) |
| Filtre automatique | Ajouter des filtres à des plages. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| Zones | Prise en charge de plages discontinues. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| Cellules spéciales | Obtenez les cellules contenant des dates, des commentaires ou des formules dans une plage. | [Plage](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| Rechercher | Recherchez des valeurs ou des formules dans une plage ou une feuille de calcul. | [Plage](/javascript/api/excel/excel.range#find-text--criteria-)[feuille de calcul](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| Copier coller | Copier des formules, formats et valeurs d’une plage à l’autre. | [Plage](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| RangeFormat | Nouvelles fonctionnalités avec les formats de plage. | [Plage](/javascript/api/excel/excel.rangeformat) |
| Classeur enregistrer et fermer. | Enregistrez et fermez ses classeurs.  | [Workbook](/javascript/api/excel/excel.workbook) |
| Insérer le classeur | Insérer un classeur dans un autre.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Calcul | Contrôle plus étroit sur le moteur de calcul Excel. | [Application](/javascript/api/excel/excel.application) |

Les informations suivantes sont une liste complète des fichiers texte.

| Classe | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Renvoie des informations sur la version de Microsoft Excel dans laquelle le classeur a été entièrement recalculé. Type de données Long en lecture seule. En lecture seule.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Renvoie un CalculationState qui indique l’état de calcul de l’application. Pour plus d’informations, voir Excel.CalculationState. En lecture seule.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Capture d’écran des paramètres de calcul itératif.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Interrompt le calcul jusqu'à ce que la prochaine méthode « context.sync() » soit appelée.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[Appliquer (plage : plage \| chaîne, columnIndex ? : nombre, critères ? : Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Applique le filtre automatique sur une plage et permet de filtrer la colonne en colonne indexer et filtrer critères sont spécifiés.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Efface les critères si le filtre automatique a filtres|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Cette propriété renvoie un objet Range qui représente la plage sur laquelle s'applique le filtre automatique spécifié.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|S’il existe un objet plage associé avec le filtre automatique, cette méthode le renvoie.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Tableau qui conserve tous les critères de filtre dans une plage filtrée automatiquement. En lecture seule.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Indique si le filtre automatique est activé ou non. En lecture seule.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Indique si le filtre automatique comporte des critères de filtre. En lecture seule.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Applique l’objet Autofilter spécifié actuellement sur la plage.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Supprime le filtre automatique pour la plage.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)||
||[style](/javascript/api/excel/excel.cellborder#style)||
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)||
||[weight](/javascript/api/excel/excel.cellborder#weight)||
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)||
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)||
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)||
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)||
||[left](/javascript/api/excel/excel.cellbordercollection#left)||
||[right](/javascript/api/excel/excel.cellbordercollection#right)||
||[top](/javascript/api/excel/excel.cellbordercollection#top)||
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)||
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[adresse](/javascript/api/excel/excel.cellproperties#address)||
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)||
||[hasSpill](/javascript/api/excel/excel.cellproperties#hasspill)||
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)||
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)||
||[pattern](/javascript/api/excel/excel.cellpropertiesfill#pattern)||
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)||
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)||
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)||
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)||
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)||
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)||
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)||
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)||
||[Subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)||
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)||
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)||
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)||
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)||
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)||
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)||
||[police](/javascript/api/excel/excel.cellpropertiesformat#font)||
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)||
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)||
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)||
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)||
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)||
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)||
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)||
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)||
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)||
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Crée et ouvre un nouveau classeur.  Si vous le souhaitez, le classeur peut être renseigné avec un fichier .xlsx codé en base 64.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)||
||[locked](/javascript/api/excel/excel.cellpropertiesprotection#locked)||
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Active la feuille de calcul dans l’interface utilisateur Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Encapsule les options pour le graphique croisé dynamique. En lecture seule.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Renvoie ou définit un entier qui représente le jeu de couleurs pour le graphique. Lecture/écriture.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|True si la zone graphique du graphique possède des coins arrondis. Type de données Boolean en lecture-écriture. Lecture/écriture.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|True si le format numérique est lié aux cellules (de sorte que le format de nombre est modifié dans les étiquettes lorsqu'il est modifié dans les cellules).|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Renvoie ou définit si le débordement bin activé dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Renvoie ou définit si le débordement bin est activé dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[count](/javascript/api/excel/excel.chartbinoptions#count)|Renvoie ou définit si le nombre de corbeille d’un histogramme ou un graphique de pareto. Lecture/écriture.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Renvoie ou définit la valeur du débordement bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Renvoie ou définit le type de débordement bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Renvoie ou définit la valeur du débordement bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Renvoie ou définit la valeur du débordement bin dans un histogramme ou un graphique de pareto. Lecture/écriture.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Renvoie ou définit le type de calcul quartile d’un graphique zone et valeur. Lecture/écriture.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Renvoie ou définit si les points internes sont affichés dans un graphique de zone et valeur. Lecture/écriture.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Renvoie ou définit si les lignes sont affichées dans un graphique de zone et valeur. Lecture/écriture.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Renvoie ou définit si les marqueurs sont affichés dans un graphique de zone et valeur. Lecture/écriture.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Renvoie ou définit si les points hors normes sont affichés dans un graphique de zone et valeur. Lecture/écriture.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Valeur booléenne si le format numérique est lié aux cellules (de sorte que le format de nombre est modifié dans les étiquettes lorsqu'il est modifié dans les cellules).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Représente si le format numérique est lié aux cellules (de sorte que le format de nombre est modifié dans les étiquettes lorsqu'il est modifié dans les cellules).|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Spécifie si vous disposez de la lettrine style fin pour barres d’erreur.|
||[inclure](/javascript/api/excel/excel.charterrorbars#include)|Spécifie les parties de la barre d'erreur à inclure. Pour plus d’informations, voir Excel.ChartErrorBarsInclude.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Représente le format du quadrillage de graphique.|
||[type](/javascript/api/excel/excel.charterrorbars#type)|Spécifie la plage marquée par des barres d'erreur. Pour plus d’informations, voir Excel.ChartErrorBarsInclude.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Représente si les barres d’erreur s’affichent.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Représente le format des lignes du graphique.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Renvoie ou de définition de stratégie d’étiquettes de carte série d’un graphique de carte région. Lecture/écriture.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Renvoie ou définit une zone de carte série d’un graphique de carte région. Lecture/écriture.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Renvoie ou définit le type de projection d’un graphique de carte région. Lecture/écriture.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique ou non.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Représente si vous voulez afficher tous les boutons de champ dans un graphique croisé dynamique.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Renvoie ou définit le facteur d’échelle des bulles dans le groupe graphique spécifié. Peut être une valeur d’entier entre 0 (zéro) et 300 correspondant à un pourcentage de la taille par défaut. S’applique uniquement aux graphiques en courbes. Lecture/écriture.|
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
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Renvoie ou définit la zone de stratégie séries parent étiquette d’un graphique de compartimentage. Lecture/écriture.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Encapsule les options bin uniquement pour les histogramme et graphique de pareto. En lecture seule.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Résume les options pour le graphique croisé de zone et valeur. En lecture seule.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Encapsule les options pour le graphique carte. En lecture seule.|
||[xerrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Représente l’objet de la barre d’erreur pour une série de graphique.|
||[yerrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Représente l’objet de la barre d’erreur pour une série de graphique.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Renvoie ou définit si les lignes de connexion s’affichent dans un graphique en cascade. Lecture/écriture.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|True si Microsoft Excel n’affiche pas leaderlines pour chaque datalabel de série. Lecture/écriture.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Cette propriété renvoie ou définit le seuil de la valeur séparant les deux sections d'un graphique en secteurs de secteur ou d'un graphique en barres de secteur. Lecture/écriture.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Valeur booléenne si le format numérique est lié aux cellules (de sorte que le format de nombre est modifié dans les étiquettes lorsqu'il est modifié dans les cellules).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[adresse](/javascript/api/excel/excel.columnproperties#address)||
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)||
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)||
||[hasSpill](/javascript/api/excel/excel.columnproperties#hasspill)||
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Obtenir/définir le contenu.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Supprime le thread de commentaires.|
||[id](/javascript/api/excel/excel.comment#id)|Représente l’identificateur de commentaire. En lecture seule.|
||[isParent](/javascript/api/excel/excel.comment#isparent)|Représente si cela est un fil de commentaires ou une réponse. Toujours retourner true ici. En lecture seule.|
||[Réponses](/javascript/api/excel/excel.comment#replies)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Crée un nouveau commentaire (fil de commentaires) basé sur les emplacements des cellules et le contenu. L’argument non valide est levé si l’emplacement est supérieur à une cellule.|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Crée un nouveau commentaire (fil de commentaires) basé sur les emplacements des cellules et le contenu. L’argument non valide est levé si l’emplacement est supérieur à une cellule.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Obtient le nombre de commentaires de la collection.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Renvoie un commentaire identifié via son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Obtient une colonne en fonction de sa position dans la collection.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Obtient un commentaire sur la cellule dans la collection de sites spécifique.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Obtient un commentaire lié à son ID dans la collection de réponse.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentcollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Obtenir/définir le contenu.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Supprime la réponse de commentaire.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Obtenir son commentaire parent de cette réponse.|
||[id](/javascript/api/excel/excel.commentreply#id)|Représente l’identificateur de réponse du commentaire. En lecture seule.|
||[isParent](/javascript/api/excel/excel.commentreply#isparent)|Représente si cela est un fil de commentaires ou une réponse. Toujours retourner false ici. En lecture seule.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Crée une réponse à un commentaire pour un commentaire.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Obtient le nombre de réponses aux commentaires de la collection.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Renvoie une réponse de commentaire identifié via son ID. En lecture seule.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Obtient une réponse de commentaire en fonction de sa position dans la collection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.commentreplycollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Renvoie le RangeAreas comprenant une ou plusieurs plages rectangulaires, le format conditionnel est appliqué. En lecture seule.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Renvoie un RangeAreas comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valide. Si toutes les valeurs de cellule sont valides, cette fonction génère une erreur ItemNotFound.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Renvoie un RangeAreas comprenant une ou plusieurs plages rectangulaires, avec des valeurs de cellule non valide. Si toutes les valeurs de cellule sont valides, cette fonction renverra une valeur null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|La propriété utilisée par le filtre pour faire filtre enrichi sur richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Représente l’identificateur de forme. En lecture seule.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Renvoie l’objet de la forme de la forme géométrique. En lecture seule.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Obtient le nombre de formes de la collection. En lecture seule.|
||[getItem(name: string)](/javascript/api/excel/excel.groupshapecollection#getitem-name-)|Extrait un graphique à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.groupshapecollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
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
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Format du corps renvoyé. En lecture seule.|
||[id](/javascript/api/excel/excel.image#id)|Représente l’identificateur de forme pour l’objet d’image. En lecture seule.|
||[shape](/javascript/api/excel/excel.image#shape)|Renvoie l’objet de la forme de l’image. En lecture seule.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|Cette propriété a la valeur True si Microsoft Excel utilise l'itération pour résoudre des références circulaires.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Cette propriété renvoie ou définit l'écart maximal utilisé pour chaque itération pendant que Microsoft Excel résout des références circulaires. |
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Cette propriété renvoie ou définit le nombre maximal d'itérations que Microsoft Excel peut utiliser pour résoudre une référence circulaire. |
|[Line](/javascript/api/excel/excel.line)|[connectorType](/javascript/api/excel/excel.line#connectortype)|Représente le type de connecteur pour la ligne.|
||[id](/javascript/api/excel/excel.line#id)|Représente l’identificateur de forme. En lecture seule.|
||[shape](/javascript/api/excel/excel.line#shape)|Renvoie l’objet de la forme de la ligne. En lecture seule.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[source](/javascript/api/excel/excel.listdatavalidation#source)|Source de la liste de validation des données|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Supprime un objet de saut de page.|
||[getStartCell()](/javascript/api/excel/excel.pagebreak#getstartcell--)|Obtient la première cellule après le saut de page.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Représente l’index de colonne pour le saut de page|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Représente l’index de la rangée pour le saut de page|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[Ajouter (pageBreakRange : plage \| chaîne)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Ajoute un saut de page avant la cellule en haut à gauche de la plage spécifiée.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Obtient le nombre de pages de la collection.|
||[getItem(index : numérique)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Obtient un objet de saut de page via l’index.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.pagebreakcollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
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
||[setPrintMargins (unité : « Points » \| « Pouces » \| « Centimètres », marginOptions : Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Définit les marges de page de la feuille de calcul avec des unités.|
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
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues (sortby : « Croissant » \| « Décroissant », valuesHierarchy : Excel.DataPivotHierarchy, pivotItemScope ? : matrice < PivotItem \| chaîne >)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Trie le PivotField par valeurs spécifiées dans une étendue donnée. L’étendue définit les valeurs spécifiques permettant de trier quand|
||[sortByValues (sortby : Excel.SortBy, valuesHierarchy : Excel.DataPivotHierarchy, pivotItemScope ? : matrice < PivotItem \| chaîne >)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Trie le PivotField par valeurs spécifiées dans une étendue donnée. L’étendue définit les valeurs spécifiques permettant de trier quand|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|True si la mise en forme sera automatiquement formaté lorsqu’il est actualisé ou lorsque les champs sont déplacés.|
||[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|True si la liste des champs dois être affichée ou masquée à partir de l’interface utilisateur.|
||[getCell (dataHierarchy : DataPivotHierarchy \| chaîne rowItems : matrice < PivotItem \| chaîne >, columnItems : matrice < PivotItem \| chaîne >)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Dispose de la cellule dans le corps de données de tableau croisé dynamique qui contient la valeur pour l’intersection des dataHierarchy spécifié, rowItems et columnItems.|
||[getDataHierarchy (cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Obtient DataHierarchy servant à calculer la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[getPivotItems (axe : « Inconnu » \| « Ligne » \| « Colonne » \| « Données » \| « Filtrer », la cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtient le PivotItems à partir d’un axe qui composent la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[getPivotItems (axe : Excel.PivotAxis, cellule : plage \| chaîne)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtient le PivotItems à partir d’un axe qui composent la valeur dans une plage spécifiée dans le tableau croisé dynamique.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|True si la mise en forme est conservée lorsque le rapport est actualisé ou recalculé par des opérations telles que par glissement, le tri ou en modifiant des éléments de champ de page.|
||[setAutosortOnCell (cellule : plage \| chaîne, sortby : « Croissant » \| « Décroissant »)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Définit un tri automatique à l’aide de la cellule spécifiée pour sélectionner automatiquement tous les critères et contexte pour le tri.|
||[setAutosortOnCell(cell: Range \| string, sortby: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Définit un tri automatique à l’aide de la cellule spécifiée pour sélectionner automatiquement tous les critères et contexte pour le tri.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|True si le tableau croisé dynamique doit utiliser des listes personnalisées lors du tri.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|True si le tableau croisé dynamique doit utiliser des listes personnalisées lors du tri.|
|[Plage](/javascript/api/excel/excel.range)|[recopie incrémentée (destinationRange : plage \| chaîne, autoFillType ? : « FillDefault » \| « FillCopy » \| « FillSeries » \| « FillFormats » \| « FillValues » \| « FillDays » \|« FillWeekdays » \| « FillMonths » \| « FillYears » \| « LinearTrend » \| « GrowthTrend » \| « Remplissage instantané »)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)||
||[recopie incrémentée (destinationRange : plage \| chaîne, autoFillType ? : Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)||
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Convertit la plage de cellules avec des types de données en texte.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Convertit la plage de cellules en type de données liée dans la feuille de calcul.|
||[copyFrom (sourceRange : plage \| RangeAreas \| chaîne, copyType ? : « Toute » \| « Formules » \| « Valeurs » \| « Formats », IgnorerVides ? : transpose booléenne, ? : booléen)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copie les cellules de données ou de mise en forme à partir de la plage source ou RangeAreas à la plage active.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copie les cellules de données ou de mise en forme à partir de la plage source ou RangeAreas à la plage active.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Recherche la chaîne donnée basée sur les critères spécifiés.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Renvoie une plage en 2D, qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Renvoie une plage à dimension unique, qui comprend les données de char colonne de police, de remplissage, de bordures, d’alignement, etc. de la plage.  Pour les propriétés ne sont pas cohérentes au sein de chaque cellule dans une colonne donnée, null est renvoyé.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Renvoie une plage à dimension unique , qui comprend les données de police, de remplissage, de bordures, d’alignement, etc. de la plage.  Pour les propriétés ne sont pas cohérentes au sein de chaque cellule dans une rangée donnée, la valeur null est renvoyée.|
||[getSpecialCells (cellType : « ConditionalFormats » \| « DataValidations » \| « Vides » \| « Commentaires » \| « Constantes » \| « Formules » \| « SameConditionalFormat » \|« SameDataValidation » \| cellValueType « Visible » ? : « Tout » \| « Erreurs » \| « ErrorsLogical » \| « ErrorsNumbers » \| « ErrorsText » \| » ErrorsLogicalNumber » \| « ErrorsLogicalText » \| « ErrorsNumberText » \| « Logique » \| « LogicalNumbers » \| « LogicalText » \| « LogicalNumbersText » \| « Nombres » \| « NumbersText » \| « Texte »)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages rectangulaires qui représente toutes les cellules qui correspondent au type et la valeur spécifiés.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages rectangulaires qui représente toutes les cellules qui correspondent au type et la valeur spécifiés.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Comments" \| "Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages rectangulaires qui représente les cellules qui correspondent au type et à la valeur spécifiés.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtient l’objet RangeAreas, comprenant une ou plusieurs plages rectangulaires qui représente les cellules qui correspondent au type et à la valeur spécifiés.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Obtient une collection de tableaux qui se chevauchent avec la plage dans l’étendue.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Représente l’état du type de données de chaque cellule. En lecture seule.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Supprime les valeurs dupliquées de la plage spécifiée par les colonnes.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Met à jour la plage basée sur une matrice 2D des propriétés de la cellule, résumant les éléments tels que la police, remplissage, bordures, alignement et ainsi de suite.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Met à jour la plage basée sur une matrice à une dimension des propriétés de la cellule, résumant les éléments tels que la police, remplissage, bordures, alignement et ainsi de suite.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Cette méthode désigne une plage qui doit être recalculée lorsque le recalcul suivant se produit.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Met à jour la plage basée sur une matrice à une dimension des propriétés de la cellule, résumant les éléments tels que la police, remplissage, bordures, alignement et ainsi de suite.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Calcule toutes les cellules de la RangeAreas.|
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Efface les valeurs, format, remplissage, bordure, etc. sur chacune des zones qui composent cet objet RangeAreas.|
||[Effacer (applyTo ? : Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Efface les valeurs, format, remplissage, bordure, etc. sur chacune des zones qui composent cet objet RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Convertit toutes les cellules de RangeAreas avec des types de données en texte.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Convertit toutes les cellules de RangeAreas avec des types de données en texte.|
||[copyFrom (sourceRange : plage \| RangeAreas \| chaîne, copyType ? : « Toute » \| « Formules » \| « Valeurs » \| « Formats », IgnorerVides ? : transpose booléenne, ? : booléen)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copie les cellules de données ou de mise en forme à partir de la plage source ou RangeAreas à la plage active.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copie les cellules de données ou de mise en forme à partir de la plage source ou RangeAreas à la plage active.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Renvoie un objet qui représente la colonne entière de la RangeAreas (par exemple, si la RangeAreas actuelle représente les cellules «B4:E11, H2 », elle renvoie une plage RangeAreas qui représente les colonnes « B:E, H:H»).|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Renvoie un objet RangeAreas qui représente la colonne entière de la RangeAreas (par exemple, si la RangeAreas actuelle représente les cellules «B4:E11 », elle renvoie une RangeAreas qui représente les rangées « 4:11»).|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Obtient l’objet de plage qui représente l’intersection des plages données ou RangeAreas. Si aucune intersection n’est trouvée, une erreur ItemNotFound sera levée.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Obtient l’objet de plage qui représente l’intersection des plages données ou RangeAreas. Si aucune intersection n’est trouvée, renvoie un objet null.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Renvoie un objet RangeAreas est décalé vers le décalage de lignes et des colonnes spécifiques. Les dimensions de la plage renvoyée correspondent à cette plage. Si la plage obtenue se retrouve en dehors des limites de grille de la feuille de calcul, une erreur est déclenchée.|
||[getSpecialCells (cellType : « ConditionalFormats » \| « DataValidations » \| « Vides » \| « Commentaires » \| « Constantes » \| « Formules » \| « SameConditionalFormat » \|« SameDataValidation » \| cellValueType « Visible » ? : « Tout » \| « Erreurs » \| « ErrorsLogical » \| « ErrorsNumbers » \| « ErrorsText » \| » ErrorsLogicalNumber » \| « ErrorsLogicalText » \| « ErrorsNumberText » \| « Logique » \| « LogicalNumbers » \| « LogicalText » \| « LogicalNumbersText » \| « Nombres » \| « NumbersText » \| « Texte »)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Renvoie un objet RangeAreas qui représente toutes les cellules correspondant au type et à la valeur spécifiés. Lève une erreur si aucune cellule spéciale n’est trouvée qui corresponde au critère.|
||[getSpecialCells (cellType : Excel.SpecialCellType, cellValueType ? : Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Renvoie un objet RangeAreas qui représente toutes les cellules correspondant au type et à la valeur spécifiés. Lève une erreur si aucune cellule spéciale n’est trouvée qui corresponde au critère.|
||[getSpecialCells (cellType : « ConditionalFormats » \| « DataValidations » \| « Vides » \| « Commentaires » \| « Constantes » \| « Formules » \| « SameConditionalFormat » \|« SameDataValidation » \| cellValueType « Visible » ? : « Tout » \| « Erreurs » \| « ErrorsLogical » \| « ErrorsNumbers » \| « ErrorsText » \| » ErrorsLogicalNumber » \| « ErrorsLogicalText » \| « ErrorsNumberText » \| « Logique » \| « LogicalNumbers » \| « LogicalText » \| « LogicalNumbersText » \| « Nombres » \| « NumbersText » \| « Texte »)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Renvoie un objet RangeAreas qui représente toutes les cellules correspondant au type et à la valeur spécifiés. Lève un objet null si aucune cellule spéciale n’est trouvée qui corresponde au critère.|
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
||[track()](/javascript/api/excel/excel.rangeareas#track--)|Effectuer le suivi de l’objet pour l’ajustement automatique en fonction environnant des modifications dans le document. Cet appel est abréviations context.trackedObjects.add(thisObject). Si vous utilisez cet objet au sein de « .sync » appels et en dehors de l’exécution séquentielle d’un lot de « .run » et rencontrez un message d’erreur « InvalidObjectPath » lors de la définition d’une propriété ou appeler une méthode sur l’objet, vous devez ajouter l’objet à l’objet de suivi collection de sites lors de l’objet a été créé.|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|Publication mémoire associée à cet objet si elle a été précédemment suivie. Cet appel est abréviations context.trackedObjects.add(thisObject). Vous rencontrez de nombreux objets suivies ralentit l’application hôte, donc n’oubliez pas de libérer les objets que l'on ajoute, une fois que vous avez terminé à les utiliser. Vous devez appeler « context.sync() » avant la publication de mémoire prend effet.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Renvoie ou définit un double qui s’éclaire ou assombrit une couleur de bordure de la plage, la valeur est comprise entre -1 (zones les plus sombres) et 1 (plus clair) avec 0 pour la couleur d’origine.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Renvoie ou définit un double qui s’éclaire ou assombrit une couleur de bordure de la plage, la valeur est comprise entre -1 (zones les plus sombres) et 1 (plus clair) avec 0 pour la couleur d’origine.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Renvoie le nombre de pages de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Renvoie la plage d’objet selon sa position dans la RangeCollection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.rangecollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
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
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[adresse](/javascript/api/excel/excel.rowproperties#address)||
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)||
||[hasSpill](/javascript/api/excel/excel.rowproperties#hasspill)||
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)||
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Spécifie si la correspondance doit être complète ou partielle. La valeur par défaut est False (partielle).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Spécifie si la correspondance respecte ou non la casse. Par défaut est false (ou minuscules).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Détermine le sens de la recherche. Par défaut est transférer. Voir Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)||
||[lien hypertexte](/javascript/api/excel/excel.settablecellproperties#hyperlink)||
||[style](/javascript/api/excel/excel.settablecellproperties#style)||
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)||
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)||
|[Paramètre](/javascript/api/excel/excel.setting)|[](/javascript/api/excel/excel.setting#replacestringdatewithdate)||
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Cette propriété renvoie ou définit la chaîne de texte descriptif (de remplacement) d'un objet Shape lors de l'enregistrement de l'objet dans une page Web. données String en lecture-écriture.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Cette propriété renvoie ou définit la chaîne de texte descriptif (de remplacement) d'un objet Shape lors de l'enregistrement de l'objet dans une page Web.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Supprime la forme.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Représente le type de forme géométrique de la forme spécifiée. Voir Excel.GeometricShapeType des détails. Renvoie null si la forme n’est pas géométrique, par exemple, se GeometricShapeType d’une ligne ou un graphique renverra null.|
||[height](/javascript/api/excel/excel.shape#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Déplace horizontalement la forme spécifiée selon le nombre de points indiqué.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|Fait pivoter la forme spécifiée, selon le nombre de degrés spécifié, autour de l'axe z.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Décale vers le haut la forme spécifiée selon le nombre de points spécifié.|
||[left](/javascript/api/excel/excel.shape#left)|La distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Représente si les proportions de ratio ont bloqué, en booléen, la forme.|
||[name](/javascript/api/excel/excel.shape#name)|Représente le nom de la forme.|
||[placement](/javascript/api/excel/excel.shape#placement)|Cette propriété représente le placment, valeur qui représente le mode d'attachement de l'objet aux cellules.|
||[fill](/javascript/api/excel/excel.shape#fill)|Renvoie la mise en forme de remplissage de l’objet de forme. En lecture seule.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Renvoie l’objet de la forme de la forme géométrique. L’erreur sera levée si l’objet de forme est autre type de forme (par exemple, Image, SmartArt, etc.) au lieu de GeometricShape.|
||[groupe](/javascript/api/excel/excel.shape#group)|Renvoie le groupe forme de l’objet forme. L’erreur sera levée si l’objet de forme est autre type de forme (par exemple, Image, SmartArt, etc.) au lieu de GroupShape.|
||[id](/javascript/api/excel/excel.shape#id)|Représente l’identificateur de forme. En lecture seule.|
||[image](/javascript/api/excel/excel.shape#image)|Renvoie l’image de l’objet forme. L’erreur sera levée si l’objet de forme est autre type de forme (par exemple, Image, SmartArt, etc.) au lieu de l’Image.|
||[level](/javascript/api/excel/excel.shape#level)|Représente le titre de la forme spécifiée. Niveau 0 signifie que la forme ne fait pas partie d’un groupe, niveau 1 signifie que la forme fait partie d’un groupe de niveau supérieur, etc..|
||[line](/javascript/api/excel/excel.shape#line)|Renvoie l’objet de la ligne de l’objet forme. L’erreur sera levée si l’objet de forme est autre type de forme (par exemple, Image, SmartArt, etc.) au lieu de l’Image.|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Renvoie la mise en forme de la ligne de l’objet de forme. En lecture seule.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Se produit lorsque la forme est activée.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Se produit lorsque la forme est activée.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Représente le groupe parent de la forme spécifiée.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Renvoie l’objet textFrame d’une forme. En lecture seule.|
||[type](/javascript/api/excel/excel.shape#type)|Renvoie le type de la forme spécifiée. En lecture seule. Voir Excel.GeometricShapeType des détails.|
||[zorderPosition](/javascript/api/excel/excel.shape#zorderposition)|Renvoie la position de la forme spécifiée dans l’ordre z, valeur z de commande de la forme tout en bas est égal à 0. En lecture seule.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Représente la rotation en degrés, de la forme.|
||[saveAsPicture (format : « Inconnu » \| « BMP » \| « JPEG » \| « GIF » \| « PNG » \| « SVG »)](/javascript/api/excel/excel.shape#saveaspicture-format-)|Enregistre la forme en tant qu’une image et renvoie l’image sous forme de chaîne encodé en base 64, en utilisant le PPP définit 96. Prise en charge uniquement enregistre estimer Excel.PictureFormat.BMP, Excel.PictureFormat.PNG, Excel.PictureFormat.JPEG et Excel.PictureFormat.GIF.|
||[saveAsPicture(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#saveaspicture-format-)|Enregistre la forme en tant qu’une image et renvoie l’image sous forme de chaîne encodé en base 64, en utilisant le PPP définit 96. Prise en charge uniquement enregistre estimer Excel.PictureFormat.BMP, Excel.PictureFormat.PNG, Excel.PictureFormat.JPEG et Excel.PictureFormat.GIF.|
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Met la hauteur de la forme à l’échelle en utilisant un facteur spécifié. Pour des images, vous pouvez indiquer si vous souhaitez mettre la forme à l’échelle par rapport à la taille d’origine ou la taille actuelle. Les formes autres que des images sont toujours mis à l’échelle par rapport à leur hauteur actuelle.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Met la hauteur de la forme à l’échelle en utilisant un facteur spécifié. Pour des images, vous pouvez indiquer si vous souhaitez mettre la forme à l’échelle par rapport à la taille d’origine ou la taille actuelle. Les formes autres que des images sont toujours mis à l’échelle par rapport à leur hauteur actuelle.|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Met la largeur de la forme à l’échelle en utilisant un facteur spécifié. Pour des images, vous pouvez indiquer si vous souhaitez mettre la forme à l’échelle par rapport à la taille d’origine ou la taille actuelle. Les formes autres que des images sont toujours mis à l’échelle par rapport à leur hauteur actuelle.|
||[scaleHeight (scaleFactor : numéro scaleType : Excel.ShapeScaleType, scaleFrom ? : Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Met la largeur de la forme à l’échelle en utilisant un facteur spécifié. Pour des images, vous pouvez indiquer si vous souhaitez mettre la forme à l’échelle par rapport à la taille d’origine ou la taille actuelle. Les formes autres que des images sont toujours mis à l’échelle par rapport à leur hauteur actuelle.|
||[setZOrder(value: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-value-)|Déplace la forme spécifiée devant ou derrière les autres formes dans la collection de (autrement dit, modifie position de la forme dans l’ordre z).|
||[setZOrder(value: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-value-)|Déplace la forme spécifiée devant ou derrière les autres formes dans la collection de (autrement dit, modifie position de la forme dans l’ordre z).|
||[top](/javascript/api/excel/excel.shape#top)|Distance, en points, du bord supérieur de l’objet au bord supérieur de la feuille de calcul.|
||[visible](/javascript/api/excel/excel.shape#visible)|Représente la visibilité de type boolean, de la forme spécifiée.|
||[width](/javascript/api/excel/excel.shape#width)|Représente la largeur, en points, de la forme.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Obtient l’id du graphique qui est activé.|
||[type](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la forme est activée.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus", left: number, top: number, width: number, height: number)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype--left--top--width--height-)|Ajoute une forme géométrique à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addGeometricShape (geometricShapeType : Excel.GeometricShapeType gauche : numéro haut : nombre, largeur : nombre, hauteur : nombre)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype--left--top--width--height-)|Ajoute une forme géométrique à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addGroup (valeurs : matrice < chaîne \| forme >)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Groupe un sous-ensemble de formes dans une feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Crée une image à partir d’une chaîne en base 64 et il est ajouté à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addLine (startLeft : startTop nombre, : endLeft nombre, : endTop nombre, : connectorType numéro ? : « Apostrophes » \| « en « angle \| « Courbe »)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Ajoute une ligne à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Ajoute une ligne à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addSVG(xmlImageString: string)](/javascript/api/excel/excel.shapecollection#addsvg-xmlimagestring-)|Crée une image à partir d’une chaîne en base 64 et il est ajouté à la feuille de calcul. Renvoie un objet Shape qui représente la nouvelle forme.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Ajoute une zone de texte à la feuille de calcul en indiquant à son contenu de texte. Elle renvoie un objet Shape qui représente la nouvelle zone de texte.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Renvoie le nombre de graphiques dans la feuille de calcul. En lecture seule.|
||[getItem(name: string)](/javascript/api/excel/excel.shapecollection#getitem-name-)|Extrait un graphique à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.shapecollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Obtient l’id de la forme qui est désactivée.|
||[type](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle la forme est désactivée.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Renvoie la mise en forme de remplissage de l’objet de forme.|
||[foreColor](/javascript/api/excel/excel.shapefill#forecolor)|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[type](/javascript/api/excel/excel.shapefill#type)|Renvoie le type de remplissage de la forme. En lecture seule. Voir Excel.GeometricShapeType des détails.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Définit la mise en forme de remplissage d’un objet de forme sur une couleur uniforme type remplissage modification remplissage uni.|
||[Transparency](/javascript/api/excel/excel.shapefill#transparency)|Renvoie ou définit le degré de transparence du remplissage spécifié sous la forme d’une valeur comprise entre 0,0 (opaque) et 1,0 (transparent). Types de forme API ne pas pris en charge ou un type spécial remplissage avec transparences incohérentes, renvoient null. Par exemple, type de remplissage dégradé pourrait avoir transparences incohérentes.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Représente le format de police Gras. Retourne la valeur null le TextRange inclut les deux fragments de texte en gras et en non.|
||[color](/javascript/api/excel/excel.shapefont#color)|Représentation sous forme de code couleur HTML de la couleur du texte. Par exemple, #FF0000 représente le rouge. Renvoie la valeur null si le TextRange inclut les fragments de texte avec les différentes couleurs.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Représente le format de police Italique. Renvoie null si le TextRange inclut les deux fragments de texte en italique et non italique.|
||[name](/javascript/api/excel/excel.shapefont#name)|Représente le nom de la police (par exemple « Calibri ») Si le texte est un langage de Script complexe ou Asie de l’est, représente le nom de la police correspondante ; dans le cas contraire représente nom de police de caractères latins.|
||[size](/javascript/api/excel/excel.shapefont#size)|Représente la taille de police en points (par exemple, 11). Renvoie la valeur null si le TextRange inclut les fragments de texte avec les différentes couleurs.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Type de soulignement appliqué à la police. Renvoie la valeur null si le TextRange inclut les fragments de texte avec les différentes couleurs. Pour plus d’informations, voir Excel.ShapeFontUnderlineStyle.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Représente l’identificateur de forme. En lecture seule.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Renvoie l’objet de la forme de la forme géométrique. En lecture seule.|
||[Formes](/javascript/api/excel/excel.shapegroup#shapes)|Renvoie la collection de forme dans le groupe. En lecture seule.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Dissocie toutes les formes groupées dans la forme spécifiée.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Code couleur HTML qui représente la couleur de remplissage, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Représente le style de trait de la forme. Renvoie la valeur null lors de la ligne n’est pas visible ou a mixte propriété du style tiret ligne (par exemple, type de groupe de forme). Pour plus d’informations, voir Excel.ShapeFontUnderlineStyle.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Représente le style de trait de l’objet de la forme. Renvoie la valeur null lors de la ligne n’est pas visible ou a mixte propriété du style tiret ligne (par exemple, type de groupe de forme). Pour plus d’informations, voir Excel.ShapeFontUnderlineStyle.|
||[Transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Renvoie ou définit le degré de transparence du remplissage spécifié sous la forme d’une valeur comprise entre 0,0 (opaque) et 1,0 (transparent). Renvoie la valeur null si la forme a une propriété de ligne de transparence mixte (par exemple, groupe type de forme).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Indique si la mise en forme de la ligne d’un élément de forme est visible. Renvoie la valeur null si la forme a une propriété de ligne de transparence mixte (par exemple, groupe type de forme).|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Représente l’épaisseur de bordure, en points. Renvoie la valeur null lors de la ligne n’est pas visible ou a mixte propriété du style tiret ligne (par exemple, type de groupe de forme).|
|[Segment](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Représente la légende de segment.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Supprime tous les filtres appliqués actuellement sur le tableau.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Supprime le segment.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Renvoie une matrice de noms d’éléments sélectionnés. En lecture seule.|
||[height](/javascript/api/excel/excel.slicer#height)|Représente la hauteur, exprimée en points, de l’axe de graphique.|
||[left](/javascript/api/excel/excel.slicer#left)|Représente la distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
||[name](/javascript/api/excel/excel.slicer#name)|Représente le nom de la forme.|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom utilisé dans la formule.|
||[id](/javascript/api/excel/excel.slicer#id)|Représente l’id unique du segment. En lecture seule.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|True si tous les filtres appliqués actuellement sur le tableau.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Représente la collection de SlicerItems qui font partie du segment. En lecture seule.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Obtenir la feuille de calcul contenant la plage. En lecture seule.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Éléments de segment à sélection multiple en fonction de leur nom. Sélection précédente est désactivée.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Représente l’ordre de tri des éléments dans le segment.|
||[style](/javascript/api/excel/excel.slicer#style)|Valeur de constante qui représente le style du tableau. Les valeurs possibles sont : SlicerStyleLight1 exécutée via SlicerStyleLight6, TableStyleOther1 exécutée via TableStyleOther2, SlicerStyleDark1 exécutée via SlicerStyleDark6. Vous pouvez également indiquer un style personnalisé présent dans le classeur.|
||[top](/javascript/api/excel/excel.slicer#top)|Représente la distance, en points, du bord supérieur de la section à la partie droite de la feuille de calcul.|
||[width](/javascript/api/excel/excel.slicer#width)|Représente la largeur, en points, de la forme.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[Ajouter (slicerSource : chaîne \| tableau croisé dynamique \| Table, sourceField : chaîne \| PivotField \| nombre \| TableColumn, slicerDestination ? : chaîne \| feuille de calcul)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Ajoute un nouveau segment au classeur.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Renvoie le nombre de séries de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID. Si la feuille de calcul n’existe pas, renvoie un objet null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.slicercollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|True si l’élément de slicer est sélectionné ; sinon False. Définir cette valeur n’efface pas état autres SlicerItems sélectionné.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|True si l’élément de segment comporte des données.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Représente la valeur unique représentant l’élément de segment.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Représente la valeur affichée dans l’interface utilisateur.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Renvoie le nombre de segment les éléments dans le segment.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Obtient un segment objet de l’élément à l’aide de son nom ou clé.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Obtient un segment de l’élément à l’aide de son nom ou clé. Si le paramètre n’existe pas, renvoie un objet null.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.sliceritemcollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Représente les sous-champs est le nom de la propriété cible d’une valeur enrichi effectuer le tri.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Obtient le nombre de tableaux de la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Obtient une forme en fonction de sa position dans la collection.|
|[Tableau](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[autoFilter](/javascript/api/excel/excel.table#autofilter)|Représente l’objet de filtre automatique de la table. En lecture seule.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Obtient l’ID du tableau. En lecture seule.|
||[type](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est ajouté.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Se produit lorsque la nouvelle table est ajoutée dans un classeur.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Se produit lorsque le tableau spécifié est supprimé dans un classeur.|
||[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Spécifie la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Obtient l’ID du tableau. En lecture seule.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Spécifie le nom du tableau qui est supprimé.|
||[type](/javascript/api/excel/excel.tabledeletedeventargs#type)|Spécifie le type du champ. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le graphique est ajouté.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Indique le nom de la table à laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Représente l’id de la feuille de calcul qui contient le tableau.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Obtient le nombre de tableaux de la collection.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Obtient le premier tableau de cette collection. Les tables dans la collection de sont triées de haut en bas et gauche vers la droite par ce tableau supérieure gauche afin que le premier tableau soit dans la collection de sites.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Obtient un tableau à l’aide de son nom ou de son ID.|
||[load(option?: OfficeExtension.LoadOption)](/javascript/api/excel/excel.tablescopedcollection#load-option-)|Files d’attente de la commande pour charger les propriétés de l’objet spécifié. Vous devez appeler « context.sync() » avant de lire les propriétés.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSize](/javascript/api/excel/excel.textframe#autosize)|Obtient ou définit le paramètres de texte de dimensionnement automatique. Un bloc de texte peut être défini à redimensionnement automatique de texte pour l’ajuster le bloc de texte ou la taille automatique le bloc de texte pour l’ajuster le texte ou sans redimensionnement automatique.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Représente la marge bas, en points du cadre du texte.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Supprime tout le texte dans la textframe.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Représente l’alignement horizontal pour le style.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Représente le type de débordement horizontal du cadre du texte.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Représente la marge gauche, en points du cadre du texte.|
||[Orientation](/javascript/api/excel/excel.textframe#orientation)|Représente l’orientation du texte de l’encadrement de texte.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Représente l’ordre de lecture du cadre texte RTL ou LTR.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Spécifie si la TextFrame contient du texte.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)||
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Représente la marge droite, en points du cadre du texte.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Représente la marge du haut, en points du cadre du texte.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Représente l’alignement vertical pour le style.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Représente le type de débordement vertical du cadre du texte.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getCharacters (début : nombre, longueur ? : nombre)](/javascript/api/excel/excel.textrange#getcharacters-start--length-)|Renvoie un objet TextRange pour les caractères dans la plage de donnée.|
||[police](/javascript/api/excel/excel.textrange#font)|Renvoie un objet ShapeFont qui représente les attributs de police pour la plage de texte. En lecture seule.|
||[text](/javascript/api/excel/excel.textrange#text)|Représente le contenu de texte brut de la plage de texte.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True si tous les graphiques dans le classeur suivent les points de données réelles auquel qu’il sont joints.|
||[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Obtient la feuille de calcul active du classeur. S’il n’existe aucun graphique actif, génère des exceptions lorsque appeler cette déclaration|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Obtient la feuille de calcul active du classeur. S’il n’existe aucun graphique actif, renverra l’objet null|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Obtient le segment actif actuel du classeur. S’il n’existe aucun segment actif, génère des exceptions lorsque appeler cette déclaration.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Obtient le segment actif actuel du classeur. S’il n’existe aucun graphique actif, il renverra l’objet null|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True si le classeur est modifié par plusieurs utilisateurs (co-création).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Obtient la ou les plage(s) sélectionnée(s) actuelle(s) dans le classeur. Contrairement aux getSelectedRange(), cette méthode renvoie un objet RangeAreas qui représente toutes les plages sélectionnées.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|True si le classeur spécifié n’a pas été modifié depuis son dernier enregistrement.Il a été récemment enregistré.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|True si le classeur se trouve dans l’enregistrement en mode automatique.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Renvoie un nombre sur la version de moteur de calcul Excel. En lecture seule.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Représente une collection de styles associés au classeur. En lecture seule.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Se produit lorsque le paramètre de l’enregistrement automatique est modifié dans le classeur.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|True si le classeur a jamais été enregistré localement ou en ligne.|
||[Slicers](/javascript/api/excel/excel.workbook#slicers)|Représente une collection de styles associés au classeur. En lecture seule.|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True si les calculs réalisés dans ce classeur utiliseront uniquement la précision des nombres tels qu’ils sont affichés. |
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Obtient ou définit EnableCalculation, propriété de la feuille de calcul.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Trouve toutes les occurrences de la chaîne donnée en fonction des critères spécifiées et renvoie un objet RangeAreas comprenant une ou plusieurs plages rectangulaires.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Trouve toutes les occurrences de la chaîne donnée en fonction des critères spécifiées et renvoie un objet RangeAreas comprenant une ou plusieurs plages rectangulaires.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Obtient l’objet RangeAreas représentant un ou plusieurs blocs de plages rectangulaires, spécifiés par nom ou l’adresse.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Représente l’objet AutoFilter de filtre automatique de la feuille de calcul. En lecture seule.|
||[comments](/javascript/api/excel/excel.worksheet#comments)|Renvoie une collection de tous les objets Lecteur sur l’ordinateur. En lecture seule.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Obtient la collection de saut de page horizontal pour la feuille de calcul. Cette collection contient uniquement les sauts de page manuels.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Se produit lorsque le filtre est modifié sur un tableau spécifique.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Obtient l’objet PageLayout de la feuille de calcul.|
||[Formes](/javascript/api/excel/excel.worksheet#shapes)|Renvoie une collection de tous les objets Forme sur la feuille de calcul. En lecture seule.|
||[Slicers](/javascript/api/excel/excel.worksheet#slicers)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Obtient la collection de saut de page vertical pour la feuille de calcul. Cette collection contient uniquement les sauts de page manuels.|
||[replaceAll (texte : chaîne remplacement : chaîne critères : Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Détecte et remplace la chaîne donnée basée sur les critères spécifiés dans la plage active.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64 (base64File : chaîne, sheetNamesToInsert ? : [] chaîne, positionType ? : « None » \| « Avant les caractères » \| « Après » \| « Depuis » \| « Fin », relativeTo ? : feuille de calcul \| chaîne)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Se produit lorsqu’une feuille de calcul dans le classeur est modifié.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Se produit lorsqu’une feuille de calcul dans le classeur a un format modifié.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Cet événement survient lorsque la sélection change dans une feuille de calcul.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Représente le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Indique le nom du tableau auquel le filtre est appliqué.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique. Il peut renvoyer un objet null.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Spécifie si la correspondance doit être complète ou partielle. La valeur par défaut est False (partielle).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Spécifie si la correspondance respecte ou non la casse. Par défaut est false (ou minuscules).|

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
