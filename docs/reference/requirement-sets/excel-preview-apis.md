---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir
ms.date: 03/19/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fda0721bd5d7cbec6349c4800a97132d61a26ab9
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891200"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Paramètres de culture](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings-preview) | Obtient les paramètres du système culturel pour le classeur, tels que la mise en forme des nombres. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [application](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Insérer un classeur](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insérer un classeur dans un autre.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Filtres de tableau croisé dynamique | Applique des filtres pilotés par valeur aux champs d’un tableau croisé dynamique. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| Classeur [enregistrer](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) et [fermer](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) | Enregistrez et fermez ses classeurs.  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en version préliminaire. Pour afficher la liste complète de toutes les API JavaScript pour Excel (y compris les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview).

| Class | Champs | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fournit des informations basées sur les paramètres de culture système actuels. Cela inclut les noms de culture, la mise en forme de numéros et d’autres paramètres dépendants de la culture.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres locaux d’Excel.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres locaux d’Excel.|
||[UseSystemSeparators,](/javascript/api/excel/excel.application#usesystemseparators)|Indique si les séparateurs système de Microsoft Excel sont activés.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Représente l’angle auquel le texte est orienté pour le titre de l’axe du graphique. La valeur doit être un entier compris entre-90 et 90 ou l’entier 180 pour le texte orienté verticalement.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension : Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtient les valeurs d’une dimension unique de la série de graphiques. Il peut s’agir de valeurs de catégorie ou de valeurs de données, en fonction de la dimension spécifiée et de la façon dont les données sont mappées pour la série de graphiques.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtient le type de contenu du commentaire.|
||[évaluation](/javascript/api/excel/excel.comment#resolved)|Obtient ou définit l’état du thème de commentaire. La valeur « true » signifie que le thread est résolu.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Obtient le type de contenu de la réponse.|
||[évaluation](/javascript/api/excel/excel.commentreply#resolved)|Obtient ou définit l’état de la réponse. La valeur « true » signifie que la réponse est à l’État résolu.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Définit le format d’affichage de la date et de l’heure approprié pour la culture. Cette fonction est basée sur les paramètres de culture actuelle du système.|
||[name](/javascript/api/excel/excel.cultureinfo#name)|Obtient le nom de la culture au format languagecode2-Country/regioncode2 (par exemple « zh-CN » ou « en-US »). Cette fonction est basée sur les paramètres système actuels.|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Définit le format d’affichage des nombres approprié pour la culture. Cette fonction est basée sur les paramètres de culture actuelle du système.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[DateSeparator,](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtient la chaîne utilisée comme séparateur de date. Cette fonction est basée sur les paramètres système actuels.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtient la chaîne de format pour une valeur de date longue. Cette fonction est basée sur les paramètres système actuels.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtient la chaîne de format pour une valeur d’heure longue. Cette fonction est basée sur les paramètres système actuels.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtient la chaîne de format pour une valeur de date courte. Cette fonction est basée sur les paramètres système actuels.|
||[TimeSeparator,](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtient la chaîne utilisée comme séparateur d’heure. Cette fonction est basée sur les paramètres système actuels.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtient la chaîne utilisée comme séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres système actuels.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtient la chaîne utilisée pour séparer les groupes de chiffres à gauche du séparateur décimal pour les valeurs numériques. Cette fonction est basée sur les paramètres système actuels.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[identifie](/javascript/api/excel/excel.pivotdatefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées. Le type de comparaison est défini par la condition.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Indique la condition pour le filtre, qui définit les critères de filtrage nécessaires.|
||[consenti](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Si la valeur est true, Filter *exclut* les éléments qui répondent aux critères. La valeur par défaut est false (filtre pour inclure les éléments qui satisfont les critères).|
||[Inférieures](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Limite inférieure de la plage de la `Between` condition de filtre.|
||[Haute](/javascript/api/excel/excel.pivotdatefilter#upperbound)|La limite supérieure de la plage pour la `Between` condition de filtre.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Pour `Equals`, `Before`, `After`, et `Between` conditions de filtre, indique si les comparaisons doivent être effectuées comme des journées entières.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filtre : PivotValueFilter \| PivotLabelFilter \| PivotManualFilter \| PivotDateFilter \| PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Définit un ou plusieurs éléments de la valeur de la propriété PivotFilters actuelle du champ et les applique au champ.|
||[ClearAllFilters, ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Efface tous les critères de tous les filtres du champ. Cela supprime tout filtrage actif sur le champ.|
||[clearFilter (filterType : Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Efface tous les critères existants du filtre du champ du type donné (s’il est déjà appliqué).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtient tous les filtres actuellement appliqués sur le champ.|
||[isFiltered (filterType ?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Vérifie s’il existe des filtres appliqués sur le champ.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|Filtre date d’application du champ PivotField. NULL si aucune n’est appliquée.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|Filtre d’étiquette du champ de tableau croisé dynamique actuellement appliqué. NULL si aucune n’est appliquée.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|Filtre manuel actuellement appliqué au champ de tableau croisé dynamique. NULL si aucune n’est appliquée.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|Filtre de valeur actuellement appliqué au champ PivotField. NULL si aucune n’est appliquée.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[identifie](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées. Le type de comparaison est défini par la condition.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Indique la condition pour le filtre, qui définit les critères de filtrage nécessaires.|
||[consenti](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Si la valeur est true, Filter *exclut* les éléments qui répondent aux critères. La valeur par défaut est false (filtre pour inclure les éléments qui satisfont les critères).|
||[Inférieures](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|La limite inférieure de la plage pour la condition entre le filtre.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Sous-chaîne utilisée pour `BeginsWith`les `EndsWith`conditions de `Contains` filtre,, et.|
||[Haute](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|La limite supérieure de la plage pour la condition entre le filtre.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives. La cellule renvoyée est l’intersection de la ligne donnée et une colonne qui contient les données à partir de la hiérarchie donnée. Cette méthode est l’inverse de l’appel getPivotItems et getDataHierarchy sur une cellule particulière.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Style appliqué au tableau croisé dynamique.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Définit le style appliqué au tableau croisé dynamique.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Liste des éléments sélectionnés à filtrer manuellement. Ces éléments doivent être existants et valides dans le champ choisi.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Indique si le tableau croisé dynamique autorise l’application de plusieurs PivotFilters sur un champ PivotField donné dans le tableau.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtient le nombre de tableaux croisés dynamiques dans la collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtient le premier tableau croisé dynamique de la collection. Les tableaux croisés dynamiques de la collection sont triés de haut en bas et de gauche à droite, de sorte que le tableau supérieur gauche est le premier tableau croisé dynamique de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtient un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[identifie](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées. Le type de comparaison est défini par la condition.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Indique la condition pour le filtre, qui définit les critères de filtrage nécessaires.|
||[consenti](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Si la valeur est true, Filter *exclut* les éléments qui répondent aux critères. La valeur par défaut est false (filtre pour inclure les éléments qui satisfont les critères).|
||[Inférieures](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Limite inférieure de la plage de la `Between` condition de filtre.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Indique si le filtre est destiné aux N éléments supérieurs/inférieurs, aux N pourcentages supérieur/inférieur ou supérieur/inférieur N.|
||[seuil](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Le nombre de seuils « N » d’éléments, pourcentage ou somme à filtrer pour une condition de filtre de haut en bas.|
||[Haute](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|La limite supérieure de la plage pour la `Between` condition de filtre.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nom de la « valeur » sélectionnée dans le champ à utiliser pour filtrer.|
|[Range](/javascript/api/excel/excel.range)|[getPivotTables (fullyContained ?: booléen)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtient une collection d’étendues de tableaux croisés dynamiques qui se chevauchent avec la plage.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. En lecture seule.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. Échoue si appliqué à une plage comportant plusieurs cellules. En lecture seule.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. En lecture seule.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Représente la catégorie de format numérique de chaque cellule. En lecture seule.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Style appliqué au Slicer.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Définit le style appliqué au segment.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Style appliqué au tableau.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Définit le style appliqué au segment.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtient l’ID de la table dans laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Fermer (closeBehavior ? : Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fermer le classeur actif.|
||[Enregistrer (saveBehavior ? : Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Enregistrer le classeur actif.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtient une collection de propriétés personnalisées au niveau de la feuille de calcul.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[adresse](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Adresse de la plage qui a terminé le calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Survient lorsque l’état masqué d’une ou plusieurs lignes a été modifié sur une feuille de calcul spécifique.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtient la clé de la propriété personnalisée. En lecture seule.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtient la valeur de la propriété personnalisée. En lecture seule.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtient le nombre de propriétés personnalisées sur cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Lève une exception si la propriété personnalisée n’existe pas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[adresse](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtient l’adresse de plage qui représente la zone modifiée dans une feuille de calcul spécifique.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtient le type de modification qui représente la manière dont l’événement a été déclenché. Pour `Excel.RowHiddenChangeType` plus d’informations, voir.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtient la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle les données sont modifiées.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
