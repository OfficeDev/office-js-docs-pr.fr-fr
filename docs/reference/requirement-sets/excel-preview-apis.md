---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir
ms.date: 06/29/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d1701ad393b96e33f0007bfcb5609c93c13608a2
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430764"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Paramètres de culture](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) de la date et de l’heure | Donne accès à des paramètres culturels supplémentaires par rapport à la mise en forme de la date et de l’heure. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [application](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Insérer un classeur](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Insérer un classeur dans un autre.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Filtres de tableau croisé dynamique | Applique des filtres pilotés par valeur aux champs d’un tableau croisé dynamique. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
|Plage de débordement | Permet aux compléments de trouver des plages associées aux résultats de [tableau dynamique](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en version préliminaire. Pour afficher la liste complète de toutes les API JavaScript pour Excel (y compris les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension : Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtient les valeurs d’une dimension unique de la série de graphiques. Il peut s’agir de valeurs de catégorie ou de valeurs de données, en fonction de la dimension spécifiée et de la façon dont les données sont mappées pour la série de graphiques.|
|[Commentaire](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtient le type de contenu du commentaire.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtient le `CommentDetail` tableau qui contient le numéro de commentaire et les ID des réponses associées.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Spécifie la source de l’événement. `Excel.EventSource`Pour plus d’informations, voir.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtient le type de l’événement. `Excel.EventType`Pour plus d’informations, voir.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtient le type de modification qui représente le mode de déclenchement de l’événement modifié.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtient le `CommentDetail` tableau qui contient le code de commentaire et les ID des réponses associées.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Spécifie la source de l’événement. `Excel.EventSource`Pour plus d’informations, voir.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtient le type de l’événement. `Excel.EventType`Pour plus d’informations, voir.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Se produit lors de l’ajout de commentaires.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Survient lorsque des commentaires ou des réponses dans une collection de commentaires sont modifiés, y compris quand les réponses sont supprimées.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Cet événement se produit lorsque des commentaires sont supprimés dans la collection comment.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtient le `CommentDetail` tableau qui contient le numéro de commentaire et les ID des réponses associées.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Spécifie la source de l’événement. `Excel.EventSource`Pour plus d’informations, voir.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtient le type de l’événement. `Excel.EventType`Pour plus d’informations, voir.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Représente l’ID du commentaire.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Représente les ID des réponses associées appartenant au commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Type de contenu de la réponse.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Définit le format d’affichage de la date et de l’heure approprié pour la culture. Cette fonction est basée sur les paramètres de culture actuelle du système.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[DateSeparator,](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtient la chaîne utilisée comme séparateur de date. Cette fonction est basée sur les paramètres système actuels.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtient la chaîne de format pour une valeur de date longue. Cette fonction est basée sur les paramètres système actuels.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtient la chaîne de format pour une valeur d’heure longue. Cette fonction est basée sur les paramètres système actuels.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtient la chaîne de format pour une valeur de date courte. Cette fonction est basée sur les paramètres système actuels.|
||[TimeSeparator,](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtient la chaîne utilisée comme séparateur d’heure. Cette fonction est basée sur les paramètres système actuels.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Active l’affichage tableau. Cela équivaut à l’utilisation de l’option « Basculer vers » dans l’interface utilisateur Excel.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Supprime l’affichage tableau de la feuille de calcul.|
||[doublon (Name ?: String)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Crée une copie de l’affichage de cette feuille.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtient ou définit le nom de l’affichage tableau.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Crée une nouvelle vue de feuille portant le nom donné.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Crée et active un nouvel affichage de tableau temporaire.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Quitte l’affichage de la feuille active.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtient l’affichage de la feuille actuellement actif de la feuille de calcul.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtient le nombre d’affichages de feuille dans cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtient un affichage tableau à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtient un affichage feuille par son index dans la collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[identifie](/javascript/api/excel/excel.pivotdatefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées. Le type de comparaison est défini par la condition.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Spécifie la condition pour le filtre, qui définit les critères de filtrage nécessaires.|
||[consenti](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Si la valeur est true, Filter *exclut* les éléments qui répondent aux critères. La valeur par défaut est false (filtre pour inclure les éléments qui satisfont les critères).|
||[Inférieures](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Limite inférieure de la plage de la condition de `Between` filtre.|
||[Haute](/javascript/api/excel/excel.pivotdatefilter#upperbound)|La limite supérieure de la plage pour la `Between` condition de filtre.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Pour `Equals` , `Before` , `After` , et `Between` conditions de filtre, indique si les comparaisons doivent être effectuées comme des journées entières.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (filtre : Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Définit une ou plusieurs des informations de la valeur de champ en vigueur du champ et les applique au champ.|
||[ClearAllFilters, ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Efface tous les critères de tous les filtres du champ. Cela supprime tout filtrage actif sur le champ.|
||[clearFilter (filterType : Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Efface tous les critères existants du filtre du champ du type donné (s’il est déjà appliqué).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtient tous les filtres actuellement appliqués sur le champ.|
||[isFiltered (filterType ?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Vérifie s’il existe des filtres appliqués sur le champ.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|Filtre date d’application du champ PivotField. NULL si aucune n’est appliquée.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|Filtre d’étiquette du champ de tableau croisé dynamique actuellement appliqué. NULL si aucune n’est appliquée.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|Filtre manuel actuellement appliqué au champ de tableau croisé dynamique. NULL si aucune n’est appliquée.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|Filtre de valeur actuellement appliqué au champ PivotField. NULL si aucune n’est appliquée.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[identifie](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées. Le type de comparaison est défini par la condition.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Spécifie la condition pour le filtre, qui définit les critères de filtrage nécessaires.|
||[consenti](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Si la valeur est true, Filter *exclut* les éléments qui répondent aux critères. La valeur par défaut est false (filtre pour inclure les éléments qui satisfont les critères).|
||[Inférieures](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|La limite inférieure de la plage pour la condition entre le filtre.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Sous-chaîne utilisée pour `BeginsWith` les `EndsWith` conditions de filtre,, et `Contains` .|
||[Haute](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|La limite supérieure de la plage pour la condition entre le filtre.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives. La cellule renvoyée est l’intersection de la ligne donnée et une colonne qui contient les données à partir de la hiérarchie donnée. Cette méthode est l’inverse de l’appel getPivotItems et getDataHierarchy sur une cellule particulière.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Style appliqué au tableau croisé dynamique.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Définit le style appliqué au tableau croisé dynamique.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Liste des éléments sélectionnés à filtrer manuellement. Ces éléments doivent être existants et valides dans le champ choisi.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Indique si le tableau croisé dynamique autorise l’application de plusieurs PivotFilters sur un champ PivotField donné dans le tableau.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[identifie](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées. Le type de comparaison est défini par la condition.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Spécifie la condition pour le filtre, qui définit les critères de filtrage nécessaires.|
||[consenti](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Si la valeur est true, Filter *exclut* les éléments qui répondent aux critères. La valeur par défaut est false (filtre pour inclure les éléments qui satisfont les critères).|
||[Inférieures](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Limite inférieure de la plage de la condition de `Between` filtre.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Indique si le filtre est destiné à l’élément N haut/bas, le niveau haut/bas de N pour cent, ou la somme N-Top/Bottom.|
||[seuil](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Le nombre de seuils « N » d’éléments, pourcentage ou somme à filtrer pour une condition de filtre de haut en bas.|
||[Haute](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|La limite supérieure de la plage pour la `Between` condition de filtre.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nom de la « valeur » sélectionnée dans le champ à utiliser pour filtrer.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Renvoie un `WorkbookRangeAreas` Object qui représente la plage contenant tous les antécédents directs d’une cellule dans une même feuille de calcul ou dans plusieurs feuilles de calcul.|
||[getMergedAreas()](/javascript/api/excel/excel.range#getmergedareas--)|Renvoie un objet RangeAreas qui représente les zones fusionnées dans cette plage. Notez que si le nombre de zones fusionnées dans cette plage est supérieur à 512, l’API ne renverra pas le résultat.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Renvoie un `WorkbookRangeAreas` Object qui représente la plage contenant tous les antécédents d’une cellule dans une même feuille de calcul ou dans plusieurs feuilles de calcul.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans. Échoue si appliqué à une plage comportant plusieurs cellules.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage. Échoue si appliqué à une plage comportant plusieurs cellules.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Représente la catégorie de format numérique de chaque cellule.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si toutes les cellules sont enregistrées sous la forme d’une formule matricielle.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Obtient le nombre d’objets RangeAreas de cette collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Renvoie l’objet RangeAreas en fonction de la position dans la collection.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul. Renvoie un objet Forme qui représente la nouvelle image.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Style appliqué au Slicer.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Définit le style appliqué au segment.|
|[Tableau](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Style appliqué au tableau.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Définit le style appliqué au segment.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtient l’ID de la table dans laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[Classeur](/javascript/api/excel/excel.workbook)|[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Indique si le volet de liste de champs du tableau croisé dynamique est affiché au niveau du classeur.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (Key : chaîne)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Renvoie l' `RangeAreas` objet basé sur l’ID ou le nom de la feuille de calcul dans la collection.|
||[getRangeAreasOrNullObjectBySheet (Key : chaîne)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Renvoie l' `RangeAreas` objet basé sur le nom ou l’ID de la feuille de calcul dans la collection. Si la feuille de calcul n’existe pas, renvoie un objet null.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Renvoie un tableau d’adresses en style a1. La valeur Address contient le nom de la feuille de calcul pour chaque bloc rectangulaire de cellules (par exemple, «Sheet1 ! A1 : B4, Sheet1 ! D1 : D4 "). En lecture seule.|
||[Zones](/javascript/api/excel/excel.workbookrangeareas#areas)|Renvoie l’objet RangeAreasCollection, chaque RangeAreas de la collection représentant une ou plusieurs plages de rectangles dans une feuille de calcul.|
||[fourneau](/javascript/api/excel/excel.workbookrangeareas#ranges)|Renvoie une collection de plages qui comprend cet objet.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtient une collection de propriétés personnalisées au niveau de la feuille de calcul.|
||[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Renvoie une collection de vues de feuille présentes dans la feuille de calcul.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtient la clé de la propriété personnalisée. Les clés de propriété personnalisée ne sont pas sensibles à la casse. La clé est limitée à 255 caractères (les plus grandes valeurs entraînent la levée d’une erreur « InvalidArgument »).|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key : chaîne, value : chaîne)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Ajoute une nouvelle propriété personnalisée qui est mappée à la clé fournie. Cette opération remplace les propriétés personnalisées existantes par cette clé.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtient le nombre de propriétés personnalisées sur cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Lève une exception si la propriété personnalisée n’existe pas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](./excel-api-requirement-sets.md)
