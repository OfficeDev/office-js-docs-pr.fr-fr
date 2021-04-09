---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1.12
description: Détails sur l’ensemble de conditions requises ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d66f5797d41c8c07f66fcc8069cd4687cd8d8118
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652215"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Nouveautés de l’API JavaScript 1.12 pour Excel

ExcelApi 1.12 a augmenté la prise en charge des formules dans les plages en ajoutant des API pour le suivi des tableaux dynamiques et la recherche des antécédents directs d’une formule. Il a également ajouté le contrôle API des filtres de tableau croisé dynamique. Des améliorations ont également été apportées aux zones de fonctionnalités de commentaires, de paramètres de culture et de propriétés personnalisées.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Événements de commentaire](../../excel/excel-add-ins-comments.md#comment-events) | Ajoute des événements pour ajouter, modifier et supprimer à la collection de commentaires.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Paramètres de [culture de date et d’heure](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Donne accès à des paramètres culturels supplémentaires autour de la mise en forme de date et d’heure. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Application NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Antécédents directs](../../excel/excel-add-ins-ranges-precedents.md) | Renvoie les plages utilisées pour évaluer la formule d’une cellule.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Filtres pivot | Applique des filtres pilotés par des valeurs aux champs d’un tableau croisé dynamique. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [Étendue de plage](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | Permet aux modules de recherche de plages associées à des [résultats de tableau](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) dynamique. | [Range](/javascript/api/excel/excel.range) |
| [Propriétés personnalisées au niveau de la feuille de calcul](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Permet d’étendue des propriétés personnalisées au niveau de la feuille de calcul, en plus de l’étendue au niveau du workbook. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API de l’ensemble de conditions requises de l’API JavaScript pour Excel 1.12. Pour afficher la documentation de référence de l’API pour toutes les API prise en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1.12 ou une version antérieure, voir API Excel dans l’ensemble de conditions requises [1.12](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)ou version antérieure.

| Classe | Champs | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Spécifie l’angle vers lequel le texte est orienté pour le titre de l’axe du graphique.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtient les valeurs d’une dimension unique de la série de graphiques.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtient le type de contenu du commentaire.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtenez le tableau CommentDetail qui contient l’ID de commentaire et les ID de ses réponses connexes.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Spécifie la source de l’événement.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtient le type de modification qui représente la façon dont l’événement modifié est déclenché.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtenez le tableau CommentDetail qui contient l’ID de commentaire et les ID de ses réponses connexes.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Spécifie la source de l’événement.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Se produit lorsque les commentaires sont ajoutés.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Se produit lorsque des commentaires ou des réponses dans une collection de commentaires sont modifiés, y compris lorsque les réponses sont supprimées.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Se produit lorsque des commentaires sont supprimés dans la collection de commentaires.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtenez le tableau CommentDetail qui contient l’ID de commentaire et les ID de ses réponses connexes.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Spécifie la source de l’événement.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Représente l’ID du commentaire.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Représente les ID des réponses associées qui appartiennent au commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Type de contenu de la réponse.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Définit le format adapté à la culture de l’affichage de la date et de l’heure.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Obtient la chaîne utilisée comme séparateur de date.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Obtient la chaîne de format pour une valeur de date longue.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Obtient la chaîne de format pour une valeur de temps longue.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Obtient la chaîne de format pour une valeur de date courte.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Obtient la chaîne utilisée comme séparateur d’heure.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[comparator](/javascript/api/excel/excel.pivotdatefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées.|
||[condition](/javascript/api/excel/excel.pivotdatefilter#condition)|Spécifie la condition du filtre, qui définit les critères de filtrage nécessaires.|
||[exclusive](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Si la valeur est True, le *filtre exclut les* éléments qui répondent aux critères.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Limite inférieure de la plage pour la `Between` condition de filtre.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Limite supérieure de la plage pour la `Between` condition de filtre.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Pour `Equals` , et les conditions de `Before` `After` `Between` filtre, indique si les comparaisons doivent être réalisées en tant que jours entiers.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Définit un ou plusieurs des filtres de tableau croisé dynamique actuels du champ et les applique au champ.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Permet d’effacer tous les critères de tous les filtres du champ.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Permet d’effacer tous les critères existants du filtre du champ du type donné (s’il en existe un actuellement appliqué).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Obtient tous les filtres actuellement appliqués sur le champ.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Vérifie s’il existe des filtres appliqués sur le champ.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|Filtre de date actuellement appliqué au champ de tableau croisé dynamique.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|Filtre d’étiquette actuellement appliqué au champ de tableau croisé dynamique.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|Filtre manuel actuellement appliqué au champ de tableau croisé dynamique.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|Filtre de valeurs actuellement appliqué au champ de tableau croisé dynamique.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[comparator](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées.|
||[condition](/javascript/api/excel/excel.pivotlabelfilter#condition)|Spécifie la condition du filtre, qui définit les critères de filtrage nécessaires.|
||[exclusive](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Si la valeur est True, le *filtre exclut les* éléments qui répondent aux critères.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Limite inférieure de la plage pour la condition de filtre Between.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Sous-stration utilisée pour `BeginsWith` , et les conditions de `EndsWith` `Contains` filtre.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Limite supérieure de la plage pour la condition de filtre Between.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Liste des éléments sélectionnés à filtrer manuellement.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Spécifie si le tableau croisé dynamique autorise l’application de plusieurs filtres de tableau croisé dynamique sur un champ de tableau croisé dynamique donné dans le tableau.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtient le nombre de tableaux croisés dynamiques dans la collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtient le premier tableau croisé dynamique de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtient un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Obtient un tableau croisé dynamique par nom.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[comparator](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Spécifie la condition du filtre, qui définit les critères de filtrage nécessaires.|
||[exclusive](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Si la valeur est True, le *filtre exclut les* éléments qui répondent aux critères.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Limite inférieure de la plage pour la `Between` condition de filtre.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Spécifie si le filtre est pour les éléments N supérieur/inférieur, le pourcentage N supérieur/inférieur ou la somme N supérieure/inférieure.|
||[seuil](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Nombre seuil « N » d’éléments, de pourcentage ou de somme à filtrer pour une condition de filtre Haut/Bas.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Limite supérieure de la plage pour la `Between` condition de filtre.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nom de la « valeur » choisie dans le champ par lequel filtrer.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Renvoie un objet WorkbookRangeAreas qui représente la plage contenant tous les antécédents directs d’une cellule dans la même feuille de calcul ou dans plusieurs feuilles de calcul.|
||[getPivotTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtient une collection étendue de tableaux croisés dynamiques qui chevauchent la plage.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Obtient l’objet de la plage contenant la cellule d’ancrage d’une cellule prise renversée dans.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Obtient l’objet de la plage contenant la plage renversé lorsque appelée sur une cellule d’ancrage.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Représente si toutes les cellules ont une bordure renversée.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Représente la catégorie du format de nombre de chaque cellule.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Représente si TOUTES les cellules doivent être enregistrées sous la la mesure d’une formule ma matrice.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Obtient le nombre d’objets RangeAreas dans cette collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Renvoie l’objet RangeAreas en fonction de la position dans la collection.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Renvoie `RangeAreas` l’objet en fonction de l’ID de feuille de calcul ou du nom de la collection.|
||[getRangeAreasOrNullObjectBySheet(key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Renvoie `RangeAreas` l’objet en fonction du nom ou de l’ID de la feuille de calcul dans la collection.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Renvoie un tableau d’adresses de type A1.|
||[Zones](/javascript/api/excel/excel.workbookrangeareas#areas)|Renvoie `RangeAreasCollection` l’objet.|
||[plages](/javascript/api/excel/excel.workbookrangeareas#ranges)|Renvoie les plages qui composent cet objet dans un `RangeCollection` objet.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtient une collection de propriétés personnalisées au niveau de la feuille de calcul.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtient la clé de la propriété personnalisée.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Ajoute une nouvelle propriété personnalisée qui s’ajoute à la clé fournie.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtient le nombre de propriétés personnalisées dans cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
