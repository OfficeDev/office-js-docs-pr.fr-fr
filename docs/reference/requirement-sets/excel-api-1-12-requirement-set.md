---
title: Ensemble de conditions requises de l’API JavaScript pour Excel 1,12
description: Détails sur l’ensemble de conditions requises ExcelApi 1,12.
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a88c511e90fe48e1a9997d19cb4a2851cb718f6b
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819840"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Nouveautés de l’API JavaScript pour Excel 1,12

Le ExcelApi 1,12 a augmenté la prise en charge des formules dans les plages en ajoutant des API pour le suivi des tableaux dynamiques et en recherchant les antécédents directs d’une formule. Il a également ajouté le contrôle de l’API des filtres de tableau croisé dynamique. Des améliorations ont également été apportées aux zones de fonctionnalité commentaires, paramètres de culture et propriétés personnalisées.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Événements de commentaire](../../excel/excel-add-ins-events.md) | Ajoute des événements pour l’ajout, la modification et la suppression à la collection de commentaires.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Paramètres de culture](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) de la date et de l’heure | Donne accès à des paramètres culturels supplémentaires par rapport à la mise en forme de la date et de l’heure. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [application](/javascript/api/excel/excel.application) NumberFormatInfo |
| Antécédents directs | Renvoie les plages utilisées pour évaluer la formule d’une cellule.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Filtres de tableau croisé dynamique | Applique des filtres pilotés par valeur aux champs d’un tableau croisé dynamique. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [Plage de débordement](../../excel/excel-add-ins-ranges-advanced.md#handle-dynamic-arrays-and-spilling) | Permet aux compléments de trouver des plages associées aux résultats de [tableau dynamique](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Propriétés personnalisées au niveau de la feuille de calcul](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Permet aux propriétés personnalisées d’être étendues au niveau de la feuille de calcul, en plus d’être étendue au niveau du classeur. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API dans l’ensemble de conditions requises de l’API JavaScript pour Excel 1,12. Pour afficher la documentation de référence de l’API pour toutes les API prises en charge par l’ensemble de conditions requises de l’API JavaScript pour Excel 1,12 ou antérieure, voir [API Excel dans l’ensemble de conditions requises 1,12 ou version antérieure](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Cette énumération spécifie l’angle auquel le texte est orienté pour le titre de l’axe du graphique. La valeur doit être un entier compris entre-90 et 90 ou l’entier 180 pour le texte orienté verticalement.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (dimension : Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Obtient les valeurs d’une dimension unique de la série de graphiques. Il peut s’agir de valeurs de catégorie ou de valeurs de données, en fonction de la dimension spécifiée et de la façon dont les données sont mappées pour la série de graphiques.|
|[Commentaire](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Obtient le type de contenu du commentaire.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Obtenir le tableau CommentDetail qui contient le code de commentaire et les ID des réponses associées.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Spécifie la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Obtient le type de modification qui représente le mode de déclenchement de l’événement modifié.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Obtenir le tableau CommentDetail qui contient le code de commentaire et les ID des réponses associées.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Spécifie la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’événement s’est produit.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Se produit lors de l’ajout de commentaires.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Survient lorsque des commentaires ou des réponses dans une collection de commentaires sont modifiés, y compris quand les réponses sont supprimées.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Cet événement se produit lorsque des commentaires sont supprimés dans la collection comment.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Obtenir le tableau CommentDetail qui contient le code de commentaire et les ID des réponses associées.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Spécifie la source de l’événement. Pour plus d’informations, voir Excel.EventSource.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Obtient le type de l’événement. Pour plus d’informations, voir Excel.EventType.|
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
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Style appliqué au tableau croisé dynamique.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Liste des éléments sélectionnés à filtrer manuellement. Ces éléments doivent être existants et valides dans le champ choisi.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Indique si le tableau croisé dynamique autorise l’application de plusieurs PivotFilters sur un champ PivotField donné dans le tableau.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Obtient le nombre de tableaux croisés dynamiques dans la collection.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Obtient le premier tableau croisé dynamique de la collection. Les tableaux croisés dynamiques de la collection sont triés de haut en bas et de gauche à droite, de sorte que le tableau supérieur gauche est le premier tableau croisé dynamique de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Obtient un tableau croisé dynamique par nom.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Extrait un tableau croisé dynamique par nom. Si le tableau croisé dynamique n’existe pas, renvoie un objet null.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[identifie](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Le comparateur est la valeur statique à laquelle les autres valeurs sont comparées. Le type de comparaison est défini par la condition.|
||[condition](/javascript/api/excel/excel.pivotvaluefilter#condition)|Spécifie la condition pour le filtre, qui définit les critères de filtrage nécessaires.|
||[consenti](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Si la valeur est true, Filter *exclut* les éléments qui répondent aux critères. La valeur par défaut est false (filtre pour inclure les éléments qui satisfont les critères).|
||[Inférieures](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Limite inférieure de la plage de la condition de `Between` filtre.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Indique si le filtre est destiné à l’élément N haut/bas, le niveau haut/bas de N pour cent, ou la somme N-Top/Bottom.|
||[seuil](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Le nombre de seuils « N » d’éléments, pourcentage ou somme à filtrer pour une condition de filtre de haut en bas.|
||[Haute](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|La limite supérieure de la plage pour la `Between` condition de filtre.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Nom de la « valeur » sélectionnée dans le champ à utiliser pour filtrer.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Renvoie un objet WorkbookRangeAreas qui représente la plage contenant tous les antécédents directs d’une cellule dans une même feuille de calcul ou dans plusieurs feuilles de calcul.|
||[getPivotTables (fullyContained ?: booléen)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Obtient une collection d’étendues de tableaux croisés dynamiques qui se chevauchent avec la plage.|
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
|[Segment](/javascript/api/excel/excel.slicer)|[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Style appliqué au Slicer.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (Key : chaîne)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Renvoie l' `RangeAreas` objet basé sur l’ID ou le nom de la feuille de calcul dans la collection.|
||[getRangeAreasOrNullObjectBySheet (Key : chaîne)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Renvoie l' `RangeAreas` objet basé sur le nom ou l’ID de la feuille de calcul dans la collection. Si la feuille de calcul n’existe pas, renvoie un objet null.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Renvoie un tableau d’adresses en style a1. La valeur Address contient le nom de la feuille de calcul pour chaque bloc rectangulaire de cellules (par exemple, «Sheet1 ! A1 : B4, Sheet1 ! D1 : D4 "). En lecture seule.|
||[Zones](/javascript/api/excel/excel.workbookrangeareas#areas)|Renvoie l' `RangeAreasCollection` objet. Chaque `RangeAreas` objet de la collection représente une ou plusieurs plages de rectangles dans une feuille de calcul.|
||[fourneau](/javascript/api/excel/excel.workbookrangeareas#ranges)|Renvoie les plages qui composent cet objet dans un `RangeCollection` objet.|
|[Feuille de calcul](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Obtient une collection de propriétés personnalisées au niveau de la feuille de calcul.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Supprime la propriété personnalisée.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Obtient la clé de la propriété personnalisée. Les clés de propriété personnalisée ne sont pas sensibles à la casse. La clé est limitée à 255 caractères (les plus grandes valeurs entraînent la levée d’une erreur « InvalidArgument »).|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Obtient ou définit la valeur de la propriété personnalisée.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key : chaîne, value : chaîne)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Ajoute une nouvelle propriété personnalisée qui est mappée à la clé fournie. Cette opération remplace les propriétés personnalisées existantes par cette clé.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Obtient le nombre de propriétés personnalisées sur cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Lève une exception si la propriété personnalisée n’existe pas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Obtient un objet de propriété personnalisé par sa clé, qui ne tient pas compte de la casse. Renvoie un objet null si la propriété personnalisée n’existe pas.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
