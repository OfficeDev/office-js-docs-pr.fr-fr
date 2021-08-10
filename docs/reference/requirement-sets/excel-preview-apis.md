---
title: Version d’évaluation API JavaScript Excel
description: Détails sur les API JavaScript Excel à venir.
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d90c5e8bb2c344cb3bb297a3cd793613f017e910ab99df6dfffc456c3f715d20
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092643"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

Le tableau suivant fournit un résumé concis des API, tandis que le tableau de liste [d’API](#api-list) suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Tables de données de graphique | Contrôler l’apparence, la mise en forme et la visibilité des tables de données sur les graphiques. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Tâches de document | Transformez les commentaires en tâches affectées aux utilisateurs. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identités | Gérer les identités des utilisateurs, y compris le nom d’affichage et l’adresse e-mail. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Types de données liées | Ajoute la prise en charge des types de données connectés Excel à partir de sources externes. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Workbooks liés | Gérez les liens entre les workbooks, notamment la prise en charge de l’actualisation et de la rupture des liens de ces derniers. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Styles de tableau | Permet de contrôler la police, la bordure, la couleur de remplissage et d’autres aspects des styles de tableau. | [Tableau](/javascript/api/excel/excel.table), [Tableau croisé dynamique](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Requêtes | Récupérer les attributs de requête tels que le nom, la date d’actualisation et le nombre de requêtes. | [Query](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les Excel api JavaScript actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript Excel (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)Excel.

| Classe | Champs | Description |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Représente la direction (par exemple, vers le haut ou vers la gauche) vers le haut ou vers la gauche que les cellules restantes déplacent lorsqu’une ou plusieurs cellules sont supprimées.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Représente la direction (par exemple, vers le bas ou vers la droite) que les cellules existantes déplacent lorsqu’une ou plusieurs nouvelles cellules sont insérées.|
|[Graphique](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Obtient la table de données du graphique.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Obtient la table de données du graphique.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Représente le format d’un tableau de données de graphique, qui inclut le format de remplissage, de police et de bordure.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Spécifie s’il faut afficher la bordure horizontale de la table de données.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Spécifie s’il faut afficher le legendkey de la table de données.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Spécifie s’il faut afficher la bordure plan de la table de données.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Spécifie s’il faut afficher la bordure verticale de la table de données.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Spécifie s’il faut afficher la table de données du graphique.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#border)|Représente le format de bordure de la table de données du graphique, qui inclut la couleur, le style de trait et l’pondération.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.chartdatatableformat#font)|Représente les attributs de police (tels que le nom de la police, la taille de police et la couleur) de l’objet actuel.|
|[Commentaire](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que personne assignée.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Obtient la tâche associée à ce commentaire.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que seule personne assignée.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Renvoie une réponse de commentaire identifié via son ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Renvoie un format conditionnel identifié par son ID.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Spécifie le pourcentage d’achèvement de la tâche.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Spécifie la priorité de la tâche.|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|Renvoie une collection de personnes assignées à la tâche.|
||[modifications](/javascript/api/excel/excel.documenttask#changes)|Obtient les enregistrements de modification de la tâche.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Obtient le commentaire associé à la tâche.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Obtient l’utilisateur le plus récent à avoir effectué la tâche.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Obtient la date et l’heure de fin de la tâche.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Obtient l’utilisateur qui a créé la tâche.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Obtient la date et l’heure de création de la tâche.|
||[id](/javascript/api/excel/excel.documenttask#id)|Obtient l’ID de la tâche.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setStartAndDueDateTime_startDateTime__dueDateTime_)|Modifie le début et les dates d’échéance de la tâche.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startAndDueDateTime)|Obtient ou définit la date et l’heure de début et d’échéance de la tâche.|
||[title](/javascript/api/excel/excel.documenttask#title)|Spécifie le titre de la tâche.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Représente l’utilisateur affecté à la tâche pour un type d’enregistrement de modification ou l’utilisateur non affecté à la tâche pour `assign` un type d’enregistrement de `unassign` modification.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedBy)|Représente l’utilisateur qui a créé ou modifié la tâche.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentId)|Représente l’ID du ou des points d’ancrage de la `Comment` `CommentReply` modification de tâche.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createdDateTime)|Représente la date et l’heure de création de l’enregistrement de modification de tâche.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#dueDateTime)|Représente la date et l’heure d’échéance de la tâche, dans le fuseau horaire UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID de l’enregistrement de modification de tâche.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentComplete)|Représente le pourcentage d’achèvement de la tâche.|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|Représente la priorité de la tâche.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startDateTime)|Représente la date et l’heure de début de la tâche, dans le fuseau horaire UTC.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Représente le titre de la tâche.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Représente le type d’action de l’enregistrement de modification de tâche.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undoHistoryId)|Représente la propriété `DocumentTaskChange.id` qui a été annulée pour le type `undo` d’enregistrement de modification.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getCount__)|Obtient le nombre d’enregistrements de modification dans la collection pour la tâche.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getItemAt_index_)|Obtient un enregistrement de modification de tâche à l’aide de son index dans la collection.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getCount__)|Obtient le nombre de tâches dans la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItem_key_)|Obtient une tâche à l’aide de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getItemAt_index_)|Obtient une tâche par son index dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItemOrNullObject_key_)|Obtient une tâche à l’aide de son ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#dueDateTime)|Obtient la date et l’heure d’échéance de la tâche.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startDateTime)|Obtient la date et l’heure de début de la tâche.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Identité](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayName)|Représente le nom d’affichage de l’utilisateur.|
||[email](/javascript/api/excel/excel.identity#email)|Représente l’adresse e-mail de l’utilisateur.|
||[id](/javascript/api/excel/excel.identity#id)|Représente l’ID unique de l’utilisateur.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add_assignee_)|Ajoute une identité d’utilisateur à la collection.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear__)|Supprime toutes les identités utilisateur de la collection.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getCount__)|Obtient le nombre d'éléments dans la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getItemAt_index_)|Obtient une identité d’utilisateur de document à l’aide de son index dans la collection.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove_assignee_)|Supprime une identité d’utilisateur de la collection.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayName)|Représente le nom d’affichage de l’utilisateur.|
||[email](/javascript/api/excel/excel.identityentity#email)|Représente l’adresse e-mail de l’utilisateur.|
||[id](/javascript/api/excel/excel.identityentity#id)|Représente l’ID unique de l’utilisateur.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataProvider)|Nom du fournisseur de données pour le type de données liées.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastRefreshed)|Date et heure du fuseau horaire local depuis l’ouverture du manuel lors de la dernière actualisation du type de données liées.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nom du type de données liées.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicRefreshInterval)|Fréquence, en secondes, à laquelle le type de données liées est actualisé si elle est définie `refreshMode` sur « Périodique ».|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshMode)|Mécanisme par lequel les données du type de données liées sont récupérées.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|ID unique du type de données liées.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Renvoie un tableau avec tous les modes d’actualisation pris en charge par le type de données liées.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Effectue une demande d’actualisation du type de données liées.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Effectue une demande de modification du mode d’actualisation pour ce type de données liées.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|ID unique du nouveau type de données liées.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtient le type de l’événement.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Obtient le nombre de types de données liées dans la collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Obtient un type de données liées par ID de service.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Obtient un type de données liées par son index dans la collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Obtient un type de données liées par ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Effectue une demande d’actualisation de tous les types de données liées dans la collection.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Effectue une demande pour rompre les liens pointant vers le workbook lié.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|URL d’origine pointant vers le workbook lié.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Effectue une demande d’actualisation des données récupérées à partir du workbook lié.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Rompt tous les liens vers les workbooks liés.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Obtient des informations sur un workbook lié par son URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Obtient des informations sur un workbook lié par son URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Effectue une demande d’actualisation de tous les liens dubook.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Représente le mode de mise à jour des liens du workbook.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Obtient une vue de feuille à l’aide de son nom.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|Style appliqué au tableau croisé dynamique.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Définit le style appliqué au tableau croisé dynamique.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Obtient le premier tableau croisé dynamique de la collection.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Obtient le message d’erreur de requête à partir de la dernière actualisation de la requête.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Obtient le type d’objet de requête « chargé vers ».|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Spécifie si la requête a été chargée dans le modèle de données.|
||[name](/javascript/api/excel/excel.query#name)|Obtient le nom de la requête.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Obtient la date et l’heure de la dernière actualisation de la requête.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Obtient le nombre de lignes qui ont été chargées lors de la dernière actualisation de la requête.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Obtient le nombre de requêtes dans le workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Obtient une requête de la collection en fonction de son nom.|
||[items](/javascript/api/excel/excel.querycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Renvoie un objet qui représente la plage contenant tous les dépendants d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Renvoie un objet qui représente la plage contenant tous les antécédents d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|Mode d’actualisation du type de données liées.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|ID unique de l’objet dont le mode d’actualisation a été modifié.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtient le type de l’événement.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[actualisé](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indique si la demande d’actualisation a réussi.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|ID unique de l’objet dont la demande d’actualisation a été effectuée.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtient le type de l’événement.|
||[avertissements](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Tableau qui contient les avertissements générés à partir de la demande d’actualisation.|
|[Forme](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|Obtient le nom complet de la forme.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|Style appliqué au slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Définit le style appliqué au slicer.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Obtient un style par nom.|
|[Tableau](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Se produit lorsqu’un filtre est appliqué à une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|Style appliqué au tableau.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Définit le style appliqué au tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Se produit lorsqu’un filtre est appliqué à une table d’un workbook ou d’une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Obtient l’ID du tableau dans lequel le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Supprimez plusieurs lignes d’un tableau.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Supprimez un nombre spécifié de lignes d’un tableau, en commençant à un index donné.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[Classeur](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Renvoie une collection de types de données liées qui font partie du manuel.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Renvoie une collection de workbooks liés.|
||[requêtes](/javascript/api/excel/excel.workbook#queries)|Renvoie une collection de requêtes Power Query qui font partie du manuel.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Renvoie une collection de tâches qui sont présentes dans le workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Spécifie si le volet liste des champs du tableau croisé dynamique est affiché au niveau du workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Se produit lorsqu’un filtre est appliqué à une feuille de calcul spécifique.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Renvoie une valeur représentant cette feuille de calcul qui peut être lue par Open Office XML.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Renvoie une collection de tâches présentes dans la feuille de calcul.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Représente une modification du sens de déplacement des cellules d’une feuille de calcul lorsqu’une ou plusieurs cellules sont supprimées ou insérées.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Représente la source du déclencheur de l’événement.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Obtient l’état de protection actuel de la feuille de calcul.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle l’état de protection est modifié.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
