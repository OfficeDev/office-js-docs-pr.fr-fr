---
title: Version d’évaluation API JavaScript Excel
description: Détails sur les API JavaScript Excel à venir.
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4bceda6229270332ed7624b693913e47a065a066
ms.sourcegitcommit: 3cc8f6adee0c7c68c61a42da0d97ed5ea61be0ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53661277"
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
| Workbooks liés | Gérer les liens entre les workbooks, y compris la prise en charge de l’actualisation et de la rupture des liens de ces derniers. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Styles de tableau | Permet de contrôler la police, la bordure, la couleur de remplissage et d’autres aspects des styles de tableau. | [Tableau](/javascript/api/excel/excel.table), [Tableau croisé dynamique](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Requêtes | Récupérer les attributs de requête tels que le nom, la date d’actualisation et le nombre de requêtes. | [Query](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les Excel api JavaScript actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript Excel (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript Excel.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteshiftdirection)|Représente la direction (par exemple, vers le haut ou vers la gauche) vers le haut ou vers la gauche que les cellules restantes déplacent lorsqu’une ou plusieurs cellules sont supprimées.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertshiftdirection)|Représente la direction (par exemple, vers le bas ou vers la droite) que les cellules existantes déplacent lorsqu’une ou plusieurs nouvelles cellules sont insérées.|
|[Graphique](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getdatatable--)|Obtient la table de données du graphique.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getdatatableornullobject--)|Obtient la table de données du graphique.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Représente le format d’un tableau de données de graphique, qui inclut le format de remplissage, de police et de bordure.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showhorizontalborder)|Spécifie s’il faut afficher la bordure horizontale de la table de données.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showlegendkey)|Spécifie s’il faut afficher le legendkey de la table de données.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showoutlineborder)|Spécifie s’il faut afficher la bordure plan de la table de données.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showverticalborder)|Spécifie s’il faut afficher la bordure verticale de la table de données.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Spécifie s’il faut afficher la table de données du graphique.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#border)|Représente le format de bordure du tableau de données de graphique, qui inclut la couleur, le style de trait et l’pondération.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.chartdatatableformat#font)|Représente les attributs de police (tels que le nom de la police, la taille de police et la couleur) de l’objet actuel.|
|[Commentaire](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que personne assignée.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Obtient la tâche associée à ce commentaire.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getitemornullobject-commentid-)|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assigntask-assignee-)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que seule personne assignée.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitemornullobject-commentreplyid-)|Renvoie une réponse de commentaire identifié via son ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getitemornullobject-id-)|Renvoie un format conditionnel identifié par son ID.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentcomplete)|Spécifie le pourcentage d’achèvement de la tâche.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Spécifie la priorité de la tâche.|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|Renvoie une collection de personnes assignées à la tâche.|
||[modifications](/javascript/api/excel/excel.documenttask#changes)|Obtient les enregistrements de modification de la tâche.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Obtient le commentaire associé à la tâche.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedby)|Obtient l’utilisateur le plus récent à avoir effectué la tâche.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completeddatetime)|Obtient la date et l’heure de fin de la tâche.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdby)|Obtient l’utilisateur qui a créé la tâche.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createddatetime)|Obtient la date et l’heure de création de la tâche.|
||[id](/javascript/api/excel/excel.documenttask#id)|Obtient l’ID de la tâche.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setstartandduedatetime-startdatetime--duedatetime-)|Modifie le début et les dates d’échéance de la tâche.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startandduedatetime)|Obtient ou définit la date et l’heure de début et d’échéance de la tâche.|
||[title](/javascript/api/excel/excel.documenttask#title)|Spécifie le titre de la tâche.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Représente l’utilisateur affecté à la tâche pour un type d’enregistrement de modification ou l’utilisateur non affecté à la tâche pour `assign` un type d’enregistrement de `unassign` modification.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedby)|Représente l’utilisateur qui a créé ou modifié la tâche.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentid)|Représente l’ID du ou des points d’ancrage de la `Comment` `CommentReply` modification de tâche.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createddatetime)|Représente la date et l’heure de création de l’enregistrement de modification de tâche.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#duedatetime)|Représente la date et l’heure d’échéance de la tâche, dans le fuseau horaire UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID de l’enregistrement de modification de tâche.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentcomplete)|Représente le pourcentage d’achèvement de la tâche.|
||[priority](/javascript/api/excel/excel.documenttaskchange#priority)|Représente la priorité de la tâche.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startdatetime)|Représente la date et l’heure de début de la tâche, dans le fuseau horaire UTC.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Représente le titre de la tâche.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Représente le type d’action de l’enregistrement de modification de tâche.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undohistoryid)|Représente la propriété `DocumentTaskChange.id` qui a été annulée pour le type `undo` d’enregistrement de modification.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getcount--)|Obtient le nombre d’enregistrements de modification dans la collection pour la tâche.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getitemat-index-)|Obtient un enregistrement de modification de tâche à l’aide de son index dans la collection.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getcount--)|Obtient le nombre de tâches dans la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitem-key-)|Obtient une tâche à l’aide de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getitemat-index-)|Obtient une tâche par son index dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getitemornullobject-key-)|Obtient une tâche à l’aide de son ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#duedatetime)|Obtient la date et l’heure d’échéance de la tâche.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startdatetime)|Obtient la date et l’heure de début de la tâche.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Identité](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|Représente le nom d’affichage de l’utilisateur.|
||[email](/javascript/api/excel/excel.identity#email)|Représente l’adresse e-mail de l’utilisateur.|
||[id](/javascript/api/excel/excel.identity#id)|Représente l’ID unique de l’utilisateur.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|Ajoute une identité d’utilisateur à la collection.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|Supprime toutes les identités utilisateur de la collection.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|Obtient le nombre d'éléments dans la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|Obtient une identité d’utilisateur de document à l’aide de son index dans la collection.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|Supprime une identité d’utilisateur de la collection.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayname)|Représente le nom d’affichage de l’utilisateur.|
||[email](/javascript/api/excel/excel.identityentity#email)|Représente l’adresse e-mail de l’utilisateur.|
||[id](/javascript/api/excel/excel.identityentity#id)|Représente l’ID unique de l’utilisateur.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Nom du fournisseur de données pour le type de données liées.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Date et heure du fuseau horaire local depuis l’ouverture du manuel lors de la dernière actualisation du type de données liées.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nom du type de données liées.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Fréquence, en secondes, à laquelle le type de données liées est actualisé si elle est définie `refreshMode` sur « Périodique ».|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Mécanisme par lequel les données du type de données liées sont récupérées.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|ID unique du type de données liées.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Renvoie un tableau avec tous les modes d’actualisation pris en charge par le type de données liées.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Effectue une demande d’actualisation du type de données liées.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Effectue une demande de modification du mode d’actualisation pour ce type de données liées.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|ID unique du nouveau type de données liées.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtient le type de l’événement.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtient le nombre de types de données liées dans la collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtient un type de données liées par ID de service.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtient un type de données liées par son index dans la collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtient un type de données liées par ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Effectue une demande d’actualisation de tous les types de données liées dans la collection.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breaklinks--)|Effectue une demande pour rompre les liens pointant vers le workbook lié.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|URL d’origine pointant vers le workbook lié.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh--)|Effectue une demande d’actualisation des données récupérées à partir du workbook lié.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakalllinks--)|Rompt tous les liens vers les workbooks liés.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitem-key-)|Obtient des informations sur un workbook lié par son URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getitemornullobject-key-)|Obtient des informations sur un workbook lié par son URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshall--)|Effectue une demande d’actualisation de tous les liens dubook.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbooklinksrefreshmode)|Représente le mode de mise à jour des liens du workbook.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|Obtient une vue de feuille à l’aide de son nom.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Style appliqué au tableau croisé dynamique.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Définit le style appliqué au tableau croisé dynamique.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|Obtient le premier tableau croisé dynamique de la collection.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Obtient le message d’erreur de requête à partir de la dernière actualisation de la requête.|
||[loadedTo](/javascript/api/excel/excel.query#loadedto)|Obtient le type d’objet de requête « chargé vers ».|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedtodatamodel)|Spécifie si la requête a été chargée dans le modèle de données.|
||[name](/javascript/api/excel/excel.query#name)|Obtient le nom de la requête.|
||[refreshDate](/javascript/api/excel/excel.query#refreshdate)|Obtient la date et l’heure de la dernière actualisation de la requête.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsloadedcount)|Obtient le nombre de lignes qui ont été chargées lors de la dernière actualisation de la requête.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getcount--)|Obtient le nombre de requêtes dans le manuel.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getitem-key-)|Obtient une requête de la collection en fonction de son nom.|
||[items](/javascript/api/excel/excel.querycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|Renvoie un objet qui représente la plage contenant tous les dépendants d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Renvoie un objet qui représente la plage contenant tous les antécédents d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Mode d’actualisation du type de données liées.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|ID unique de l’objet dont le mode d’actualisation a été modifié.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtient le type de l’événement.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[actualisé](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indique si la demande d’actualisation a réussi.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|ID unique de l’objet dont la demande d’actualisation a été effectuée.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtient le type de l’événement.|
||[avertissements](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Tableau qui contient les avertissements générés à partir de la demande d’actualisation.|
|[Forme](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayname)|Obtient le nom complet de la forme.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Style appliqué au slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Définit le style appliqué au slicer.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|Obtient un style par nom.|
|[Tableau](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsqu’un filtre est appliqué à une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Style appliqué au tableau.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Définit le style appliqué au tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsqu’un filtre est appliqué à une table d’un workbook ou d’une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtient l’ID du tableau dans lequel le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleterows-rows-)|Supprimez plusieurs lignes d’un tableau.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleterowsat-index--count-)|Supprimez un nombre spécifié de lignes d’un tableau, en commençant à un index donné.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[Classeur](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Renvoie une collection de types de données liées qui font partie du manuel.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedworkbooks)|Renvoie une collection de workbooks liés.|
||[requêtes](/javascript/api/excel/excel.workbook#queries)|Renvoie une collection de requêtes Power Query qui font partie du manuel.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Renvoie une collection de tâches présentes dans le manuel.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Spécifie si le volet liste des champs du tableau croisé dynamique est affiché au niveau du workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsqu’un filtre est appliqué sur une feuille de calcul spécifique.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onprotectionchanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
||[tabId](/javascript/api/excel/excel.worksheet#tabid)|Renvoie une valeur représentant cette feuille de calcul qui peut être lue par Open Office XML.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Renvoie une collection de tâches présentes dans la feuille de calcul.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changedirectionstate)|Représente une modification du sens de déplacement des cellules d’une feuille de calcul lorsqu’une ou plusieurs cellules sont supprimées ou insérées.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggersource)|Représente la source du déclencheur de l’événement.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onprotectionchanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isprotected)|Obtient l’état de protection actuel de la feuille de calcul.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle l’état de protection est modifié.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
