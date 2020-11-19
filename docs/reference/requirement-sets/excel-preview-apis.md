---
title: Version d’évaluation API JavaScript Excel
description: Informations détaillées sur les API JavaScript pour Excel à venir.
ms.date: 11/17/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 083741d35d3e881c2e46b186c4e93591bf7f4834
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131765"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Types de données liées | Prend en charge les types de données connectés à Excel à partir de sources externes. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Affichages de feuille nommée | Fournit un contrôle par programme des affichages de feuille de calcul par utilisateur. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| Tâches | Transformez les commentaires en tâches affectées aux utilisateurs. | [Tâche](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en version préliminaire. Pour obtenir la liste complète des API JavaScript pour Excel (dont les API d’aperçu et les API précédemment publiées), voir [toutes les API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Class | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (email : chaîne)](/javascript/api/excel/excel.comment#assigntask-email-)|Affecte la tâche jointe au commentaire à l’utilisateur donné en tant que cessionnaire unique.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Obtient la tâche associée à ce commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (email : chaîne)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Affecte la tâche jointe au commentaire à l’utilisateur donné en tant que cessionnaire unique.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Obtient la tâche associée à ce commentaire.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Nom du fournisseur de données pour le type de données liées.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Date et heure locales du fuseau horaire depuis l’ouverture du classeur lors de la dernière actualisation du type de données liées.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nom du type de données liées.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Fréquence, en secondes, à laquelle le type de données liées est actualisé si `refreshMode` est défini sur « périodique ».|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Mécanisme par lequel les données du type de données liées sont récupérées.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|ID unique du type de données liées.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Renvoie un tableau avec tous les modes d’actualisation pris en charge par le type de données liées.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Effectue une demande pour actualiser le type de données liées.|
||[requestSetRefreshMode (refreshMode : Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Effectue une demande pour modifier le mode d’actualisation de ce type de données liées.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|ID unique du nouveau type de données liées.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtient le type de l’événement.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtient le nombre de types de données liées dans la collection.|
||[getItem (Key : nombre)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtient un type de données liées par ID de service.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtient un type de données liées par son index dans la collection.|
||[getItemOrNullObject (Key : nombre)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtient un type de données liées par ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Effectue une demande d’actualisation de tous les types de données liées dans la collection.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Active l’affichage tableau.|
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
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Description du texte de remplacement du tableau croisé dynamique.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Titre de texte de remplacement du tableau croisé dynamique.|
||[displayBlankLineAfterEachItem (Display : Boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Définit si une ligne vide doit être affichée après chaque élément.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texte qui est rempli automatiquement dans une cellule vide du tableau croisé dynamique si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Indique si les cellules vides dans le tableau croisé dynamique doivent être renseignées avec le `emptyCellText` .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Style appliqué au tableau croisé dynamique.|
||[repeatAllItemLabels (repeatLabels : booléen)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Définit le paramètre « répéter toutes les étiquettes d’éléments » sur tous les champs du tableau croisé dynamique.|
||[setStyle (style : String \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Définit le style appliqué au tableau croisé dynamique.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Indique si le tableau croisé dynamique affiche les en-têtes de champ (légendes de champ et listes déroulantes de filtre).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Indique si le tableau croisé dynamique est actualisé lors de l’ouverture du classeur.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Renvoie un `WorkbookRangeAreas` Object qui représente la plage contenant tous les antécédents d’une cellule dans une même feuille de calcul ou dans plusieurs feuilles de calcul.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Mode d’actualisation du type de données liées.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|ID unique de l’objet dont le mode d’actualisation a été modifié.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtient le type de l’événement.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[Actualisé](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indique si la demande d’actualisation a réussi.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|ID unique de l’objet dont la demande d’actualisation a été exécutée.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtient le type de l’événement.|
||[affichés](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Tableau qui contient tous les avertissements générés à partir de la demande d’actualisation.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Style appliqué au Slicer.|
||[setStyle (style : String \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Définit le style appliqué au segment.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Style appliqué au tableau.|
||[setStyle (style : String \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Définit le style appliqué au tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtient l’ID de la table dans laquelle le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[Tâche](/javascript/api/excel/excel.task)|[addAssignee (email : chaîne)](/javascript/api/excel/excel.task#addassignee-email-)|Ajoute un cessionnaire à la tâche.|
||[applyChanges (taskChanges : Excel. TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Applique les modifications apportées à la tâche.|
||[utilisateurs](/javascript/api/excel/excel.task#assignees)|Obtient les utilisateurs auxquels la tâche est affectée.|
||[comment](/javascript/api/excel/excel.task#comment)|Obtient le commentaire associé à la tâche.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Obtient la date et l’heure d’échéance de la tâche.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Obtient les enregistrements d’historique de la tâche.|
||[id](/javascript/api/excel/excel.task#id)|Obtient l’ID de la tâche.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Obtient le pourcentage d’achèvement de la tâche.|
||[priorité](/javascript/api/excel/excel.task#priority)|Obtient la priorité de la tâche.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Obtient la date et l’heure de début de la tâche.|
||[title](/javascript/api/excel/excel.task#title)|Obtient le titre de la tâche.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Supprime tous les utilisateurs de la tâche.|
||[removeAssignee (email : chaîne)](/javascript/api/excel/excel.task#removeassignee-email-)|Supprime une personne affectée de la tâche.|
||[setPercentComplete (percentComplete : Number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Modifie l’exécution de la tâche.|
||[setPriority (priorité : nombre)](/javascript/api/excel/excel.task#setpriority-priority-)|Modifie la priorité de la tâche.|
||[setStartDateAndDueDate (DateDébut : date, dueDate : date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Modifie le début et les dates d’échéance de la tâche.|
||[setTitle (titre : chaîne)](/javascript/api/excel/excel.task#settitle-title-)|Modifie le titre de la tâche.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Définit une nouvelle date d’échéance pour la tâche, au format de fuseau horaire UTC.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Définit les adresses de messagerie des utilisateurs à affecter à la tâche.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Définit les adresses de messagerie des utilisateurs dont l’affectation doit être désaffectée de la tâche.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Définit un nouveau pourcentage d’achèvement de la tâche.|
||[priorité](/javascript/api/excel/excel.taskchanges#priority)|Définit une nouvelle priorité pour la tâche.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Définit si la modification doit supprimer tous les utilisateurs précédents de la tâche.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Définit une nouvelle date de début pour la tâche, au format de fuseau horaire UTC.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Définit un nouveau titre pour la tâche.|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Obtient le nombre de tâches dans la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Obtient une tâche à l’aide de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Obtient une tâche par son index dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Obtient une tâche à l’aide de son ID.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Représente l’ID de l’objet auquel la tâche est ancrée (par exemple, commentId pour les tâches jointes aux commentaires).|
||[utilisateur](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Représente l’utilisateur affecté à la tâche pour un type d’enregistrement d’historique « attribuer », ou l’utilisateur à annuler l’affectation de la tâche pour un type d’enregistrement d’historique « annuler l’affectation ».|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Représente l’utilisateur qui a créé ou modifié la tâche.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Représente la date d’échéance de la tâche.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Représente la date de création de l’enregistrement de l’historique des tâches.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID de l’enregistrement de l’historique.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Représente le pourcentage d’achèvement de la tâche.|
||[priorité](/javascript/api/excel/excel.taskhistoryrecord#priority)|Représente la priorité de la tâche.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Représente la date de début de la tâche.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Représente le titre de la tâche.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Représente le type de l’enregistrement de l’historique des tâches.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Représente la propriété TaskHistoryRecord.id qui a été annulée pour le type d’enregistrement d’historique « annuler ».|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Obtient le nombre d’enregistrements d’historique dans la collection de la tâche.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Obtient un enregistrement d’historique de tâche à l’aide de son index dans la collection.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Utilisateur](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Représente le nom d’affichage de l’utilisateur.|
||[email](/javascript/api/excel/excel.user#email)|Représente l’adresse e-mail de l’utilisateur.|
||[uid](/javascript/api/excel/excel.user#uid)|Représente l’ID unique de l’utilisateur.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Renvoie une collection de types de données liées qui font partie du classeur.|
||[décrites](/javascript/api/excel/excel.workbook#tasks)|Renvoie une collection de tâches présentes dans le classeur.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Indique si le volet de liste de champs du tableau croisé dynamique est affiché au niveau du classeur.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Renvoie une collection de vues de feuille présentes dans la feuille de calcul.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[décrites](/javascript/api/excel/excel.worksheet#tasks)|Renvoie une collection de tâches présentes dans la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
