---
title: Version d’évaluation API JavaScript Excel
description: Détails sur les API JavaScript Excel à venir.
ms.date: 01/26/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 10057123cc159af0c00a6b6e6345d8f6ab316822
ms.sourcegitcommit: 3123b9819c5225ee45a5312f64be79e46cbd0e3c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/29/2021
ms.locfileid: "50043896"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

Le premier tableau fournit un résumé concis des API, tandis que le tableau suivant fournit une liste détaillée.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Types de données liées | Ajoute la prise en charge des types de données connectés à Excel à partir de sources externes. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Vues de feuille nommée | Permet de contrôler par programme les affichages de feuille de calcul par utilisateur. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| Tâches | Transformez les commentaires en tâches affectées aux utilisateurs. | [Tâche](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript pour Excel (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript pour Excel.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(email: string)](/javascript/api/excel/excel.comment#assigntask-email-)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que seule personne assignée.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Obtient la tâche associée à ce commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(email: string)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que seule personne assignée.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Obtient la tâche associée à ce commentaire.|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Adresse de la cellule qui contient la formule modifiée.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Représente la formule précédente, avant qu’elle n’a été modifiée.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|Nom du fournisseur de données pour le type de données liées.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|Date et heure du fuseau horaire local depuis l’ouverture du manuel lors de la dernière actualisation du type de données liées.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|Nom du type de données liées.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|Fréquence, en secondes, à laquelle le type de données liées est actualisé si elle est définie `refreshMode` sur « Périodique ».|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|Mécanisme par lequel les données du type de données liées sont récupérées.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|ID unique du type de données liées.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Renvoie un tableau avec tous les modes d’actualisation pris en charge par le type de données liées.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Effectue une demande d’actualisation du type de données liées.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Effectue une demande de modification du mode d’actualisation pour ce type de données liées.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|ID unique du nouveau type de données liées.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtient le type de l’événement.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtient le nombre de types de données liées dans la collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtient un type de données liées par ID de service.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtient un type de données liées par son index dans la collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtient un type de données liées par ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Effectue une demande d’actualisation de tous les types de données liées dans la collection.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Active cette vue de feuille.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Supprime l’affichage Feuille de la feuille de calcul.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Crée une copie de cette vue de feuille.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtient ou définit le nom de l’affichage Feuille.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Crée un affichage feuille avec le nom donné.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Crée et active un nouvel affichage de feuille temporaire.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Quitte l’affichage feuille actif.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtient l’affichage feuille de calcul actif.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtient le nombre d’affichages de feuille dans cette feuille de calcul.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtient une vue de feuille à l’aide de son nom.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtient une vue de feuille par son index dans la collection.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Description de texte de alt du tableau croisé dynamique.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Titre de texte de alt du tableau croisé dynamique.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Définit si une ligne vide doit être affichée après chaque élément.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texte qui est automatiquement rempli dans n’importe quelle cellule vide du tableau croisé dynamique si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Spécifie si les cellules vides du tableau croisé dynamique doivent être remplies avec le `emptyCellText` .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Style appliqué au tableau croisé dynamique.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Définit le paramètre « Répéter toutes les étiquettes d’éléments » dans tous les champs du tableau croisé dynamique.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Définit le style appliqué au tableau croisé dynamique.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Spécifie si le tableau croisé dynamique affiche les en-têtes de champ (légendes de champ et les drop-downs de filtre).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Spécifie si le tableau croisé dynamique est actualisé à l’ouverture du manuel.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Renvoie un objet qui représente la plage contenant tous les antécédents d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|Mode d’actualisation du type de données liées.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|ID unique de l’objet dont le mode d’actualisation a été modifié.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtient le type de l’événement.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[actualisé](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indique si la demande d’actualisation a réussi.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|ID unique de l’objet dont la demande d’actualisation a été effectuée.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtient le type de l’événement.|
||[avertissements](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Tableau qui contient les avertissements générés à partir de la demande d’actualisation.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Style appliqué au slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Définit le style appliqué au slicer.|
|[Tableau](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsque le filtre est appliqué sur une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Style appliqué au tableau.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Définit le style appliqué au tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsque le filtre est appliqué sur n’importe quel tableau dans un classeur ou une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtient l’id du tableau dans lequel le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtient l’id de la feuille de calcul qui contient le tableau.|
|[Tâche](/javascript/api/excel/excel.task)|[addAssignee(email: string)](/javascript/api/excel/excel.task#addassignee-email-)|Ajoute une personne assignée à la tâche.|
||[applyChanges(taskChanges: Excel.TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Applique les modifications données à la tâche.|
||[assignees](/javascript/api/excel/excel.task#assignees)|Obtient les utilisateurs auxquels la tâche est affectée.|
||[comment](/javascript/api/excel/excel.task#comment)|Obtient le commentaire associé à la tâche.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Obtient la date et l’heure d’échéance de la tâche.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Obtient les enregistrements d’historique de la tâche.|
||[id](/javascript/api/excel/excel.task#id)|Obtient l’ID de la tâche.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Obtient le pourcentage d’achèvement de la tâche.|
||[priorité](/javascript/api/excel/excel.task#priority)|Obtient la priorité de la tâche.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Obtient la date et l’heure de début de la tâche.|
||[title](/javascript/api/excel/excel.task#title)|Obtient le titre de la tâche.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Supprime tous les personnes assignées de la tâche.|
||[removeAssignee(email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|Supprime une personne assignée de la tâche.|
||[setPercentComplete(percentComplete: number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Modifie l’achèvement de la tâche.|
||[setPriority(priority: number)](/javascript/api/excel/excel.task#setpriority-priority-)|Modifie la priorité de la tâche.|
||[setStartDateAndDueDate(startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Modifie le début et les dates d’échéance de la tâche.|
||[setTitle(title: string)](/javascript/api/excel/excel.task#settitle-title-)|Modifie le titre de la tâche.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Définit une nouvelle date d’échéance pour la tâche, dans le fuseau horaire UTC.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Définit les adresses de messagerie des utilisateurs à affecter à la tâche.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Définit les adresses e-mail des utilisateurs à désattribuer à la tâche.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Définit un nouveau pourcentage d’achèvement pour la tâche.|
||[priorité](/javascript/api/excel/excel.taskchanges#priority)|Définit une nouvelle priorité pour la tâche.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Définit si la modification doit supprimer tous les anciens assignés de la tâche.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Définit une nouvelle date de début pour la tâche, dans le fuseau horaire UTC.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Définit un nouveau titre pour la tâche.|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Obtient le nombre de tâches dans la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Obtient une tâche à l’aide de son ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Obtient une tâche par son index dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Obtient une tâche à l’aide de son ID.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Représente l’ID de l’objet auquel la tâche est ancrée (par exemple, commentId pour les tâches jointes aux commentaires).|
||[assignee](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Représente l’utilisateur affecté à la tâche pour un type d’enregistrement d’historique « Assigner » ou l’utilisateur à désattribuer à la tâche pour un type d’enregistrement d’historique « Non affecté ».|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Représente l’utilisateur qui a créé ou modifié la tâche.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Représente la date d’échéance de la tâche.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Représente la date de création de l’enregistrement d’historique des tâches.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID de l’enregistrement d’historique.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Représente le pourcentage d’achèvement de la tâche.|
||[priorité](/javascript/api/excel/excel.taskhistoryrecord#priority)|Représente la priorité de la tâche.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Représente la date de début de la tâche.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Représente le titre de la tâche.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Représente le type d’enregistrement de l’historique des tâches.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Représente la propriété TaskHistoryRecord.id qui a été annulée pour le type d’enregistrement d’historique « Annuler ».|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Obtient le nombre d’enregistrements d’historique dans la collection pour la tâche.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Obtient un enregistrement d’historique des tâches à l’aide de son index dans la collection.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Utilisateur](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Représente le nom d’affichage de l’utilisateur.|
||[email](/javascript/api/excel/excel.user#email)|Représente l’adresse e-mail de l’utilisateur.|
||[uid](/javascript/api/excel/excel.user#uid)|Représente l’ID unique de l’utilisateur.|
|[Classeur](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Renvoie une collection de types de données liées qui font partie du manuel.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Renvoie une collection de tâches présentes dans le manuel.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Spécifie si le volet liste des champs du tableau croisé dynamique est affiché au niveau du workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Renvoie une collection d’affichages de feuille présents dans la feuille de calcul.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsque le filtre est appliqué sur un tableau spécifique.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans cette feuille de calcul.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Renvoie une collection de tâches présentes dans la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans une feuille de calcul de cette collection.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtient l’id de la feuille de calcul dans laquelle le filtre est appliqué.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Obtient un tableau d’objets FormulaChangedEventDetail, qui contiennent les détails sur toutes les formules modifiées.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle la formule a été modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
