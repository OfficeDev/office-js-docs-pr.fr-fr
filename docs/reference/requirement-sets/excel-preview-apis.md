---
title: Version d’évaluation API JavaScript Excel
description: Détails sur les API JavaScript Excel à venir.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e4e0066830d10ad3b466d33b5a59d31efe4b9777
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604658"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

Le tableau suivant fournit un résumé concis des API, tandis que le tableau de liste [d’API](#api-list) suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Tâches de document | Transformez les commentaires en tâches affectées aux utilisateurs. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Événements de changement de formule | Suivre les modifications apportées aux formules, y compris la source et le type d’événement à l’origine d’une modification. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| Types de données liées | Ajoute la prise en charge des types de données connectés à Excel à partir de sources externes. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| PivotLayout de tableau croisé dynamique | Développement de la classe PivotLayout, y compris la nouvelle prise en charge du texte de alt et de la gestion des cellules vides. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |
| Styles de tableau | Permet de contrôler la police, la bordure, la couleur de remplissage et d’autres aspects des styles de tableau. | [Tableau](/javascript/api/excel/excel.table), [Tableau croisé dynamique](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les API JavaScript pour Excel actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript pour Excel (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript pour Excel.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|Efface les critères de filtre du filtre automatique.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assigntask-assignee-)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que personne assignée.|
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
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Adresse de la cellule qui contient la formule modifiée.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Représente la formule précédente, avant qu’elle n’a été modifiée.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getitemornullobject-key-)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Identité](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayname)|Représente le nom d’affichage de l’utilisateur.|
||[email](/javascript/api/excel/excel.identity#email)|Représente l’adresse e-mail de l’utilisateur.|
||[id](/javascript/api/excel/excel.identity#id)|Représente l’ID unique de l’utilisateur.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add-assignee-)|Ajoute une identité d’utilisateur à la collection.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear--)|Supprime toutes les identités utilisateur de la collection.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getcount--)|Obtient le nombre d'éléments dans la collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getitemat-index-)|Obtient une identité d’utilisateur de document à l’aide de son index dans la collection.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove-assignee-)|Supprime une identité d’utilisateur de la collection.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|Position d’insertion, dans le livre de calcul actuel, des nouvelles feuilles de calcul.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|Feuille de calcul du manuel actuel référencé pour le `WorksheetPositionType` paramètre.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|Noms des feuilles de calcul individuelles à insérer.|
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
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtient un type de données lié par ID de service.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtient un type de données liées par son index dans la collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtient un type de données liées par ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Effectue une demande d’actualisation de tous les types de données liées dans la collection.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitemornullobject-key-)|Obtient une vue de feuille à l’aide de son nom.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|Description de texte de alt du tableau croisé dynamique.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|Titre de texte de alt du tableau croisé dynamique.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Définit si une ligne vide doit être affichée après chaque élément.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Texte qui est automatiquement rempli dans n’importe quelle cellule vide du tableau croisé dynamique si `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Spécifie si les cellules vides du tableau croisé dynamique doivent être remplies avec le `emptyCellText` .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Style appliqué au tableau croisé dynamique.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Définit le paramètre « Répéter toutes les étiquettes d’éléments » sur tous les champs du tableau croisé dynamique.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Définit le style appliqué au tableau croisé dynamique.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Spécifie si le tableau croisé dynamique affiche les en-têtes de champ (légendes de champ et les drop-downs de filtre).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Spécifie si le tableau croisé dynamique est actualisé à l’ouverture du manuel.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getfirstornullobject--)|Obtient le premier tableau croisé dynamique de la collection.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getdependents--)|Renvoie un objet qui représente la plage contenant tous les dépendants d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Renvoie un objet qui représente la plage contenant tous les dépendants directs d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Renvoie un objet RangeAreas qui représente les zones fusionnées dans cette plage.|
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
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getitemornullobject-key-)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|Style appliqué au slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Définit le style appliqué au slicer.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getitemornullobject-name-)|Obtient un style par nom.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Se produit lorsqu’un filtre est appliqué à une table spécifique.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Style appliqué au tableau.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Définit le style appliqué au tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Se produit lorsqu’un filtre est appliqué à une table d’un workbook ou d’une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtient l’ID du tableau dans lequel le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitemornullobject-key-)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Insère les feuilles de calcul spécifiées à partir d’un workbook source dans le workbook actuel.|
||[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Renvoie une collection de types de données liées qui font partie du manuel.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Se produit lorsque le workbook est activé.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Renvoie une collection de tâches qui sont présentes dans le workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Spécifie si le volet liste des champs du tableau croisé dynamique est affiché au niveau du workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True si le classeur utilise le calendrier depuis 1904.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Obtient le type de l’événement.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Se produit lorsqu’un filtre est appliqué sur une feuille de calcul spécifique.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans cette feuille de calcul.|
||[tabId](/javascript/api/excel/excel.worksheet#tabid)|Renvoie une valeur représentant cette feuille de calcul qui peut être lue par Open Office XML.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Renvoie une collection de tâches présentes dans la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Se produit lorsqu’une ou plusieurs formules sont modifiées dans une feuille de calcul de cette collection.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Obtient un tableau `FormulaChangedEventDetail` d’objets, qui contient les détails sur toutes les formules modifiées.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Obtient l’ID de la feuille de calcul dans laquelle la formule a été modifiée.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
