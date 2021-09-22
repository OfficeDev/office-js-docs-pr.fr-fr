---
title: Version d’évaluation API JavaScript Excel
description: Détails sur les API JavaScript Excel à venir.
ms.date: 09/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: bd36d9ba1be4e9e0caafdd49e63d8e7cdea01c59
ms.sourcegitcommit: a854a2fd2ad9f379a3ef712f307e0b1bb9b5b00d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/22/2021
ms.locfileid: "59474349"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

Le tableau suivant fournit un résumé concis des API, tandis que le tableau de liste [d’API](#api-list) suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| Tables de données de graphique | Contrôler l’apparence, la mise en forme et la visibilité des tables de données sur les graphiques. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Types de données personnalisés | Extension des types de données Excel existants, y compris la prise en charge des nombres formatés et des images web. | [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| Erreurs de types de données personnalisés| Objets d’erreur qui la prise en charge des types de données personnalisés. | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValuee](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| Tâches de document | Transformez les commentaires en tâches affectées aux utilisateurs. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identités | Gérer les identités des utilisateurs, y compris le nom d’affichage et l’adresse e-mail. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Types de données liées | Ajoute la prise en charge des types de données connectés Excel à partir de sources externes. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Styles de tableau | Permet de contrôler la police, la bordure, la couleur de remplissage et d’autres aspects des styles de tableau. | [Tableau](/javascript/api/excel/excel.table), [Tableau croisé dynamique](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Requêtes | Récupérer les attributs de requête tels que le nom, la date d’actualisation et le nombre de requêtes. | [Query](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| Protection de feuille de calcul | Empêcher les utilisateurs non autorisés d’apporter des modifications à des plages spécifiées dans une feuille de calcul. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les EXCEL JavaScript actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript Excel (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)Excel.

| Classe | Champs | Description |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[adresse](/javascript/api/excel/excel.alloweditrange#address)|Spécifie la plage associée à l’objet.|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|Supprime cet objet du `AllowEditRangeCollection` .|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|Suspend la protection de feuille de calcul pour `AllowEditRange` l’objet donné pour l’utilisateur dans une session donnée.|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|Spécifie si le mot `AllowEditRange` de passe est protégé.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|Modifie le mot de passe associé au `AllowEditRange` .|
||[title](/javascript/api/excel/excel.alloweditrange#title)|Spécifie le titre de l’objet.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel. AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|Ajoute un `AllowEditRange` objet à la collection.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|Renvoie le nombre `AllowEditRange` d’objets de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|Obtient `AllowEditRange` l’objet par son titre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|Renvoie un `AllowEditRange` objet par son index dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|Obtient `AllowEditRange` l’objet par son titre.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|Suspend la protection de feuille de calcul pour tous les objets de la collection qui ont le mot de passe donné pour l’utilisateur `AllowEditRange` dans une session donnée.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[mot de passe](/javascript/api/excel/excel.alloweditrangeoptions#password)|Mot de passe associé au `AllowEditRange` .|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|Représente le type de `BlockedErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.blockederrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.blockederrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[primitive](/javascript/api/excel/excel.booleancellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.booleancellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.booleancellvalue#type)|Représente le type de cette valeur de cellule.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|Représente le type de `BusyErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.busyerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.busyerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|Représente le type de `CalcErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.calcerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.calcerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|Représente une URL vers une licence ou une source qui décrit comment cette propriété peut être utilisée.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|Représente un nom pour la licence qui régit cette propriété.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|Représente une URL vers la source du `CellValue` .|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|Représente un nom pour la source du `CellValue` .|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|Représente la propriété de description du fournisseur utilisée en affichage carte si aucun logo n’est spécifié.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|Représente une URL utilisée pour télécharger une image qui sera utilisée comme logo en affichage carte.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|Représente une URL qui est la cible de navigation si l’utilisateur clique sur l’élément logo en affichage carte.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Représente la direction (par exemple, vers le haut ou vers la gauche) vers le haut ou vers la gauche que les cellules restantes déplacent lorsqu’une ou plusieurs cellules sont supprimées.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Représente la direction (par exemple, vers le bas ou vers la droite) que les cellules existantes déplacent lorsqu’une ou plusieurs nouvelles cellules sont insérées.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Obtient la table de données du graphique.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Obtient la table de données du graphique.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Représente le format d’un tableau de données de graphique, qui inclut le format de remplissage, de police et de bordure.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Spécifie s’il faut afficher la bordure horizontale de la table de données.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Spécifie s’il faut afficher le clé de légende de la table de données.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Spécifie s’il faut afficher la bordure de plan de la table de données.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Spécifie s’il faut afficher la bordure verticale de la table de données.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Spécifie s’il faut afficher la table de données du graphique.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[bordure](/javascript/api/excel/excel.chartdatatableformat#border)|Représente le format de bordure de la table de données du graphique, qui inclut la couleur, le style de trait et l’pondération.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Représente le format de remplissage d’un objet, qui comprend des informations de mise en forme d’arrière-plan.|
||[police](/javascript/api/excel/excel.chartdatatableformat#font)|Représente les attributs de police (tels que le nom de la police, la taille de police et la couleur) de l’objet actuel.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que personne assignée.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Obtient la tâche associée à ce commentaire.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Obtient un commentaire à partir de la collection de sites en fonction de son ID.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que seule personne assignée.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Renvoie une réponse de commentaire identifié via son ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Renvoie un format conditionnel identifié par son ID.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|Représente le type de `ConnectErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.connecterrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.connecterrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.div0errorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.div0errorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.div0errorcellvalue#type)|Représente le type de cette valeur de cellule.|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[primitive](/javascript/api/excel/excel.doublecellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.doublecellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.doublecellvalue#type)|Représente le type de cette valeur de cellule.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[primitive](/javascript/api/excel/excel.emptycellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.emptycellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.emptycellvalue#type)|Représente le type de cette valeur de cellule.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|Représente le type de `FieldErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.fielderrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.fielderrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|Renvoie la chaîne de format numérique utilisée pour afficher cette valeur.|
||[primitive](/javascript/api/excel/excel.formattednumbercellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.formattednumbercellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#type)|Représente le type de cette valeur de cellule.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.gettingdataerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.gettingdataerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
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
|[NaErrorCellValue](/javascript/api/excel/excel.naerrorcellvalue)|[errorType](/javascript/api/excel/excel.naerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.naerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.naerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.naerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.nameerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.nameerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Obtient une vue de feuille à l’aide de son nom.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.nullerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.nullerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.numerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.numerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.numerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|Style appliqué au tableau croisé dynamique.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Définit le style appliqué au tableau croisé dynamique.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Obtient le premier tableau croisé dynamique de la collection.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Obtient le message d’erreur de requête à partir de la dernière actualisation de la requête.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Obtient la requête chargée dans le type d’objet.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Spécifie si la requête a été chargée dans le modèle de données.|
||[name](/javascript/api/excel/excel.query#name)|Obtient le nom de la requête.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Obtient la date et l’heure de la dernière actualisation de la requête.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Obtient le nombre de lignes qui ont été chargées lors de la dernière actualisation de la requête.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Obtient le nombre de requêtes dans le workbook.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Obtient une requête de la collection en fonction de son nom.|
||[items](/javascript/api/excel/excel.querycollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Renvoie un objet qui représente la plage contenant tous les dépendants d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
||[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Renvoie un objet qui représente la plage contenant tous les antécédents d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|Représente le type de `RefErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.referrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.referrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.referrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|Mode d’actualisation du type de données liées.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|ID unique de l’objet dont le mode d’actualisation a été modifié.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtient le type de l’événement.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[actualisé](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indique si la demande d’actualisation a réussi.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|ID unique de l’objet dont la demande d’actualisation a été effectuée.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtient le type de l’événement.|
||[avertissements](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Tableau qui contient les avertissements générés à partir de la demande d’actualisation.|
|[Shape](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|Obtient le nom complet de la forme.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|Crée un graphique de fichiers SVG (SVG) à partir d’une chaîne XML et il est ajouté à la feuille de calcul.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Obtient une forme à l’aide de son nom ou de son ID.|
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Représente le nom du segment utilisé dans la formule.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|Style appliqué au slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Définit le style appliqué au slicer.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|Représente le type de `SpillErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.spillerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.spillerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[primitive](/javascript/api/excel/excel.stringcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.stringcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.stringcellvalue#type)|Représente le type de cette valeur de cellule.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Obtient un style par nom.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Modifie le tableau pour utiliser le style de tableau par défaut.|
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
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|Représente le type de `ValueErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[primitive](/javascript/api/excel/excel.valueerrorcellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.valueerrorcellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[primitive](/javascript/api/excel/excel.valuetypenotavailablecellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|Représente le type de cette valeur de cellule.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[adresse](/javascript/api/excel/excel.webimagecellvalue#address)|Représente l’URL à partir de laquelle l’image sera téléchargée.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|Représente le texte de remplacement qui peut être utilisé dans les scénarios d’accessibilité pour décrire ce que représente l’image.|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#attribution)|Représente les informations d’attribution pour décrire les exigences en matière de source et de licence pour l’utilisation de cette image.|
||[primitive](/javascript/api/excel/excel.webimagecellvalue#primitive)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[primitiveType](/javascript/api/excel/excel.webimagecellvalue#primitiveType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[fournisseur](/javascript/api/excel/excel.webimagecellvalue#provider)|Représente des informations qui décrivent l’entité ou la personne qui a fourni l’image.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|Représente l’URL d’une page web avec des images qui sont considérées comme étant liées à `WebImageCellValue` ce .|
||[type](/javascript/api/excel/excel.webimagecellvalue#type)|Représente le type de cette valeur de cellule.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Renvoie une collection de types de données liées qui font partie du manuel.|
||[requêtes](/javascript/api/excel/excel.workbook#queries)|Renvoie une collection de requêtes Power Query qui font partie du manuel.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Renvoie une collection de tâches qui sont présentes dans le workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Spécifie si le volet liste des champs du tableau croisé dynamique est affiché au niveau du workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Se produit lorsqu’un filtre est appliqué à une feuille de calcul spécifique.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|Se produit lorsque le nom de la feuille de calcul est modifié.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|Se produit lorsque la visibilité de la feuille de calcul est modifiée.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Renvoie une valeur représentant cette feuille de calcul qui peut être lue par Open Office XML.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Renvoie une collection de tâches présentes dans la feuille de calcul.|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Représente une modification du sens de déplacement des cellules d’une feuille de calcul lorsqu’une ou plusieurs cellules sont supprimées ou insérées.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Représente la source du déclencheur de l’événement.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
||[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|Se produit lorsqu’une feuille de calcul est déplacée par un utilisateur dans un workbook.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|Se produit lorsque le nom de la feuille de calcul est modifié dans la collection de feuilles de calcul.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Se produit lorsque l’état de protection de la feuille de calcul est modifié.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|Se produit lorsque la visibilité de la feuille de calcul est modifiée dans la collection de feuilles de calcul.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|Obtient la nouvelle position de la feuille de calcul, après le déplacement.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|Obtient la position précédente de la feuille de calcul, avant le déplacement.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul qui a été déplacée.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|Obtient le nouveau nom de la feuille de calcul, après la modification du nom.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|Obtient le nom précédent de la feuille de calcul, avant que le nom ne soit modifié.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul avec le nouveau nom.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|Spécifie si le mot de passe peut être utilisé pour déverrouiller la protection de feuille de calcul.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|Suspend la protection de feuille de calcul pour l’objet de feuille de calcul donné pour l’utilisateur dans une session donnée.|
||[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|Spécifie les `AllowEditRangeCollection` trouvés dans cette feuille de calcul.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|Spécifie si la protection peut être suspendue pour cette feuille de calcul.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|Spécifie si la feuille est protégée par mot de passe.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|Spécifie si la protection de feuille de calcul est suspendue.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|Reprend la protection de feuille de calcul pour l’objet de feuille de calcul donné pour l’utilisateur dans une session donnée.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|Modifie le mot de passe associé à `WorksheetProtection` l’objet.|
||[updateOptions(options: Excel. WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|Modifiez les options de protection de feuille de calcul associées à `WorksheetProtection` l’objet.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|Indique si l’un des `AllowEditRange` objets a changé.|
||[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Obtient l’état de protection actuel de la feuille de calcul.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|Spécifie si les changements `WorksheetProtectionOptions` ont été faits.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|Indique si le mot de passe de la feuille de calcul a changé.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle l’état de protection est modifié.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|Source de l'événement.|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|Obtient le type de l’événement.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|Obtient le nouveau paramètre de visibilité de la feuille de calcul, après la modification de la visibilité.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|Obtient le paramètre de visibilité précédent de la feuille de calcul, avant la modification de la visibilité.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dont la visibilité a changé.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
