---
title: Version d’évaluation API JavaScript Excel
description: Détails sur les API JavaScript Excel à venir.
ms.date: 11/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f62b71532f323ad17f541979d3956f217ab5d07d
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/03/2021
ms.locfileid: "60683783"
---
# <a name="excel-javascript-preview-apis"></a>Version d’évaluation API JavaScript Excel

Les nouvelles Excel JavaScript APIs introduits dans « Aperçu » et versions ultérieures deviennent partie intégrante d’une exigence spécifique, numérotée définie après que test suffisantes se produit et les commentaires des utilisateurs sont acquis.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

Le tableau suivant fournit un résumé concis des API, tandis que le tableau de liste [d’API](#api-list) suivant fournit une liste détaillée.

| Fonctionnalité | Description | Objets pertinents |
|:--- |:--- |:--- |
| [Types de données](../../excel/excel-data-types-overview.md) | Extension des types de données Excel existants, y compris la prise en charge des nombres formatés et des images web. | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue), [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [Erreurs de types de données](../../excel/excel-data-types-concepts.md#improved-error-support) | Objets d’erreur qui supportent les types de données développés. | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValuee](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| Tâches de document | Transformez les commentaires en tâches affectées aux utilisateurs. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identités | Gérer les identités des utilisateurs, y compris le nom d’affichage et l’adresse e-mail. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Types de données liées | Ajoute la prise en charge des types de données connectés Excel à partir de sources externes. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |
| Styles de tableau | Permet de contrôler la police, la bordure, la couleur de remplissage et d’autres aspects des styles de tableau. | [Tableau](/javascript/api/excel/excel.table), [Tableau croisé dynamique](/javascript/api/excel/excel.pivottable), [Slicer](/javascript/api/excel/excel.slicer) |
| Protection de feuille de calcul | Empêcher les utilisateurs non autorisés d’apporter des modifications à des plages spécifiées dans une feuille de calcul. | [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection), [AllowEditRange](/javascript/api/excel/excel.alloweditrange), [AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection), [AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions) |

## <a name="api-list"></a>Liste des API

Le tableau suivant répertorie les Excel api JavaScript actuellement en prévisualisation. Pour obtenir la liste complète de toutes les API JavaScript Excel (y compris les API de prévisualisation et les API publiées précédemment), voir toutes les API [JavaScript Excel.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Classe | Champs | Description |
|:---|:---|:---|
|[AllowEditRange](/javascript/api/excel/excel.alloweditrange)|[adresse](/javascript/api/excel/excel.alloweditrange#address)|Spécifie la plage associée à l’objet.|
||[delete()](/javascript/api/excel/excel.alloweditrange#delete__)|Supprime cet objet du `AllowEditRangeCollection` .|
||[isPasswordProtected](/javascript/api/excel/excel.alloweditrange#isPasswordProtected)|Spécifie si le mot `AllowEditRange` de passe est protégé.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.alloweditrange#pauseProtection_password_)|Suspend la protection de feuille de calcul pour `AllowEditRange` l’objet donné pour l’utilisateur dans une session donnée.|
||[setPassword(password?: string)](/javascript/api/excel/excel.alloweditrange#setPassword_password_)|Modifie le mot de passe associé au `AllowEditRange` .|
||[title](/javascript/api/excel/excel.alloweditrange#title)|Spécifie le titre de l’objet.|
|[AllowEditRangeCollection](/javascript/api/excel/excel.alloweditrangecollection)|[add(title: string, rangeAddress: string, options?: Excel. AllowEditRangeOptions)](/javascript/api/excel/excel.alloweditrangecollection#add_title__rangeAddress__options_)|Ajoute un `AllowEditRange` objet à la collection.|
||[getCount()](/javascript/api/excel/excel.alloweditrangecollection#getCount__)|Renvoie le nombre `AllowEditRange` d’objets de la collection.|
||[getItem(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItem_key_)|Obtient `AllowEditRange` l’objet par son titre.|
||[getItemAt(index: number)](/javascript/api/excel/excel.alloweditrangecollection#getItemAt_index_)|Renvoie un `AllowEditRange` objet par son index dans la collection.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.alloweditrangecollection#getItemOrNullObject_key_)|Obtient `AllowEditRange` l’objet par son titre.|
||[items](/javascript/api/excel/excel.alloweditrangecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[pauseProtection(password: string)](/javascript/api/excel/excel.alloweditrangecollection#pauseProtection_password_)|Suspend la protection de feuille de calcul pour tous les objets de la collection qui ont le mot de passe donné pour l’utilisateur `AllowEditRange` dans une session donnée.|
|[AllowEditRangeOptions](/javascript/api/excel/excel.alloweditrangeoptions)|[mot de passe](/javascript/api/excel/excel.alloweditrangeoptions#password)|Mot de passe associé au `AllowEditRange` .|
|[ArrayCellValue](/javascript/api/excel/excel.arraycellvalue)|[basicType](/javascript/api/excel/excel.arraycellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.arraycellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[éléments](/javascript/api/excel/excel.arraycellvalue#elements)|Représente les éléments du tableau.|
||[type](/javascript/api/excel/excel.arraycellvalue#type)|Représente le type de cette valeur de cellule.|
|[BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue)|[basicType](/javascript/api/excel/excel.blockederrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.blockederrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.blockederrorcellvalue#errorSubType)|Représente le type de `BlockedErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.blockederrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.blockederrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[BooleanCellValue](/javascript/api/excel/excel.booleancellvalue)|[basicType](/javascript/api/excel/excel.booleancellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.booleancellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.booleancellvalue#type)|Représente le type de cette valeur de cellule.|
|[BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue)|[basicType](/javascript/api/excel/excel.busyerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.busyerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.busyerrorcellvalue#errorSubType)|Représente le type de `BusyErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.busyerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.busyerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue)|[basicType](/javascript/api/excel/excel.calcerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.calcerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.calcerrorcellvalue#errorSubType)|Représente le type de `CalcErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.calcerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.calcerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes)|[licenseAddress](/javascript/api/excel/excel.cellvalueattributionattributes#licenseAddress)|Représente une URL vers une licence ou une source qui décrit comment cette propriété peut être utilisée.|
||[licenseText](/javascript/api/excel/excel.cellvalueattributionattributes#licenseText)|Représente un nom pour la licence qui régit cette propriété.|
||[sourceAddress](/javascript/api/excel/excel.cellvalueattributionattributes#sourceAddress)|Représente une URL vers la source du `CellValue` .|
||[sourceText](/javascript/api/excel/excel.cellvalueattributionattributes#sourceText)|Représente un nom pour la source du `CellValue` .|
|[CellValuePropertyMetadata](/javascript/api/excel/excel.cellvaluepropertymetadata)|[attribution](/javascript/api/excel/excel.cellvaluepropertymetadata#attribution)|Représente les informations d’attribution pour décrire les exigences en matière de source et de licence pour l’utilisation de cette propriété.|
||[excludeFrom](/javascript/api/excel/excel.cellvaluepropertymetadata#excludeFrom)|Représente les fonctionnalités dont cette propriété est exclue.|
||[sublabel](/javascript/api/excel/excel.cellvaluepropertymetadata#sublabel)|Représente la sous-identité de cette propriété affichée en affichage carte.|
|[CellValuePropertyMetadataExclusions](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions)|[autoComplete](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#autoComplete)|True représente que la propriété est exclue des propriétés affichées par la mise à jour automatique.|
||[calcCompare](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#calcCompare)|True représente que la propriété est exclue des propriétés utilisées pour comparer les valeurs des cellules lors du recalc.|
||[cardView](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#cardView)|True représente que la propriété est exclue des propriétés affichées en affichage carte.|
||[dotNotation](/javascript/api/excel/excel.cellvaluepropertymetadataexclusions#dotNotation)|True représente que la propriété est exclue des propriétés accessibles via la fonction FIELDVALUE.|
|[CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes)|[description](/javascript/api/excel/excel.cellvalueproviderattributes#description)|Représente la propriété de description du fournisseur utilisée en affichage carte si aucun logo n’est spécifié.|
||[logoSourceAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoSourceAddress)|Représente une URL utilisée pour télécharger une image qui sera utilisée comme logo en affichage carte.|
||[logoTargetAddress](/javascript/api/excel/excel.cellvalueproviderattributes#logoTargetAddress)|Représente une URL qui est la cible de navigation si l’utilisateur clique sur l’élément logo en affichage carte.|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que personne assignée.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Obtient la tâche associée à ce commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Obtient la tâche associée à ce commentaire.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Affecte la tâche liée au commentaire à l’utilisateur donné en tant que seule personne assignée.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Obtient la tâche associée au fil de discussion de cette réponse de commentaire.|
|[ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue)|[basicType](/javascript/api/excel/excel.connecterrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.connecterrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.connecterrorcellvalue#errorSubType)|Représente le type de `ConnectErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.connecterrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.connecterrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue)|[basicType](/javascript/api/excel/excel.div0errorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.div0errorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorType](/javascript/api/excel/excel.div0errorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.div0errorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[assignees](/javascript/api/excel/excel.documenttask#assignees)|Renvoie une collection de personnes assignées à la tâche.|
||[modifications](/javascript/api/excel/excel.documenttask#changes)|Obtient les enregistrements de modification de la tâche.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Obtient le commentaire associé à la tâche.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Obtient l’utilisateur le plus récent à avoir effectué la tâche.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Obtient la date et l’heure de fin de la tâche.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Obtient l’utilisateur qui a créé la tâche.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Obtient la date et l’heure de création de la tâche.|
||[id](/javascript/api/excel/excel.documenttask#id)|Obtient l’ID de la tâche.|
||[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Spécifie le pourcentage d’achèvement de la tâche.|
||[priority](/javascript/api/excel/excel.documenttask#priority)|Spécifie la priorité de la tâche.|
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
|[DoubleCellValue](/javascript/api/excel/excel.doublecellvalue)|[basicType](/javascript/api/excel/excel.doublecellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.doublecellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.doublecellvalue#type)|Représente le type de cette valeur de cellule.|
|[EmptyCellValue](/javascript/api/excel/excel.emptycellvalue)|[basicType](/javascript/api/excel/excel.emptycellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.emptycellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.emptycellvalue#type)|Représente le type de cette valeur de cellule.|
|[EntityCellValue](/javascript/api/excel/excel.entitycellvalue)|[basicType](/javascript/api/excel/excel.entitycellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.entitycellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[properties: { [key: string]: CellValue & { propertyMetadata](/javascript/api/excel/excel.entitycellvalue#properties)|Représente les propriétés de cette entité et leurs métadonnées.|
||[propertyMetadata](/javascript/api/excel/excel.entitycellvalue#propertyMetadata)||
||[text](/javascript/api/excel/excel.entitycellvalue#text)|Représente le texte affiché lorsqu’une cellule avec cette valeur est affichée.|
||[type](/javascript/api/excel/excel.entitycellvalue#type)|Représente le type de cette valeur de cellule.|
|[FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue)|[basicType](/javascript/api/excel/excel.fielderrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.fielderrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.fielderrorcellvalue#errorSubType)|Représente le type de `FieldErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.fielderrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.fielderrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue)|[basicType](/javascript/api/excel/excel.formattednumbercellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.formattednumbercellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[numberFormat](/javascript/api/excel/excel.formattednumbercellvalue#numberFormat)|Renvoie la chaîne de format numérique utilisée pour afficher cette valeur.|
||[type](/javascript/api/excel/excel.formattednumbercellvalue#type)|Représente le type de cette valeur de cellule.|
|[GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue)|[basicType](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.gettingdataerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorType](/javascript/api/excel/excel.gettingdataerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.gettingdataerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
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
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Effectue une demande d’actualisation du type de données liées.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Effectue une demande de modification du mode d’actualisation pour ce type de données liées.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|ID unique du type de données liées.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Renvoie un tableau avec tous les modes d’actualisation pris en charge par le type de données liées.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|ID unique du nouveau type de données liées.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtient la source de l’événement.|
||[type](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtient le type de l’événement.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Obtient le nombre de types de données liées dans la collection.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Obtient un type de données liées par ID de service.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Obtient un type de données liées par son index dans la collection.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Obtient un type de données liées par ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtient l’élément enfant chargé dans cette collection de sites.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Effectue une demande d’actualisation de tous les types de données liées dans la collection.|
|[NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue)|[basicType](/javascript/api/excel/excel.nameerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.nameerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorType](/javascript/api/excel/excel.nameerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.nameerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Obtient une vue de feuille à l’aide de son nom.|
|[NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue)|[basicType](/javascript/api/excel/excel.notavailableerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.notavailableerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorType](/javascript/api/excel/excel.notavailableerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.notavailableerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue)|[basicType](/javascript/api/excel/excel.nullerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.nullerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorType](/javascript/api/excel/excel.nullerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.nullerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue)|[basicType](/javascript/api/excel/excel.numerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.numerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorType](/javascript/api/excel/excel.numerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.numerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Obtient une cellule unique dans le tableau croisé dynamique basé sur une hiérarchie de données ainsi que les éléments de ligne et de colonne de leurs hiérarchies respectives.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|Style appliqué au tableau croisé dynamique.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Définit le style appliqué au tableau croisé dynamique.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[getDataSourceString()](/javascript/api/excel/excel.pivottable#getDataSourceString__)|Renvoie la représentation sous la chaîne de la source de données pour le tableau croisé dynamique.|
||[getDataSourceType()](/javascript/api/excel/excel.pivottable#getDataSourceType__)|Obtient le type de la source de données pour le tableau croisé dynamique.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Obtient le premier tableau croisé dynamique de la collection.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Renvoie un objet qui représente la plage contenant tous les dépendants d’une cellule dans la même feuille de calcul ou `WorkbookRangeAreas` dans plusieurs feuilles de calcul.|
|[RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue)|[basicType](/javascript/api/excel/excel.referrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.referrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.referrorcellvalue#errorSubType)|Représente le type de `RefErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.referrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
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
|[Segment](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Représente le nom du segment utilisé dans la formule.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Définit le style appliqué au slicer.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|Style appliqué au slicer.|
|[SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue)|[basicType](/javascript/api/excel/excel.spillerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.spillerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.spillerrorcellvalue#errorSubType)|Représente le type de `SpillErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.spillerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[spilledColumns](/javascript/api/excel/excel.spillerrorcellvalue#spilledColumns)|Représente le nombre de colonnes qui seraient déversées en l’absence de #SPILL ! erreur.|
||[spilledRows](/javascript/api/excel/excel.spillerrorcellvalue#spilledRows)|Représente le nombre de lignes qui seraient surdessin s’il n’y avait #SPILL ! erreur.|
||[type](/javascript/api/excel/excel.spillerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[StringCellValue](/javascript/api/excel/excel.stringcellvalue)|[basicType](/javascript/api/excel/excel.stringcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.stringcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.stringcellvalue#type)|Représente le type de cette valeur de cellule.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Modifie le tableau pour utiliser le style de tableau par défaut.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Se produit lorsqu’un filtre est appliqué à une table spécifique.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Définit le style appliqué au tableau.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|Style appliqué au tableau.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Se produit lorsqu’un filtre est appliqué à une table d’un workbook ou d’une feuille de calcul.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Obtient l’ID du tableau dans lequel le filtre est appliqué.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Obtient l’ID de la feuille de calcul qui contient le tableau.|
|[ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|[basicType](/javascript/api/excel/excel.valueerrorcellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.valueerrorcellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[errorSubType](/javascript/api/excel/excel.valueerrorcellvalue#errorSubType)|Représente le type de `ValueErrorCellValue` .|
||[errorType](/javascript/api/excel/excel.valueerrorcellvalue#errorType)|Représente le type de `ErrorCellValue` .|
||[type](/javascript/api/excel/excel.valueerrorcellvalue#type)|Représente le type de cette valeur de cellule.|
|[ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue)|[basicType](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[type](/javascript/api/excel/excel.valuetypenotavailablecellvalue#type)|Représente le type de cette valeur de cellule.|
|[WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue)|[adresse](/javascript/api/excel/excel.webimagecellvalue#address)|Représente l’URL à partir de laquelle l’image sera téléchargée.|
||[altText](/javascript/api/excel/excel.webimagecellvalue#altText)|Représente le texte de remplacement qui peut être utilisé dans les scénarios d’accessibilité pour décrire ce que représente l’image.|
||[attribution](/javascript/api/excel/excel.webimagecellvalue#attribution)|Représente les informations d’attribution pour décrire les exigences en matière de source et de licence pour l’utilisation de cette image.|
||[basicType](/javascript/api/excel/excel.webimagecellvalue#basicType)|Représente la valeur renvoyée par une `Range.valueTypes` cellule avec cette valeur.|
||[basicValue](/javascript/api/excel/excel.webimagecellvalue#basicValue)|Représente la valeur renvoyée par une `Range.values` cellule avec cette valeur.|
||[fournisseur](/javascript/api/excel/excel.webimagecellvalue#provider)|Représente des informations qui décrivent l’entité ou la personne qui a fourni l’image.|
||[relatedImagesAddress](/javascript/api/excel/excel.webimagecellvalue#relatedImagesAddress)|Représente l’URL d’une page web avec des images qui sont considérées comme étant liées à `WebImageCellValue` ce .|
||[type](/javascript/api/excel/excel.webimagecellvalue#type)|Représente le type de cette valeur de cellule.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Renvoie une collection de types de données liées qui font partie du manuel.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Spécifie si le volet liste des champs du tableau croisé dynamique est affiché au niveau du workbook.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Renvoie une collection de tâches qui sont présentes dans le workbook.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True si le classeur utilise le calendrier depuis 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Se produit lorsqu’un filtre est appliqué sur une feuille de calcul spécifique.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Renvoie une collection de tâches présentes dans la feuille de calcul.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Insère les feuilles de calcul spécifiées d’un classeur dans le classeur actif.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Se produit lorsqu’un filtre de la feuille de calcul est appliqué dans le classeur.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtient le type de l’événement.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Obtient l’ID de la feuille de calcul dans laquelle le filtre est appliqué.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[allowEditRanges](/javascript/api/excel/excel.worksheetprotection#allowEditRanges)|Spécifie les `AllowEditRangeCollection` trouvés dans cette feuille de calcul.|
||[canPauseProtection](/javascript/api/excel/excel.worksheetprotection#canPauseProtection)|Spécifie si la protection peut être suspendue pour cette feuille de calcul.|
||[checkPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#checkPassword_password_)|Spécifie si le mot de passe peut être utilisé pour déverrouiller la protection de feuille de calcul.|
||[isPasswordProtected](/javascript/api/excel/excel.worksheetprotection#isPasswordProtected)|Spécifie si la feuille est protégée par mot de passe.|
||[isPaused](/javascript/api/excel/excel.worksheetprotection#isPaused)|Spécifie si la protection de feuille de calcul est suspendue.|
||[pauseProtection(password?: string)](/javascript/api/excel/excel.worksheetprotection#pauseProtection_password_)|Suspend la protection de feuille de calcul pour l’objet de feuille de calcul donné pour l’utilisateur dans une session donnée.|
||[resumeProtection()](/javascript/api/excel/excel.worksheetprotection#resumeProtection__)|Reprend la protection de feuille de calcul pour l’objet de feuille de calcul donné pour l’utilisateur dans une session donnée.|
||[setPassword(password?: string)](/javascript/api/excel/excel.worksheetprotection#setPassword_password_)|Modifie le mot de passe associé à `WorksheetProtection` l’objet.|
||[updateOptions(options: Excel. WorksheetProtectionOptions)](/javascript/api/excel/excel.worksheetprotection#updateOptions_options_)|Modifiez les options de protection de feuille de calcul associées à `WorksheetProtection` l’objet.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[allowEditRangesChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#allowEditRangesChanged)|Spécifie si l’un des `AllowEditRange` objets a changé.|
||[protectionOptionsChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#protectionOptionsChanged)|Spécifie si les changements `WorksheetProtectionOptions` ont été faits.|
||[sheetPasswordChanged](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#sheetPasswordChanged)|Indique si le mot de passe de la feuille de calcul a changé.|

## <a name="see-also"></a>Voir aussi

- [Documentation référence de l’API JavaScript pour Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Ensembles de conditions requises de l’API JavaScript pour Excel](excel-api-requirement-sets.md)
