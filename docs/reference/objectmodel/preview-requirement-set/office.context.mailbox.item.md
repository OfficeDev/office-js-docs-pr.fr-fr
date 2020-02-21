---
title: Office. Context. Mailbox. Item-Preview ensemble de conditions requises
description: ''
ms.date: 02/03/2020
localization_priority: Normal
ms.openlocfilehash: d4a65fd4e53d86d39c6b2e5bcd2145e32ece2225
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163856"
---
# <a name="item"></a>élément

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément à l’aide `itemType` de la propriété.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)|Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)|Organisateur de rendez-vous, participant au rendez-vous,<br>Composition de message ou lecture de message|

## <a name="properties"></a>Propriétés

| Propriété | Minimale<br>niveau d’autorisation | Détails par mode | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| pièces jointes | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#bcc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| corps | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#cc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#cc) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#conversationid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#conversationid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#end)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#from) | [From](/javascript/api/outlook/office.from) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#internetmessageid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemclass) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemclass) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| emplacement | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) | [Emplacement](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#location) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#location)<br>(Demande de réunion) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#normalizedsubject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#normalizedsubject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#optionalattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#recurrence) | [Instances](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#recurrence) | [Instances](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#recurrence)<br>(Demande de réunion) | [Instances](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#requiredattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| expéditeur | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#seriesid) | Chaîne | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#seriesid) | Chaîne | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#seriesid) | Chaîne | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#seriesid) | Chaîne | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#start)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| à | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#to) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#to) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimale<br>niveau d’autorisation | Détails par mode | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| addFileAttachmentAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| fermer | Restreint | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getAllInternetHeadersAsync | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getallinternetheadersasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getEntities | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType | Restreint | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getInitializationContextAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| getItemIdAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| loadCustomPropertiesAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Événements

Vous pouvez vous abonner aux événements suivants et les annuler, `addHandlerAsync` à `removeHandlerAsync` l’aide de et respectivement.

| Événement | Description | Minimale<br>ensemble de conditions requises |
|---|---|:---:|
|`AppointmentTimeChanged`| La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci. | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| L’emplacement du rendez-vous sélectionné a changé. | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`RecipientsChanged`| La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| La périodicité de la série sélectionnée a été modifiée. | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

## <a name="example"></a>Exemple

L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```
