---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,9
description: La version 1,9 de l’API de boîte aux lettres Outlook du modèle objet d’élément.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: af8faeb9ba652e880b5c7bf293145a5289ad671b
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628056"
---
# <a name="item-mailbox-requirement-set-19"></a>élément (boîte aux lettres requise Set 1,9)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément à l’aide de la `itemType` propriété.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)|Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)|Organisateur de rendez-vous, participant au rendez-vous,<br>Composition de message ou lecture de message|

## <a name="properties"></a>Propriétés

| Propriété | Minimum<br>niveau d’autorisation | Détails par mode | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| pièces jointes | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#bcc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#cc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#cc) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#end)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| de | ReadWriteItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#from) | [From](/javascript/api/outlook/office.from) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#location)<br>(Demande de réunion) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#optionalattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#optionalattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#recurrence) | [Instances](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#recurrence) | [Instances](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#recurrence)<br>(Demande de réunion) | [Instances](/javascript/api/outlook/office.recurrence) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#requiredattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#requiredattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| expéditeur | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#seriesid) | String | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#start)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| au | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#to) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#to) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimum<br>niveau d’autorisation | Détails par mode | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restreint | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync (Formulairemode, [options], [Rappel]) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayreplyallformasync-formdata--options--callback-) | [1,9](outlook-requirement-set-1.9.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayreplyallformasync-formdata--options--callback-) | [1,9](outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync (Formulairemode, [options], [Rappel]) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayreplyformasync-formdata--options--callback-) | [1,9](outlook-requirement-set-1.9.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayreplyformasync-formdata--options--callback-) | [1,9](outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync ([options], [Rappel]) | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getallinternetheadersasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync (attachmentId, [options], [Rappel]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync ([options], [Rappel]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getattachmentsasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getEntities () | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restreint | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (nom) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getItemIdAsync ([options], rappel) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getitemidasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches () | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (nom) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType, [options], rappel) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getselectedentities--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getselectedregexmatches--) | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync ([options], rappel) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1,7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Événements

Vous pouvez vous abonner aux événements suivants et les annuler, à l’aide de `addHandlerAsync` et `removeHandlerAsync` respectivement.

| Événement | Description | Minimum<br>ensemble de conditions requises |
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
