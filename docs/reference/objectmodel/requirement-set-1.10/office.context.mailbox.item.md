---
title: Office.context.mailbox.item - ensemble de conditions requises 1.10
description: Outlook Ensemble de conditions requises de l’API de boîte aux lettres version 1.10 du modèle objet Item.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 6869be63694e2680325b6236eeee40d9f0f43d1a
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592039"
---
# <a name="item-mailbox-requirement-set-110"></a>élément (ensemble de conditions requises de boîte aux lettres 1.10)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément à l’aide de la `itemType` propriété.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)|Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)|Organisateur de rendez-vous, participant à un rendez-vous,<br>Composition de message ou lecture de message|

## <a name="properties"></a>Propriétés

| Propriété | Minimum<br>niveau d’autorisation | Détails par mode | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| pièces jointes | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#bcc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#cc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#conversationid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#conversationid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#end)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#internetmessageid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemclass) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemclass) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| emplacement | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#location) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#location)<br>(Demande de réunion) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#normalizedsubject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#normalizedsubject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#optionalattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#optionalattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#recurrence) | [Périodicité](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#recurrence) | [Périodicité](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#recurrence)<br>(Demande de réunion) | [Périodicité](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#requiredattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#requiredattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| expéditeur | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#seriesid) | Chaîne | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#seriesid) | Chaîne | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#seriesid) | Chaîne | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#seriesid) | Chaîne | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#start)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| au | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#to) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimum<br>niveau d’autorisation | Détails par mode | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restreint | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| disableClientSignatureAsync([options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#disableclientsignatureasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#disableclientsignatureasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| displayReplyAllForm(formData) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync(formData, [options], [callback]) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyallformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyallformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync(formData, [options], [callback]) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#displayreplyformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#displayreplyformasync-formdata--options--callback-) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync([options], [callback]) | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getallinternetheadersasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync(attachmentId, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync([options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getComposeTypeAsync([options], callback) | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getcomposetypeasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| getEntities() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restreint | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getItemIdAsync([options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync([options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| isClientSignatureEnabledAsync([options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#isclientsignatureenabledasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#isclientsignatureenabledasync-options--callback-) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-1.10&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.10&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Événements

Vous pouvez vous abonner aux événements suivants et les désabonner à l’aide et `addHandlerAsync` `removeHandlerAsync` respectivement.

> [!IMPORTANT]
> Les événements sont uniquement disponibles avec l’implémentation du volet Des tâches.

| Événement | Description | Minimum<br>ensemble de conditions requises |
|---|---|:---:|
|`AppointmentTimeChanged`| La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci. | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| L’emplacement du rendez-vous sélectionné a changé. | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`RecipientsChanged`| La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| La périodicité de la série sélectionnée a été modifiée. | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

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
