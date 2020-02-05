---
title: Office. Context. Mailbox. Item-ensemble de conditions requises 1,4
description: ''
ms.date: 02/03/2020
localization_priority: Normal
ms.openlocfilehash: a904770a63f137abd0209ed06031942da0b21156
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773585"
---
# <a name="item"></a>élément

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément à l’aide `itemType` de la propriété.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|Restreinte|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Organisateur de rendez-vous, participant au rendez-vous,<br>Composition de message ou lecture de message|

## <a name="properties"></a>Propriétés

| Propriété | Minimale<br>niveau d’autorisation | Détails par mode | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| pièces jointes | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#bcc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| corps | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#cc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#cc) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#conversationid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#conversationid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#end)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#internetmessageid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#itemclass) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#itemclass) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#itemid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#itemid) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#location) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#location)<br>(Demande de réunion) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#normalizedsubject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#normalizedsubject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#optionalattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#optionalattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#requiredattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#requiredattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| start | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#start)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sujet | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| à | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#to) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#to) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimale<br>niveau d’autorisation | Détails par mode | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| addFileAttachmentAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| fermer | Restreint | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType | Restreint | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| loadCustomPropertiesAsync | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.4#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.4#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| saveAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.4#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.4#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

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
