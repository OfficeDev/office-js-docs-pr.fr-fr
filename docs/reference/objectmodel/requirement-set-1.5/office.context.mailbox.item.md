---
title: Office.context.mailbox.item - ensemble de conditions requises 1.5
description: La version 1,5 de l’API de boîte aux lettres Outlook du modèle objet d’élément.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 6290527cb4df550f71e6ffb76d95a4e4477d4471
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891382"
---
# <a name="item-mailbox-requirement-set-15"></a>élément (boîte aux lettres requise Set 1,5)

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
| pièces jointes | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#bcc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#body) | [Corps](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#cc) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#cc) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#end)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#itemtype) | [MailboxEnums. ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#location)<br>(Demande de réunion) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#optionalattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#optionalattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#requiredattendees) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#requiredattendees) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| expéditeur | ReadItem | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| start | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#start)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| à | ReadItem | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#to) | [Destinataires](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#to) | Tableau. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimale<br>niveau d’autorisation | Détails par mode | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restreint | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities () | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restreint | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (nom) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches () | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (nom) | ReadItem | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType, [options], rappel) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant à un rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.5#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Lecture de message](/javascript/api/outlook/office.messageread?view=outlook-js-1.5#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.5#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composition de message](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.5#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

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
