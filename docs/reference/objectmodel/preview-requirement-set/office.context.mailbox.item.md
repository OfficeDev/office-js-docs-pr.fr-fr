---
title: 'Office.context.mailbox.item : ensemble de conditions requises d’aperçu'
description: Outlook Version de l’ensemble de conditions requises de l’API de boîte aux lettres du modèle objet Item.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: ca4afe583022761a764783293d5c036a3eda6ce7
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681513"
---
# <a name="item-mailbox-preview-requirement-set"></a>élément (ensemble de conditions requises d’aperçu de boîte aux lettres)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément à l’aide de la `itemType` propriété.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)|Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)|Organisateur de rendez-vous, participant à un rendez-vous,<br>Composition de message ou lecture de message|

> [!IMPORTANT]
> Android et iOS : il existe des restrictions sur le moment où les applications sont activées et les API disponibles. Pour plus d’informations, reportez-vous à [Ajouter une prise en charge mobile à un complément Outlook](../../../outlook/add-mobile-support.md#compose-mode-and-appointments).

## <a name="properties"></a>Propriétés

| Propriété | Minimum<br>niveau d’autorisation | Détails par mode | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| pièces jointes | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-preview&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-preview&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#bcc) | [Destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#body) | [Corps](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#cc) | [Destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-preview&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#conversationId) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#conversationId) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#dateTimeCreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#dateTimeCreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#dateTimeModified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#dateTimeModified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| delayDeliveryTime | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#delayDeliveryTime) | [DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true) | [Aperçu](outlook-requirement-set-preview.md) |
| end | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#end)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| de | ReadWriteItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#from) | [From](/javascript/api/outlook/office.from?view=outlook-js-preview&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#internetHeaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-preview&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#internetMessageId) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| isAllDayEvent | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#isAllDayEvent) | [IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true) | [Aperçu](outlook-requirement-set-preview.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent) | Valeur booléenne | [Aperçu](outlook-requirement-set-preview.md) |
| itemClass | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#itemClass) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#itemId) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| emplacement | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) | [Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#location)<br>(Demande de réunion) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-preview&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-preview&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-preview&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-preview&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalAttendees) | [Destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#optionalattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-preview&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#organizer) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-preview&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-preview&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-preview&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#recurrence)<br>(Demande de réunion) | [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-preview&preserve-view=true) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredAttendees) | [Destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#requiredAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-preview&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sensitivity | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#sensitivity) | [Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true) | [Aperçu](outlook-requirement-set-preview.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity) | [MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true) | [Aperçu](outlook-requirement-set-preview.md) |
| seriesId | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#seriesId) | Chaîne | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#seriesId) | Chaîne | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#seriesId) | Chaîne | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| sessionData | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#sessionData) | [SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-preview&preserve-view=true) | [1.11](../requirement-set-1.11/outlook-requirement-set-1.11.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#sessionData) | [SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-preview&preserve-view=true) | [1.11](../requirement-set-1.11/outlook-requirement-set-1.11.md) |
| start | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#start)<br>(Demande de réunion) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sujet | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#subject) | [Sujet](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#subject) | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| au | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#to) | [Destinataires](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-preview&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimum<br>niveau d’autorisation | Détails par mode | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restreint | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| disableClientSignatureAsync([options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#disableClientSignatureAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| displayReplyAllForm(formData) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync(formData, [options], [callback]) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync(formData, [options], [callback]) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync([options], [callback]) | ReadItem | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getAllInternetHeadersAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync(attachmentId, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync([options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getComposeTypeAsync([options], callback) | ReadItem | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| getEntities() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restreint | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getInitializationContextAsync([options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getInitializationContextAsync_options__callback_) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| getItemIdAsync([options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync([options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| isClientSignatureEnabledAsync([options], callback) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#isClientSignatureEnabledAsync_options__callback_) | [1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Participant au rendez-vous](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [Message lu](/javascript/api/outlook/office.messageread?view=outlook-js-preview&preserve-view=true#removeHandlerAsync_eventType__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [Organisateur de rendez-vous](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [Composer un message](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>Événements

Vous pouvez vous abonner aux événements suivants et les désabonner à l’aide et `addHandlerAsync` `removeHandlerAsync` respectivement.

> [!IMPORTANT]
> Les événements sont uniquement disponibles avec l’implémentation du volet Des tâches.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) | Description | Minimum<br>ensemble de conditions requises |
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
