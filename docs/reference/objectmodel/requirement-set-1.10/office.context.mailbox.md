---
title: Office.context.mailbox - ensemble de conditions requises 1.10
description: Outlook’ensemble de conditions requises de l’API de boîte aux lettres version 1.10 du modèle objet Mailbox.
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# <a name="mailbox-requirement-set-110"></a>boîte aux lettres (ensemble de conditions requises 1.10)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)| Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

## <a name="properties"></a>Propriétés

| Propriété | Minimum<br>niveau d’autorisation | Modes | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-diagnostics-member) | ReadItem | Composition<br>Lecture | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-ewsurl-member) | ReadItem | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restreint | Composition<br>Lecture | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-mastercategories-member) | ReadWriteMailbox | Composition<br>Lecture | [Catégoriesmaître](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.10&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-resturl-member) | ReadItem | Composition<br>Lecture | Chaîne | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-userprofile-member) | ReadItem | Composition<br>Lecture | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.10&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimum<br>niveau d’autorisation | Modes | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) | ReadItem | Composition<br>Lecture | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttoewsid-member(1)) | Restreint | Composition<br>Lecture | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttorestid-member(1)) | Restreint | Composition<br>Lire | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | ReadItem | Composition<br>Lire | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1)) | ReadItem | Composition<br>Lecture | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1)) | ReadItem | Composition<br>Lecture | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | ReadItem | Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1)) | ReadItem | Lire | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewmessageform-member(1)) | ReadItem | Lire | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1)) | ReadItem | Lecture | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | ReadItem | Composition<br>Lecture | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(2)) | ReadItem | Composition<br>Lecture | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | ReadWriteMailbox | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) | ReadItem | Composition<br>Lecture | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Événements

Abonnez-vous aux événements suivants et supprimez-les à l’aide de [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) et [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) , respectivement.

> [!IMPORTANT]
> Les événements sont uniquement disponibles avec l’implémentation du volet Des tâches.

| [Event](/javascript/api/office/office.eventtype?view=outlook-js-1.10&preserve-view=true) | Description | Minimum<br>ensemble de conditions requises |
|---|---|:---:|
|`ItemChanged`| Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
