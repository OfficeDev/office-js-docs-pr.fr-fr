---
title: Office.context – ensemble de conditions requises 1.5
description: Outlook Ensemble de conditions requises de l’API de boîte aux lettres version 1.5 du modèle objet Mailbox.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: af407bd797e5a873031482bf7efefc7b41267f28
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893616"
---
# <a name="mailbox-requirement-set-15"></a>boîte aux lettres (ensemble de conditions requises 1.5)

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
| [diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#diagnostics) | ReadItem | Composition<br>Lecture | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.5&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#ewsurl) | ReadItem | Composition<br>Lecture | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restreint | Composition<br>Lecture | [Élément](/javascript/api/outlook/office.item?view=outlook-js-1.5&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#resturl) | ReadItem | Composition<br>Lecture | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#userprofile) | ReadItem | Composition<br>Lecture | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.5&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimum<br>niveau d’autorisation | Modes | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Composition<br>Lecture | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#converttoewsid-itemid--restversion-) | Restreint | Composition<br>Lecture | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#converttolocalclienttime-timevalue-) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#converttorestid-itemid--restversion-) | Restreint | Composition<br>Lecture | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#converttoutcclienttime-input-) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#displayappointmentform-itemid-) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#displaymessageform-itemid-) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#displaynewappointmentform-parameters-) | ReadItem | Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#getcallbacktokenasync-options--callback-) | ReadItem | Composition<br>Lecture | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#getcallbacktokenasync-callback--usercontext-) | ReadItem | Composition<br>Lecture | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#removehandlerasync-eventtype--options--callback-) | ReadItem | Composition<br>Lecture | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Événements

Vous pouvez vous abonner aux événements suivants et les désabonner à l’aide de [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) et [removeHandlerAsync,](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true#removehandlerasync-eventtype--options--callback-) respectivement.

> [!IMPORTANT]
> Les événements sont uniquement disponibles avec l’implémentation du volet Des tâches.

| [Event](/javascript/api/office/office.eventtype) | Description | Minimum<br>ensemble de conditions requises |
|---|---|:---:|
|`ItemChanged`| Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
