---
title: Office. Context. Mailbox-ensemble de conditions requises 1,6
description: Ensemble de conditions requises de l’API de boîte aux lettres Outlook 1,6 du modèle objet boîte aux lettres.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 9473ce4d0407736623f0338822efb887a16b55c3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612065"
---
# <a name="mailbox-requirement-set-16"></a>Mailbox (ensemble de conditions requises 1,6)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)| Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

## <a name="properties"></a>Propriétés

| Propriété | Minimale<br>niveau d’autorisation | Modes | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| [Diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#diagnostics) | ReadItem | Composition<br>Read | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#ewsurl) | ReadItem | Composition<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restreint | Composition<br>Read | [Élément](/javascript/api/outlook/office.item?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#resturl) | ReadItem | Composition<br>Read | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5#userprofile) | ReadItem | Composition<br>Read | [Profil](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimale<br>niveau d’autorisation | Modes | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Composition<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttoewsid-itemid--restversion-) | Restreint | Composition<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime (timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttolocalclienttime-timevalue-) | ReadItem | Composition<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId (itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttorestid-itemid--restversion-) | Restreint | Composition<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime (entrée)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#converttoutcclienttime-input-) | ReadItem | Composition<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displayappointmentform-itemid-) | ReadItem | Composition<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaymessageform-itemid-) | ReadItem | Composition<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaynewappointmentform-parameters-) | ReadItem | Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm (paramètres)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#displaynewmessageform-parameters-) | ReadItem | Composition<br>Read | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getcallbacktokenasync-options--callback-) | ReadItem | Composition<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getcallbacktokenasync-callback--usercontext-) | ReadItem | Composition<br>Read | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Composition<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Composition<br>Read | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#removehandlerasync-eventtype--options--callback-) | ReadItem | Composition<br>Read | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Évènements

Vous pouvez vous abonner et annuler l’abonnement aux événements suivants à l’aide de [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#addhandlerasync-eventtype--handler--options--callback-) et [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6#removehandlerasync-eventtype--options--callback-) , respectivement.

| Événement | Description | Minimale<br>ensemble de conditions requises |
|---|---|:---:|
|`ItemChanged`| Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
