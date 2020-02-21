---
title: Office. Context. Mailbox-ensemble de conditions requises 1,8
description: ''
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 9bf71808b538f3048b6cc5f74a1bb0b9f88bea68
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163632"
---
# <a name="mailbox"></a>boîte aux lettres

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
| [Diagnostics](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#diagnostics) | ReadItem | Composition<br>Lecture | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.8) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#ewsurl) | ReadItem | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restreint | Composition<br>Lecture | [Élément](/javascript/api/outlook/office.item?view=outlook-js-1.8) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#mastercategories) | ReadWriteMailbox | Composition<br>Lecture | [Catégoriesmaître](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8) | [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#resturl) | ReadItem | Composition<br>Lecture | Chaîne | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#userprofile) | ReadItem | Composition<br>Lire | [Profil](/javascript/api/outlook/office.userprofile?view=outlook-js-1.8) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>Méthodes

| Méthode | Minimale<br>niveau d’autorisation | Modes | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-) | ReadItem | Composition<br>Lire | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttoewsid-itemid--restversion-) | Restreint | Composition<br>Lire | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttolocalclienttime-timevalue-) | ReadItem | Composition<br>Lire | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttorestid-itemid--restversion-) | Restreint | Composition<br>Lire | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#converttoutcclienttime-input-) | ReadItem | Composition<br>Lire | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displayappointmentform-itemid-) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displaymessageform-itemid-) | ReadItem | Composition<br>Lire | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displaynewappointmentform-parameters-) | ReadItem | Lire | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewMessageForm](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#displaynewmessageform-parameters-) | ReadItem | Composition<br>Lecture | [1,6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#getcallbacktokenasync-options--callback-) | ReadItem | Composition<br>Lecture | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#getcallbacktokenasync-callback--usercontext-) | ReadItem | Composition<br>Lecture | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#getuseridentitytokenasync-callback--usercontext-) | ReadItem | Composition<br>Lecture | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#makeewsrequestasync-data--callback--usercontext-) | ReadWriteMailbox | Composition<br>Lire | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-) | ReadItem | Composition<br>Lecture | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>Événements

Vous pouvez vous abonner et annuler l’abonnement aux événements suivants à l’aide de [addHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#addhandlerasync-eventtype--handler--options--callback-) et [removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8#removehandlerasync-eventtype--options--callback-) , respectivement.

| Événement | Description | Minimale<br>ensemble de conditions requises |
|---|---|:---:|
|`ItemChanged`| Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
