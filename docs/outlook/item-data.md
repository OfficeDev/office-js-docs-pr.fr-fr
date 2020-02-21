---
title: Obtenir ou définir des données d’élément dans un complément Outlook
description: Selon qu’un complément est activé dans un formulaire de lecture ou de composition, les propriétés disponibles pour le complément sur un élément diffèrent.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: be7d14a6c417d01c0537e3375524da5cc807d749
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166111"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition

À compter de la version 1.1 du schéma de manifeste des compléments Office, Outlook peut activer les compléments lorsque l’utilisateur affiche ou compose un élément. Selon qu’un complément est activé dans un formulaire de lecture ou de composition, les propriétés disponibles pour le complément sur l’élément diffèrent également.

Par exemple, les propriétés [dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) sont définies uniquement pour un élément qui a déjà été envoyé (l’élément est affiché par la suite dans un formulaire de lecture), mais pas lorsque l’élément est en cours de création (dans un formulaire de composition). Un autre exemple est la propriété [Cci](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) qui est pertinente uniquement lorsqu’un message est en cours de création (dans un formulaire de composition) et n’est pas accessible à l’utilisateur dans un formulaire de lecture.

## <a name="item-properties-available-in-compose-and-read-forms"></a>Propriétés d’éléments disponibles dans les formulaires de composition et de lecture

Le Tableau 1 indique les propriétés au niveau de l’élément dans l’interface API JavaScript pour Office qui sont disponibles dans mode (lecture et écriture) des compléments de messagerie. En règle générale, ces propriétés disponibles dans les formulaires de lecture sont en lecture seule et celles disponibles dans les formulaires de composition sont en lecture/écriture, à l’exception des propriétés [itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), [conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) qui sont toujours en lecture seule.

Pour les propriétés restantes au niveau de l’élément disponibles dans les formulaires de composition, étant donné que le complément et l’utilisateur peuvent lire ou écrire la même propriété simultanément, les méthodes pour les obtenir ou les définir dans le mode de composition sont asynchrones et par conséquent, les types des objets renvoyés par ces propriétés peuvent également être différents dans les formulaires de compositions et les formulaires de lecture. Pour plus d’informations sur l’utilisation des méthodes asynchrones pour obtenir ou définir des propriétés au niveau de l’élément en mode de composition, reportez-vous à [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md).


**Tableau 1. Propriétés d’éléments disponibles dans les formulaires de composition et de lecture**

<br/>

|**Type d’élément**|**Propriété**|**Type de propriété dans les formulaires de lecture**|**Type de propriété dans les formulaires de composition**|
|:-----|:-----|:-----|:-----|
|Rendez-vous et messages|[dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objet **Date** JavaScript|Propriété non disponible|
|Rendez-vous et messages|[dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objet **Date** JavaScript|Propriété non disponible|
|Rendez-vous et messages|[itemClass](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|Propriété non disponible|
|Rendez-vous et messages|[itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|Propriété non disponible|
|Rendez-vous et messages|[itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Chaîne dans l’énumération [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)|Chaîne dans l’énumération [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) (lecture seule)|
|Rendez-vous et messages|[attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|Propriété non disponible|
|Rendez-vous et messages|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[Corps](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|Rendez-vous et messages|[normalizedSubject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|Propriété non disponible|
|Rendez-vous et messages|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|[Subject](/javascript/api/outlook/office.subject)|
|Rendez-vous|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objet **Date** JavaScript|[Heure](/javascript/api/outlook/office.time)|
|Rendez-vous|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|[Emplacement](/javascript/api/outlook/office.location)|
|Rendez-vous|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|
|Rendez-vous|[organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)|
|Rendez-vous|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|
|Rendez-vous|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Objet **Date** JavaScript|[Time](/javascript/api/outlook/office.time)|
|Messages|[bbc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Propriété non disponible|[Destinataires](/javascript/api/outlook/office.recipients)|
|Messages|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|
|Messages|[conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|String (lecture seule)|
|Messages|[from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)|
|Messages|[internetMessageId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Entier|Propriété non disponible|
|Messages|[sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|Propriété non disponible|
|Messages|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>Utilisation de jetons de rappel Exchange Server à partir d’un complément de lecture

Si votre complément Outlook est activé dans les formulaires de lecture, vous pouvez obtenir un jeton de rappel Exchange. Ce jeton peut être utilisé dans le code côté serveur pour accéder à l’élément complet via les services web Exchange (EWS).

En spécifiant l’autorisation **ReadItem** dans le manifeste du complément, vous pouvez utiliser la méthode [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour obtenir un jeton de rappel Exchange, la propriété [mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) pour obtenir l’URL du point de terminaison EWS de la boîte aux lettres de l’utilisateur, et [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour obtenir l’ID EWS de l’élément sélectionné. Vous pouvez ensuite transmettre le jeton de rappel, l’URL du point de terminaison EWS et l’ID d’élément EWS dans le code côté serveur pour accéder à l’opération [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) afin d’obtenir d’autres propriétés de l’élément.


## <a name="access-ews-from-a-read-or-compose-add-in"></a>Accès à EWS à partir d’un complément de composition ou de lecture

Vous pouvez également utiliser la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour accéder aux opérations des services web Exchange [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) et [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) directement à partir du complément. Vous pouvez utiliser ces opérations pour obtenir et définir de nombreuses propriétés d’un élément spécifié. Cette méthode est disponible pour les compléments Outlook que le complément ait été activé ou non dans un formulaire de lecture ou de composition, tant que vous spécifiez l’autorisation **ReadWriteMailbox** dans le manifeste de complément.

Pour plus d’informations sur l’utilisation de **makeEwsRequestAsync** pour accéder aux opérations EWS, reportez-vous à [Appeler des services web à partir d’un complément Outlook](web-services.md).


## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Appeler des services web à partir d’un complément Outlook](web-services.md)
