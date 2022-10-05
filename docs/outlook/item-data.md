---
title: Obtenir ou définir des données d’élément dans un complément Outlook
description: Selon qu’un complément est activé dans un formulaire de lecture ou de composition, les propriétés disponibles pour le complément sur un élément diffèrent.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8349d81b376aa55d239a88a5d4598381fd8bfc4d
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467271"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition

À compter de la version 1.1 du schéma de manifeste des compléments Office, Outlook peut activer les compléments lorsque l’utilisateur affiche ou compose un élément. Selon qu’un complément est activé dans un formulaire de lecture ou de composition, les propriétés disponibles pour le complément sur l’élément diffèrent également.

Par exemple, les propriétés [dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) et [dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) sont définies uniquement pour un élément qui a déjà été envoyé (l’élément est affiché par la suite dans un formulaire de lecture), mais pas lorsque l’élément est en cours de création (dans un formulaire de composition). Un autre exemple est la propriété [Cci](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) qui est pertinente uniquement lorsqu’un message est en cours de création (dans un formulaire de composition) et n’est pas accessible à l’utilisateur dans un formulaire de lecture.

## <a name="item-properties-available-in-compose-and-read-forms"></a>Propriétés d’éléments disponibles dans les formulaires de composition et de lecture

Le tableau 1 présente les propriétés au niveau de l’élément dans l’API JavaScript Office qui sont disponibles dans chaque mode (lecture et composition) des compléments de messagerie. En règle générale, les propriétés disponibles dans les formulaires de lecture sont en lecture seule, et celles disponibles dans les formulaires de composition sont en lecture/écriture, à l’exception des propriétés [itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), [conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) et [itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) , qui sont toujours en lecture seule indépendamment.

Pour les propriétés restantes au niveau de l’élément disponibles dans les formulaires de composition, étant donné que le complément et l’utilisateur peuvent lire ou écrire la même propriété simultanément, les méthodes pour les obtenir ou les définir dans le mode de composition sont asynchrones et par conséquent, les types des objets renvoyés par ces propriétés peuvent également être différents dans les formulaires de compositions et les formulaires de lecture. Pour plus d’informations sur l’utilisation des méthodes asynchrones pour obtenir ou définir des propriétés au niveau de l’élément en mode de composition, reportez-vous à [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md).


**Tableau 1. Propriétés d’éléments disponibles dans les formulaires de composition et de lecture**

<br/>

|**Type d’élément**|**Propriété**|**Type de propriété dans les formulaires de lecture**|**Type de propriété dans les formulaires de composition**|
|:-----|:-----|:-----|:-----|
|Rendez-vous et messages|[dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objet **Date** JavaScript|Propriété non disponible|
|Rendez-vous et messages|[dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objet **Date** JavaScript|Propriété non disponible|
|Rendez-vous et messages|[itemClass](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Propriété non disponible|
|Rendez-vous et messages|[itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Propriété non disponible|
|Rendez-vous et messages|[itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Chaîne dans l’énumération [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)|Chaîne dans l’énumération [ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) (en lecture seule)|
|Rendez-vous et messages|[attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|Propriété non disponible|
|Rendez-vous et messages|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Corps](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|Rendez-vous et messages|[normalizedSubject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Propriété non disponible|
|Rendez-vous et messages|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|[Subject](/javascript/api/outlook/office.subject)|
|Rendez-vous|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objet **Date** JavaScript|[Heure](/javascript/api/outlook/office.time)|
|Rendez-vous|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|[Emplacement](/javascript/api/outlook/office.location)|
|Rendez-vous|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|
|Rendez-vous|[organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)|
|Rendez-vous|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|
|Rendez-vous|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Objet **Date** JavaScript|[Time](/javascript/api/outlook/office.time)|
|Messages|[bbc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Propriété non disponible|[Destinataires](/javascript/api/outlook/office.recipients)|
|Messages|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|
|Messages|[conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|Chaîne (lecture seule)|
|Messages|[from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)|
|Messages|[internetMessageId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Entier|Propriété non disponible|
|Messages|[sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|Propriété non disponible|
|Messages|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Destinataires](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>Utilisation de jetons de rappel Exchange Server à partir d’un complément de lecture

Si votre complément Outlook est activé dans les formulaires de lecture, vous pouvez obtenir un jeton de rappel Exchange. Ce jeton peut être utilisé dans le code côté serveur pour accéder à l’élément complet via les services web Exchange (EWS).

En spécifiant [l’autorisation d’élément de lecture](understanding-outlook-add-in-permissions.md#read-item-permission) dans le manifeste du complément, vous pouvez utiliser la méthode [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) pour obtenir un jeton de rappel Exchange, la propriété [mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) pour obtenir l’URL du point de terminaison EWS pour la boîte aux lettres de l’utilisateur et [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) pour obtenir l’ID EWS de l’élément sélectionné. Vous pouvez ensuite transmettre le jeton de rappel, l’URL du point de terminaison EWS et l’ID d’élément EWS dans le code côté serveur pour accéder à l’opération [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) afin d’obtenir d’autres propriétés de l’élément.

## <a name="access-ews-from-a-read-or-compose-add-in"></a>Accès à EWS à partir d’un complément de composition ou de lecture

Vous pouvez également utiliser la méthode [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) pour accéder aux opérations des services web Exchange [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) et [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) directement à partir du complément. Vous pouvez utiliser ces opérations pour obtenir et définir de nombreuses propriétés d’un élément spécifié. Cette méthode est disponible pour les compléments Outlook, que le complément ait été activé dans un formulaire de lecture ou de composition, tant que vous spécifiez l’autorisation de **boîte aux lettres en lecture/écriture** dans le manifeste du complément. Pour plus d’informations sur l’autorisation de **boîte aux lettres en lecture/écriture** , consultez [Présentation des autorisations de complément Outlook](understanding-outlook-add-in-permissions.md)

Pour plus d’informations sur l’utilisation de **makeEwsRequestAsync** pour accéder aux opérations EWS, reportez-vous à [Appeler des services web à partir d’un complément Outlook](web-services.md).


## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Appeler des services web à partir d’un complément Outlook](web-services.md)
