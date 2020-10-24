---
title: Ensemble de conditions requises de l’API du complément Outlook 1,9
description: Ensemble de conditions requises 1,9 pour l’API de complément Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: b2174052a60580a895ef82a4b5d8f00ed6899feb
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628052"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Ensemble de conditions requises de l’API du complément Outlook 1,9

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.

## <a name="whats-new-in-19"></a>Quelles sont les nouveautés de 1,9 ?

L’ensemble de conditions requises 1,9 inclut toutes les fonctionnalités de l' [ensemble de conditions requises 1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouvelles API pour les fonctionnalités ajout d’envoi, propriétés personnalisées et formulaire d’affichage.
- Prise en charge supplémentaire de `Dialog.messageChild` .

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [CustomProperties. GetAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--): ajoute une nouvelle fonction à l' `CustomProperties` objet qui obtient toutes les propriétés personnalisées.
- Ajout de la méthode [Dialog. messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): ajoute une nouvelle méthode qui remet un message à partir de la page hôte, telle qu’un volet de tâches ou un fichier de fonctions sans interface utilisateur, à une boîte de dialogue ouverte à partir de la page.
- Ajout de l' [élément de manifeste ExtendedPermissions](../../manifest/extendedpermissions.md): ajoute un élément enfant à l’élément de manifeste [VersionOverrides](../../manifest/versionoverrides.md) . Pour qu’un complément prenne en charge la [fonctionnalité Append-on-Send](../../../outlook/append-on-send.md), l' `AppendOnSend` autorisation étendue doit être incluse dans la collection des autorisations étendues.
- Ajout de la méthode [Office. Context. Mailbox. displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un rendez-vous existant. Il s’agit de la version asynchrone de la `displayAppointmentForm` méthode.
- Ajout de la méthode [Office. Context. Mailbox. displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un message existant. Il s’agit de la version asynchrone de la `displayMessageForm` méthode.
- Ajout de la méthode [Office. Context. Mailbox. displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de rendez-vous. Il s’agit de la version asynchrone de la `displayNewAppointmentForm` méthode.
- Ajout de la méthode [Office. Context. Mailbox. displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-): ajoute une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de message. Il s’agit de la version asynchrone de la `displayNewMessageForm` méthode.
- Ajout de la méthode [Office. Context. Mailbox. Item. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-): ajoute une nouvelle fonction à l' `Body` objet qui ajoute des données à la fin du corps de l’élément en mode composition.
- Ajout de la méthode [Office. Context. Mailbox. Item. displayReplyAllFormAsync](office.context.mailbox.item.md#methods): ajoute une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre à tous » en mode lecture. Il s’agit de la version asynchrone de la `displayReplyAllForm` méthode.
- Ajout de la méthode [Office. Context. Mailbox. Item. displayReplyFormAsync](office.context.mailbox.item.md#methods): ajoute une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre » en mode lecture. Il s’agit de la version asynchrone de la `displayReplyForm` méthode.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
