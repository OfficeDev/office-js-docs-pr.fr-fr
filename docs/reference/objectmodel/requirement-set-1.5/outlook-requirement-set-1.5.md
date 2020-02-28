---
title: Ensemble de conditions requises de l’API du complément Outlook 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: e9dda19a4dd9831ca1bb7924558ac99c26baf763
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324791"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Ensemble de conditions requises de l’API du complément Outlook 1.5

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-15"></a>Nouveautés de la version 1.5

L’ensemble de conditions requises de la version 1.5 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [volets Office épinglables](../../../outlook/pinnable-taskpane.md).
- Prise en charge de l’appel des [API REST](../../../outlook/use-rest-api.md).
- Possibilité de marquer une pièce jointe comme élément incorporé.
- Possibilité de fermer un volet Office ou une boîte de dialogue.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) : ajoute un gestionnaire d’événements pour un événement pris en charge.
- Ajout de la méthode [Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#methods): supprime les gestionnaires d’événements pour un type d’événement pris en charge.
- Ajout de l’énumération [Office.EventType](office.md#eventtype-string) : spécifie l’événement associé à un gestionnaire d’événements et prend en charge l’événement ItemChanger.
- Ajout de la propriété [Office.context.mailbox.restUrl](office.context.mailbox.md#properties) : obtient l’URL du point de terminaison REST de ce compte de messagerie.
- Modification de la méthode [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods) : cette nouvelle version comprend une nouvelle signature (`getCallbackTokenAsync([options], callback)`). La version d’origine est toujours disponible et reste inchangée.
- Ajout de la méthode [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Modification de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `options` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.
- Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.
- Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) : nouvelle valeur du dictionnaire `formData.attachments` appelée `isInline`. Elle indique qu’une image est incorporée dans le corps du message.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
