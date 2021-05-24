---
title: Ensemble de conditions requises de l’API du complément Outlook 1.5
description: Fonctionnalités et API introduites pour les Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1.5.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7d780538a77f54db6f1234a6d29a3bcdea9533b0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590840"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Ensemble de conditions requises de l’API du complément Outlook 1.5

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-15"></a>Nouveautés de la version 1.5

L’ensemble de conditions requises 1.5 inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). Les fonctionnalités suivantes ont été ajoutées.

- Prise en charge des [volets Office épinglables](../../../outlook/pinnable-taskpane.md).
- Prise en charge de l’appel des [API REST](../../../outlook/use-rest-api.md).
- Possibilité de marquer une pièce jointe comme élément incorporé.
- Possibilité de fermer un volet Office ou une boîte de dialogue.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods) : ajoute un gestionnaire d’événements pour un événement pris en charge.
- Ajout [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): supprime les handlers d’événements pour un type d’événement pris en charge.
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
