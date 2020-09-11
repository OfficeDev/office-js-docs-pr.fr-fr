---
title: Ensemble de conditions requises de l’API du complément Outlook 1.1
description: Les fonctionnalités et les API qui ont été introduites pour les compléments Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1,1.
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f93b6d582043641903b362121c6e5eaf89c2ad1c
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431373"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Ensemble de conditions requises de l’API du complément Outlook 1.1

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook. L’API JavaScript pour Outlook 1,1 (boîte aux lettres 1,1) est la première version de l’API.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-11"></a>Nouveautés de la version 1.1

L’ensemble de conditions requises 1,1 inclut tous les [ensembles de conditions requises d’API communs](../../requirement-sets/office-add-in-requirement-sets.md) pris en charge dans Outlook. Désormais, les compléments peuvent accéder au corps des messages et des rendez-vous et vous pouvez modifier l’élément actif.

### <a name="change-log"></a>Journal des modifications

- Ajout de l’objet [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1&preserve-view=true) : fournit des méthodes pour ajouter et mettre à jour le contenu d’un élément dans un complément Outlook.
- Ajout de l’objet [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1&preserve-view=true) : Fournit des méthodes pour obtenir et définir le lieu d’une réunion dans un complément Outlook.
- Ajout de l’objet [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1&preserve-view=true) : fournit des méthodes pour obtenir et définir les destinataires d’un rendez-vous ou d’un message dans un complément Outlook.
- Ajout de l’objet [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1&preserve-view=true) : Fournit des méthodes pour obtenir et définir l’objet d’un rendez-vous ou d’un message dans un complément Outlook.
- Ajout de l’objet [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1&preserve-view=true) : fournit des méthodes pour obtenir et définir l’heure de début ou de fin d’une réunion dans un complément Outlook.
- Ajout de la méthode [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods) : ajoute un fichier à un message ou un rendez-vous en pièce jointe.
- Ajout de la méthode [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods) : ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.
- Ajout de la méthode [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods) : supprime une pièce jointe d’un message ou d’un rendez-vous.
- Ajout de l’objet [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties) : obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.
- Ajout de la ligne [Office. Context. Mailbox. Item. BCC](office.context.mailbox.item.md#properties) d’un message.
- Ajout de l’énumération [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true) : spécifie le type de destinataire d’un rendez-vous.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
