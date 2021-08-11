---
title: Ensemble de conditions requises de l’API du complément Outlook 1.1
description: Fonctionnalités et API introduites pour les Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1.1.
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 20105efd3d7e7e978f7c184c029d6482c0db8bd947166e91d9e9f5714e775d99
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098723"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Ensemble de conditions requises de l’API du complément Outlook 1.1

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook. Outlook L’API JavaScript 1.1 (Mailbox 1.1) est la première version de l’API.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-11"></a>Nouveautés de la version 1.1

L’ensemble de conditions requises 1.1 inclut tous les ensembles de conditions requises [d’API](../../requirement-sets/office-add-in-requirement-sets.md) communes pris en charge dans Outlook. Désormais, les compléments peuvent accéder au corps des messages et des rendez-vous et vous pouvez modifier l’élément actif.

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
- Ajout de [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) d’un message.
- Ajout de l’énumération [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1&preserve-view=true) : spécifie le type de destinataire d’un rendez-vous.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
