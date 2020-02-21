---
title: Ensemble de conditions requises de l’API du complément Outlook 1.2
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: d46b705c79283049b3dbdff19b8348aa1b3c7bb0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163849"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Ensemble de conditions requises de l’API du complément Outlook 1.2

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente. 

## <a name="whats-new-in-12"></a>Nouveautés de la version 1.2

L’ensemble de conditions requises de la version 1.2 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Désormais, les compléments peuvent insérer du texte au niveau du curseur de l’utilisateur, soit dans l’objet ou le corps du message.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): retourne de façon asynchrone des données sélectionnées à partir de l’objet ou du corps d’un message.
- Ajout de la méthode [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods) : insère les données dans le corps ou l’objet d’un message de manière asynchrone.
- Modification de la fonction [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.
- Modification de la fonction [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods) : Ajout de la propriété `attachments` dans le paramètre `formData`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
