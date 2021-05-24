---
title: Ensemble de conditions requises de l’API du complément Outlook 1.6
description: Fonctionnalités et API introduites pour les Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1.6.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: cdb39eae387035f386a59b4640448b0bef25031e
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590994"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Ensemble de conditions requises de l’API du complément Outlook 1.6

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-16"></a>Nouveautés de la version 1.6

L’ensemble de conditions requises 1.6 inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Les fonctionnalités suivantes ont été ajoutées.

- Les nouvelles APIs Ajoutées pour les compléments contextuels pour que l’entité ou l’expression régulière corresponde avec l’utilisateur sélectionné pour activer le complément.
- La nouvelles API ajoutée pour ouvrir un nouveau formulaire de message.
- La possibilité ajoutée pour le complément afin de déterminer le type de compte de boîte aux lettres de l’utilisateur.

### <a name="change-log"></a>Journal des modifications

- [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods) ajouté: ajout d’une fonction qui obtient les entités figurant dans une correspondance en surbrillance sélectionnée par un utilisateur. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods) ajouté: ajout d’une fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods)-Ajout d’une nouvelle fonction qui ouvre un nouveau formulaire de message.
- [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#accounttype) ajouté: ajout d’un nouveau membre dans le profil d’utilisateur qui indique le type de compte d’utilisateur.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
