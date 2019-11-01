---
title: Ensemble de conditions requises de l’API du complément Outlook 1.6
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 46d1b4eeb260c2b0f3b94999a7f02a1384b71942
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902080"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Ensemble de conditions requises de l’API du complément Outlook 1.6

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-16"></a>Nouveautés de la version 1.6

L’ensemble de conditions requises de la version 1.6 comprend toutes les fonctionnalités de l’[Ensemble de conditions requises de la version 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). Les fonctionnalités suivantes ont été ajoutées.

- Les nouvelles APIs Ajoutées pour les compléments contextuels pour que l’entité ou l’expression régulière corresponde avec l’utilisateur sélectionné pour activer le complément.
- La nouvelles API ajoutée pour ouvrir un nouveau formulaire de message.
- La possibilité ajoutée pour le complément afin de déterminer le type de compte de boîte aux lettres de l’utilisateur.

### <a name="change-log"></a>Journal des modifications

- [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities) ajouté: ajout d’une fonction qui obtient les entités figurant dans une correspondance en surbrillance sélectionnée par un utilisateur. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object) ajouté: ajout d’une fonction qui renvoie les valeurs de chaîne dans une correspondance en surbrillance qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux compléments contextuels.
- [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)-Ajout d’une nouvelle fonction qui ouvre un nouveau formulaire de message.
- [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string) ajouté: ajout d’un nouveau membre dans le profil d’utilisateur qui indique le type de compte d’utilisateur.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](/outlook/add-ins/quick-start)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
