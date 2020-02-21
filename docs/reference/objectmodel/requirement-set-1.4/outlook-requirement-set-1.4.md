---
title: Ensemble de conditions requises de l’API du complément Outlook 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: cb4c8eecd63604aa633ade1a40eb5391b3a62ef2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165404"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Ensemble de conditions requises de l’API du complément Outlook 1.4

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-14"></a>Nouveautés de la version 1.4

L’ensemble de conditions requises de la version 1.4 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Il comprend en plus l’accès à l’espace de noms `Office.ui`.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) : Affiche une boîte de dialogue dans un hôte Office.
- Ajout de la méthode[Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Remet un message de la part de la boîte de dialogue à sa page parent/d’ouverture.
- Ajout de l’objet [Dialog](/javascript/api/office/office.dialog): objet renvoyé lorsque la méthode [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)est appelée.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
