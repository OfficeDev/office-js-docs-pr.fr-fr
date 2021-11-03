---
title: Outlook l’ensemble de conditions requises de l’API du add-in 1.11
description: Ensemble de conditions requises 1.11 pour Outlook API de votre application.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 56066d7b3a6debaeed365a9ca05a3e894762dea3
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681776"
---
# <a name="outlook-add-in-api-requirement-set-111"></a>Outlook l’ensemble de conditions requises de l’API du add-in 1.11

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.

## <a name="whats-new-in-111"></a>Nouveautés de la 1.11

L’ensemble de conditions requises 1.11 inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouveaux événements pour [l’activation basée sur des événements.](../../../outlook/autolaunch.md#supported-events)
- Ajout des API SessionData.

### <a name="change-log"></a>Journal des modifications

- Ajout [Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties): ajoute une nouvelle propriété pour gérer les données de session d’un élément en mode composition.
- Ajout de [Office. SessionData :](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true)ajoute un nouvel objet qui représente les données de session d’un élément de composition.
- Ajout de nouveaux événements pour l’activation basée [sur des](../../../outlook/autolaunch.md#supported-events)événements : ajoute la prise en charge des événements suivants.

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- Ajout de [Office. AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true): ajoute un objet qui prend en charge `OnAppointmentTimeChanged` l’événement.
- Ajout [Office. AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true): ajoute un objet qui prend en charge les `OnAppointmentAttachmentsChanged` événements et les `OnMessageAttachmentsChanged` événements.
- Ajout [Office. InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true): ajoute un objet qui prend en charge `OnInfoBarDismissClicked` l’événement.
- Ajout [Office. RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true): ajoute un objet qui prend en charge les `OnAppointmentAttendeesChanged` événements et les `OnMessageRecipientsChanged` événements.
- Ajout [Office. RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true): ajoute un objet qui prend en charge `OnAppointmentRecurrenceChanged` l’événement.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
