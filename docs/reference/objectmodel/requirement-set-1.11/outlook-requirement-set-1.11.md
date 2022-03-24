---
title: Outlook l’ensemble de conditions requises de l’API du add-in 1.11
description: Ensemble de conditions requises 1.11 pour Outlook API de votre application.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 384e872b44b213b60a1b651f85ac315cd06cf082
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744143"
---
# <a name="outlook-add-in-api-requirement-set-111"></a>Outlook l’ensemble de conditions requises de l’API du add-in 1.11

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook de gestion.

## <a name="whats-new-in-111"></a>Nouveautés de la 1.11

L’ensemble de conditions requises 1.11 inclut toutes les fonctionnalités de l’ensemble [de conditions requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouveaux événements pour [l’activation basée sur des événements](../../../outlook/autolaunch.md#supported-events).
- Ajout des API SessionData.

### <a name="change-log"></a>Journal des modifications

- Ajout [Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties) : ajoute une nouvelle propriété pour gérer les données de session d’un élément en mode composition.
- Ajout de [Office. SessionData :](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true) ajoute un nouvel objet qui représente les données de session d’un élément de composition.
- Ajout de nouveaux événements pour [l’activation basée sur des](../../../outlook/autolaunch.md#supported-events) événements : ajoute la prise en charge des événements suivants.

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- Ajout de [Office. AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true) : ajoute un objet qui prend en charge l’événement`OnAppointmentTimeChanged`.
- Ajout de [Office. AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true) : ajoute un objet qui prend en charge les événements `OnAppointmentAttachmentsChanged` et les `OnMessageAttachmentsChanged` événements.
- Ajout de [Office. InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true) : ajoute un objet qui prend en charge l’événement`OnInfoBarDismissClicked`.
- Ajout de [Office. RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true) : ajoute un objet qui prend en charge les événements `OnAppointmentAttendeesChanged` et les `OnMessageRecipientsChanged` événements.
- Ajout de [Office. RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true) : ajoute un objet qui prend en charge l’événement`OnAppointmentRecurrenceChanged`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
