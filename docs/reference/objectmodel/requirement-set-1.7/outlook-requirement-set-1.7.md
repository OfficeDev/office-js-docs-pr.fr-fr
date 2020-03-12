---
title: Ensemble de conditions requises de l’API du complément Outlook 1.7
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 793a1e1c2c3dd014f104ab264f4954369b591162
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597025"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Ensemble de conditions requises de l’API du complément Outlook 1.7

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-17"></a>Nouveautés de la version 1.7

L’ensemble de conditions requises de la version 1.7 comprend toutes les fonctionnalités de l’[Ensemble de conditions requises de la version 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Les fonctionnalités suivantes ont été ajoutées.

- Nouvelles API ajoutées concernant la périodicité sur les messages et rendez-vous qui sont des demandes de réunion.
- Modification de la propriété item.from pour également être disponibles en mode Composer.
- Prise en charge ajoutée pour les événements RecurrenceChanged, RecipientsChanged et AppointmentTimeChanged.

### <a name="change-log"></a>Journal des modifications

- Ajout de la fonctionnalité [De](/javascript/api/outlook/office.from?view=outlook-js-1.7) : ajoute un nouvel objet qui fournit une méthode pour obtenir la valeur « De ».
- Ajout de [Organisateur](/javascript/api/outlook/office.organizer?view=outlook-js-1.7): ajoute un nouvel objet qui fournit une méthode pour obtenir la valeur Organisateur.
- Ajout de [Périodicité](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7): ajoute un nouvel objet qui fournit des méthodes permettant d’obtenir et configurer la périodicité de rendez-vous, mais obtient uniquement la périodicité de messages qui sont des demandes de réunion.
- Ajout de [RecurrenceTimeZone](/javascript/api/outlook/office.recurrencetimezone?view=outlook-js-1.7): ajoute un nouvel objet qui représente la configuration de fuseau horaire de la périodicité.
- Ajout de [SeriesTime](/javascript/api/outlook/office.seriestime?view=outlook-js-1.7): ajoute un nouvel objet qui fournit des méthodes pour obtenir et définir les dates et heures de rendez-vous dans une série périodique et consulter les dates et heures de demandes de réunion dans une série périodique.
- Ajout de [Office.context.mailbox.addHandlerAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle méthode qui ajoute un gestionnaire d’événements pour un événement pris en charge.
- Modifié [Office.context.mailbox.item.from](office.context.mailbox.item.md#properties): ajoute la possibilité d’obtenir la valeur « à partir de » en mode de composition.
- Modifié [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#properties): ajoute la possibilité d’obtenir la valeur « organisateur » en mode de composition.
- Ajout de [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#properties): ajoute une nouvelle propriété qui obtient ou définit un objet qui fournit des méthodes pour gérer la périodicité d’un élément de rendez-vous. Cette propriété peut également être utilisée pour obtenir la périodicité d’un élément de demande de réunion.
- Ajout de [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle méthode qui supprime le gestionnaire d’événements pour un type d’événement pris en charge.
- Ajout de [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#properties): ajoute une nouvelle propriété qui récupère l’Id de la série à laquelle une occurrence appartient.
- Ajout de [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days?view=outlook-js-1.7): ajoute une nouvelle énumération qui spécifie le jour de semaine ou le type de journée.
- Ajout de [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month?view=outlook-js-1.7): ajoute une nouvelle énumération qui spécifie le mois.
- Ajout de [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone?view=outlook-js-1.7): ajoute une nouvelle énumération qui spécifie le fuseau horaire appliqué à la périodicité.
- Ajout de [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype?view=outlook-js-1.7): ajoute une nouvelle énumération qui spécifie le type de périodicité.
- Ajout de [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber?view=outlook-js-1.7): ajoute une nouvelle énumération qui spécifie la semaine du mois.
- Modifié [Office.EventType](/javascript/api/office/office.eventtype): ajoute la prise en charge des événements `RecurrenceChanged`, `RecipientsChanged` et `AppointmentTimeChanged`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
