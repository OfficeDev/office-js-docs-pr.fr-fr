---
title: Ensemble de conditions requises de l’API du complément Outlook 1.7
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2e233c614a902a724ead0240c4e5229e1053ee81
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432311"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Ensemble de conditions requises de l’API du complément Outlook 1.7

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

## <a name="whats-new-in-17"></a>Nouveautés de la version 1.7

L’ensemble de conditions requises de la version 1.7 comprend toutes les fonctionnalités de l’[Ensemble de conditions requises de la version 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Les fonctionnalités suivantes ont été ajoutées.

- Nouvelles API ajoutées concernant la périodicité sur les messages et rendez-vous qui sont des demandes de réunion.
- Modification de la propriété item.from pour également être disponibles en mode Composer.
- Prise en charge ajoutée pour les événements RecurrenceChanged, RecipientsChanged et AppointmentTimeChanged.

### <a name="change-log"></a>Journal des modifications

- Ajout de la fonctionnalité [De](/javascript/api/outlook_1_7/office.from) : ajoute un nouvel objet qui fournit une méthode pour obtenir la valeur « De ».
- Ajout de [Organisateur](/javascript/api/outlook_1_7/office.organizer): ajoute un nouvel objet qui fournit une méthode pour obtenir la valeur Organisateur.
- Ajout de [Périodicité](/javascript/api/outlook_1_7/office.recurrence): ajoute un nouvel objet qui fournit des méthodes permettant d’obtenir et configurer la périodicité de rendez-vous, mais obtient uniquement la périodicité de messages qui sont des demandes de réunion.
- Ajout de [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): ajoute un nouvel objet qui représente la configuration de fuseau horaire de la périodicité.
- Ajout de [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): ajoute un nouvel objet qui fournit des méthodes pour obtenir et définir les dates et heures de rendez-vous dans une série périodique et consulter les dates et heures de demandes de réunion dans une série périodique.
- Ajout de [Office.context.mailbox.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback) : ajoute une nouvelle méthode qui ajoute un gestionnaire d’événements pour un événement pris en charge.
- Modification de [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): modifie pour obtenir la valeur « De » en mode Composer.
- Modification de [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer): modifie pour obtenir la valeur Organisateur en mode Composer.
- Ajout de [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): ajoute une nouvelle propriété qui obtient ou définit un objet qui fournit des méthodes pour gérer la périodicité d’un élément de rendez-vous. Cette propriété peut également être utilisée pour obtenir la périodicité d’un élément de demande de réunion.
- Ajout de [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback) : ajoute une nouvelle méthode qui supprime un gestionnaire d’événements.
- Ajout de [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): ajoute une nouvelle propriété qui récupère l’Id de la série à laquelle une occurrence appartient.
- Ajout de [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): ajoute une nouvelle énumération qui spécifie le jour de semaine ou le type de journée.
- Ajout de [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): ajoute une nouvelle énumération qui spécifie le mois.
- Ajout de [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): ajoute une nouvelle énumération qui spécifie le fuseau horaire appliqué à la périodicité.
- Ajout de [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): ajoute une nouvelle énumération qui spécifie le type de périodicité.
- Ajout de [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): ajoute une nouvelle énumération qui spécifie la semaine du mois.
- Modification de [Office.EventType](/javascript/api/office/office.eventtype) : modifie la prise en charge des événements RecurrenceChanged, RecipientsChanged, et AppointmentTimeChanged via l’ajout respectivement d’entrées `RecurrenceChanged`,`RecipientsChanged` et `AppointmentTimeChanged`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](https://docs.microsoft.com/outlook/add-ins/quick-start)