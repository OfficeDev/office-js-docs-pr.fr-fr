---
title: Ensemble de conditions requises de l’API du complément Outlook 1.3
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: ba74102d70d2f2c38e96cb833b779d3381595f0f
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814863"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Ensemble de conditions requises de l’API du complément Outlook 1.3

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-13"></a>Nouveautés de la version 1.3

L’ensemble de conditions requises de la version 1.3 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [commandes de complément](/outlook/add-ins/add-in-commands-for-outlook).
- Possibilité d’enregistrer ou de fermer un élément en cours de composition.
- Objet [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3) amélioré pour permettre aux compléments d’obtenir ou de définir le corps entier.
- Nouvelles méthodes de conversion pour convertir les ID aux formats EWS et REST.
- Possibilité d’ajouter des messages de notification dans la barre d’informations sur les éléments.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-) : Renvoie le corps actif dans un format spécifié.
- Ajout de la méthode [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-) : Remplace l’ensemble du corps avec le texte spécifié.
- Ajout de l’objet [Event](/javascript/api/office/office.addincommands.event) : transmis comme paramètre aux fonctions de commande sans IU dans un complément Outlook. Utilisé pour signaler la fin du traitement de l’événement.
- Ajout de la méthode [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods) : Ferme l’élément en cours qui est composé.
- Ajout de la méthode [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods) : Enregistre un élément de manière asynchrone.
- Ajout de l’objet [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties) : Obtient les messages de notification pour un élément.
- Ajout de la méthode [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods) : Convertit un ID d’élément mis en forme pour REST au format EWS.
- Ajout de la méthode [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods) : Convertit un ID d’élément mis en forme pour EWS au format REST.
- Ajout de l’énumération [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3) : Spécifie le type de message de notification pour un rendez-vous ou un message.
- Ajout de l’énumération [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3) : Spécifie la version de l’API REST qui correspond à un ID d’élément au format REST.
- Ajout de l’objet [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3) : fournit des méthodes pour accéder aux messages de notification dans un complément Outlook.
- Ajout du type [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3) : renvoyé par la méthode `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](/outlook/add-ins/quick-start)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
