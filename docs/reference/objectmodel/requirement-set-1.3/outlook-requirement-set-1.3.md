---
title: Ensemble de conditions requises de l’API du complément Outlook 1.3
description: Les fonctionnalités et les API qui ont été introduites pour les compléments Outlook et les API JavaScript Office dans le cadre de l’API de boîte aux lettres 1,3.
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: bf7dc9e3977df626241cdafdebd8d4b4e473d494
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430967"
---
# <a name="outlook-add-in-api-requirement-set-13"></a>Ensemble de conditions requises de l’API du complément Outlook 1.3

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-13"></a>Nouveautés de la version 1.3

L’ensemble de conditions requises de la version 1.3 comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). Les fonctionnalités suivantes ont été ajoutées :

- Prise en charge des [commandes de complément](../../../outlook/add-in-commands-for-outlook.md).
- Possibilité d’enregistrer ou de fermer un élément en cours de composition.
- Objet [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true) amélioré pour permettre aux compléments d’obtenir ou de définir le corps entier.
- Nouvelles méthodes de conversion pour convertir les ID aux formats EWS et REST.
- Possibilité d’ajouter des messages de notification dans la barre d’informations sur les éléments.

### <a name="change-log"></a>Journal des modifications

- Ajout de la méthode [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#getasync-coerciontype--options--callback-) : Renvoie le corps actif dans un format spécifié.
- Ajout de la méthode [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3&preserve-view=true#setasync-data--options--callback-) : Remplace l’ensemble du corps avec le texte spécifié.
- Ajout de l’objet [Event](/javascript/api/office/office.addincommands.event) : transmis comme paramètre aux fonctions de commande sans IU dans un complément Outlook. Utilisé pour signaler la fin du traitement de l’événement.
- Ajout de la méthode [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods) : Ferme l’élément en cours qui est composé.
- Ajout de la méthode [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods) : Enregistre un élément de manière asynchrone.
- Ajout de l’objet [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties) : Obtient les messages de notification pour un élément.
- Ajout de la méthode [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods) : Convertit un ID d’élément mis en forme pour REST au format EWS.
- Ajout de la méthode [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods) : Convertit un ID d’élément mis en forme pour EWS au format REST.
- Ajout de l’énumération [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3&preserve-view=true) : Spécifie le type de message de notification pour un rendez-vous ou un message.
- Ajout de l’énumération [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3&preserve-view=true) : Spécifie la version de l’API REST qui correspond à un ID d’élément au format REST.
- Ajout de l’objet [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3&preserve-view=true) : fournit des méthodes pour accéder aux messages de notification dans un complément Outlook.
- Ajout du type [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3&preserve-view=true) : renvoyé par la méthode `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
