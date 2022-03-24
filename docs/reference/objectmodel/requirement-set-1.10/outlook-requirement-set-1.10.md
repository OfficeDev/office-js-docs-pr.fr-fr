---
title: Outlook l’ensemble de conditions requises de l’API du add-in 1.10
description: Ensemble de conditions requises 1.10 pour Outlook API de votre application.
ms.date: 11/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: a7412c655423d016101b406444f86da0f2be610d
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746526"
---
# <a name="outlook-add-in-api-requirement-set-110"></a>Outlook l’ensemble de conditions requises de l’API du add-in 1.10

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook de gestion.

> [!NOTE]
> Dans cette documentation, l’[ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) présenté est différent de l’ensemble de conditions requises de la version précédente.

## <a name="whats-new-in-110"></a>Nouveautés de la 1.10

L’ensemble de conditions requises 1.10 inclut toutes les fonctionnalités de l’ensemble [de conditions requises 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md). Les fonctionnalités suivantes ont été ajoutées.

- Ajout de nouvelles API pour [l’activation basée sur des](../../../outlook/autolaunch.md) événements et les fonctionnalités de signature électronique.
- Prise en charge supplémentaire de [l’objet OfficeRuntime.Stockage](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true) avec la fonctionnalité d’activation basée sur des événements.
- Ajout de la possibilité d’inclure une action personnalisée sur un message de notification.

### <a name="change-log"></a>Journal des modifications

- Ajout [d’un point d’extension LaunchEvent](../../manifest/extensionpoint.md#launchevent) : ajoute un nouveau type d’ExtensionPoint pris en charge. Il configure la fonctionnalité d’activation basée sur des événements.
- Ajout de [l’élément de manifeste LaunchEvents](../../manifest/launchevents.md) : ajoute un élément manifeste pour prendre en charge la configuration de la fonctionnalité d’activation basée sur les événements.
- Élément [manifeste Runtimes modifié](../../manifest/runtimes.md) : ajoute Outlook prise en charge. Il fait référence aux fichiers HTML et JavaScript nécessaires pour la fonctionnalité d’activation basée sur des événements.
- Ajout [Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#outlook-office-body-setsignatureasync-member(1)) : ajoute une nouvelle fonction à l’objet`Body`. Il ajoute ou remplace la signature dans le corps de l’élément en mode Composition.
- Ajout de [Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle fonction qui désactive la signature du client pour la boîte aux lettres d’envoi en mode composition.
- Ajout [Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#outlook-office-messagecompose-getcomposetypeasync-member(1)) : ajoute une nouvelle fonction qui obtient le type de composition d’un message en mode composition.
- Ajout de [Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods) : ajoute une nouvelle fonction qui vérifie si la signature du client est activée sur l’élément en mode composition.
- Ajout de [Office. MailboxEnums.ActionType :](/javascript/api/outlook/office.mailboxenums.actiontype?view=outlook-js-1.10&preserve-view=true) ajoute une nouvelle enum. Il représente le type d’action personnalisée dans un message de notification.
- Ajout de [Office. MailboxEnums.ComposeType :](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true) ajoute une nouvelle enum disponible en mode composition.
- Ajout de [Office. MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.10&preserve-view=true) : ajoute un nouveau type à l’enum`ItemNotificationMessageType`. Il représente un message de notification avec une action personnalisée.
- Ajout de [Office. NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction?view=outlook-js-1.10&preserve-view=true) : ajoute un nouvel objet afin que vous pouvez définir une action personnalisée pour votre `InsightMessage` notification.
- Ajout de [Office. NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.10&preserve-view=true#outlook-office-notificationmessagedetails-actions-member) : `InsightMessage` ajoute une nouvelle propriété qui vous permet d’ajouter une notification avec une action personnalisée.
- Modification [d’OfficeRuntime.Stockage](/javascript/api/office-runtime/officeruntime.storage?view=outlook-js-1.10&preserve-view=true) : ajoute Outlook prise en charge, mais uniquement avec la fonctionnalité d’activation basée sur des événements.

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
