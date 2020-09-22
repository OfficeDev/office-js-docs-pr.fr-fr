---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 09/21/2020
localization_priority: Normal
ms.openlocfilehash: f7c9c7c2e60a77c30e3957a0c759d0f20b22e86a
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175541"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.

> [!IMPORTANT]
> Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**. Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Vous pouvez afficher un aperçu des fonctionnalités dans Outlook sur le Web en [configurant la version ciblée sur votre client Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center). « Configurer l’accès en aperçu » est indiqué sur cette page pour les fonctionnalités applicables.
>
> Pour les autres fonctionnalités, vous pouvez demander l’accès à des bits d’aperçu pour Outlook sur le Web à l’aide de votre compte Microsoft 365 en remplissant et envoyant [ce formulaire](https://aka.ms/OWAPreview). « Demander un accès en aperçu » est indiqué sur ces fonctionnalités.

L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Activation des compléments sur les éléments protégés par la gestion des droits relatifs à l’information (IRM)

Les compléments peuvent désormais être activés sur les éléments protégés par IRM. Pour activer cette fonctionnalité, un administrateur client doit activer le droit d' `OBJMODEL` utilisation en définissant l’option autoriser la stratégie personnalisée d' **accès par programme** dans Office. Pour plus d’informations [, voir droits et descriptions d’utilisation](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) .

**Disponible dans**: Outlook sur Windows, en commençant par Build 13229,10000 (connecté à un abonnement Microsoft 365)

<br>

---

---

### <a name="additional-calendar-properties"></a>Propriétés de calendrier supplémentaires

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée entière d’un rendez-vous en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui représente le critère de diffusion d’un rendez-vous en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office. Context. Mailbox. Item. isAllDayEvent](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui indique si un rendez-vous est un événement d’une journée entière.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. Context. Mailbox. Item. Sensitivity](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui représente le critère de diffusion d’un rendez-vous.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums. AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Ajout d’une nouvelle énumération `AppointmentSensitivityType` qui représente les options de critère de diffusion disponibles sur un rendez-vous.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

<br>

---

---

### <a name="append-on-send"></a>Ajouter à l'envoi

Pour en savoir plus sur l’utilisation de la fonctionnalité Ajout à l’envoi, consultez la rubrique [implémenter Append lors de l’envoi dans votre complément Outlook](../../../outlook/append-on-send.md).

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[Office. Context. Mailbox. Item. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-)

Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute des données à la fin du corps de l’élément en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

Ajout d’un nouvel élément au manifeste dans lequel l' `AppendOnSend` autorisation étendue doit être incluse dans la collection des autorisations étendues.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="async-versions-of-display-apis"></a>Versions Async des `display` API

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[Office. Context. Mailbox. displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentformasync-itemid--options--callback-)

Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un rendez-vous existant. Il s’agit de la version asynchrone de la `displayAppointmentForm` méthode.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[Office. Context. Mailbox. displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageformasync-itemid--options--callback-)

Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un message existant. Il s’agit de la version asynchrone de la `displayMessageForm` méthode.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[Office. Context. Mailbox. displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)

Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de rendez-vous. Il s’agit de la version asynchrone de la `displayNewAppointmentForm` méthode.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[Office. Context. Mailbox. displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)

Ajout d’une nouvelle fonction à l' `Mailbox` objet qui affiche un nouveau formulaire de message. Il s’agit de la version asynchrone de la `displayNewMessageForm` méthode.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[Office. Context. Mailbox. Item. displayReplyAllFormAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre à tous » en mode lecture. Il s’agit de la version asynchrone de la `displayReplyAllForm` méthode.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[Office. Context. Mailbox. Item. displayReplyFormAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction à l' `Item` objet qui affiche le formulaire « répondre » en mode lecture. Il s’agit de la version asynchrone de la `displayReplyForm` méthode.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

<br>

---

---

### <a name="event-based-activation"></a>Activation basée sur un événement

Prise en charge supplémentaire de la fonctionnalité d’activation basée sur un événement dans les compléments Outlook. Pour en savoir plus, reportez-vous à [la rubrique Configurer votre complément Outlook pour l’activation basée sur les événements](../../../outlook/autolaunch.md) .

#### <a name="launchevent-extension-point"></a>[Point d’extension LaunchEvent](../../manifest/extensionpoint.md#launchevent-preview)

Ajout `LaunchEvent` de la prise en charge du point d’extension au manifeste. Il configure les fonctionnalités d’activation basée sur les événements.

**Disponible dans**: Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="launchevents-manifest-element"></a>[Élément de manifeste LaunchEvents](../../manifest/launchevents.md)

Ajout `LaunchEvents` de l’élément à manifest. Il prend en charge la configuration de la fonctionnalité d’activation basée sur les événements.

**Disponible dans**: Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="runtimes-manifest-element"></a>[Élément de manifeste runtimes](../../manifest/runtimes.md)

Ajout de la prise en charge d’Outlook à l' `Runtimes` élément de manifeste. Il fait référence aux fichiers HTML et JavaScript nécessaires à la fonctionnalité d’activation basée sur les événements.

**Disponible dans**: Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="get-all-custom-properties"></a>Obtenir toutes les propriétés personnalisées

#### <a name="custompropertiesgetall"></a>[CustomProperties. getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true#getall--)

Ajout d’une nouvelle fonction à l' `CustomProperties` objet qui obtient toutes les propriétés personnalisées.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne), Outlook sur Mac (connecté à un abonnement Microsoft 365), Outlook sur Android, Outlook sur iOS

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Intégration avec les messages actionnables

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (classique)

<br>

---

---

### <a name="mail-signature"></a>Signature de courrier électronique

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office. Context. Mailbox. Item. Body. setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

Ajout d’une nouvelle fonction à l' `Body` objet qui ajoute ou remplace la signature dans le corps de l’élément en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office. Context. Mailbox. Item. disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Ajout d’une fonction qui désactive la signature client pour la boîte aux lettres d’envoi en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office. Context. Mailbox. Item. getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office. Context. Mailbox. Item. isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Ajout d’une fonction qui vérifie si la signature client est activée sur l’élément en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

#### <a name="officemailboxenumscomposetype"></a>[Office. MailboxEnums. ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

Ajout d’une nouvelle énumération `ComposeType` disponible en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne, [configurer l’accès en aperçu](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="notification-messages-with-actions"></a>Messages de notification avec actions

Cette fonctionnalité permet à votre complément d’inclure un message de notification avec une action personnalisée en plus de l’action **Ignorer** par défaut.

#### <a name="officenotificationmessagedetailsactions"></a>[Office. NotificationMessageDetails. actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

Ajout d’une nouvelle propriété qui vous permet d’ajouter une `InsightMessage` notification avec une action personnalisée.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officenotificationmessageaction"></a>[Office. NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

Ajout d’un nouvel objet dans lequel vous définissez une action personnalisée pour votre `InsightMessage` notification.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officemailboxenumsactiontype"></a>[Office. MailboxEnums. ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

Ajout d’une nouvelle énumération `ActionType` .

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office. MailboxEnums. ItemNotificationMessageType. InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Ajout d’un nouveau type `InsightMessage` à l' `ItemNotificationMessageType` énumération.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

<br>

---

---

### <a name="office-theme"></a>Thème Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Ajout de la possibilité d’obtenir un thème Office.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

<br>

---

---

### <a name="session-data"></a>Données de session

#### <a name="officesessiondata"></a>[Office. SessionData](/javascript/api/outlook/office.sessiondata)

Ajout d’un nouvel objet qui représente les données de session d’un élément.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office. Context. Mailbox. Item. sessionData](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
