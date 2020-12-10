---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 2f83f81dcf7aa7ab0e3a48fff4279c1e08ba6286
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49612749"
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

L’ensemble de conditions requises pour l’aperçu inclut toutes les fonctionnalités de l' [ensemble de conditions requises 1,9](../requirement-set-1.9/outlook-requirement-set-1.9.md).

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

### <a name="integration-with-actionable-messages"></a>Intégration avec les messages actionnables

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le Web (moderne)

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

Cette fonctionnalité permet à votre complément d’inclure un message de notification avec une action personnalisée en plus de l’action **Ignorer** par défaut. Dans la session moderne Outlook sur le Web, cette fonctionnalité est disponible uniquement en mode composition.

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
