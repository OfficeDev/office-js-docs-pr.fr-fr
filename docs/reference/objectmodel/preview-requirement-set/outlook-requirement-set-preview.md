---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Fonctionnalités et API actuellement en prévisualisation pour les add-ins Outlook.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: 39dd1221f4dea9674c89cdaad20024ce408f8db3
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104839"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble de l’API de l’API JavaScript pour Outlook inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un application Outlook.

> [!IMPORTANT]
> Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**. Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Vous pourrez peut-être afficher un aperçu des fonctionnalités dans Outlook sur le web en configurant la version ciblée [sur votre client Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center) « Configurer l’accès à l’aperçu » est indiqué sur cette page pour les fonctionnalités applicables.
>
> Pour d’autres fonctionnalités, vous pouvez demander l’accès aux bits d’aperçu pour Outlook sur le web à l’aide de votre compte Microsoft 365 en remplissant et en envoyant [ce formulaire.](https://aka.ms/OWAPreview) « Demander l’accès en prévisualisation » est indiqué sur ces fonctionnalités.

L’ensemble de conditions requises preview inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Activation de complément sur des éléments protégés par la Gestion des droits de l’information (IRM)

Les add-ins peuvent désormais être activés sur les éléments protégés par IRM. Pour activer cette fonctionnalité, un administrateur client doit activer le droit d’utilisation en paramètres de stratégie personnalisée Autoriser l’accès par programme `OBJMODEL` dans Office.  Pour plus [d’informations, voir droits d’utilisation et descriptions.](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions)

**Disponible dans**: Outlook sur Windows, à partir de la build 13229.10000 (connecté à un abonnement Microsoft 365)

<br>

---

---

### <a name="additional-calendar-properties"></a>Propriétés de calendrier supplémentaires

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée d’un rendez-vous en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui représente la sensibilité d’un rendez-vous en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui représente si un rendez-vous est un événement d’une journée.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui représente la sensibilité d’un rendez-vous.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Ajout d’une nouvelle enum `AppointmentSensitivityType` qui représente les options de sensibilité disponibles sur un rendez-vous.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Activation basée sur un événement

Prise en charge supplémentaire de la fonctionnalité d’activation basée sur des événements dans les compléments Outlook. Pour [plus d’informations,](../../../outlook/autolaunch.md) voir Configurer votre complément Outlook pour l’activation basée sur des événements.

#### <a name="launchevent-extension-point"></a>[Point d’extension LaunchEvent](../../manifest/extensionpoint.md#launchevent-preview)

Ajout de `LaunchEvent` la prise en charge du point d’extension au manifeste. Il configure la fonctionnalité d’activation basée sur des événements.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="launchevents-manifest-element"></a>[Élément manifeste LaunchEvents](../../manifest/launchevents.md)

Ajout `LaunchEvents` d’un élément au manifeste. Il prend en charge la configuration de la fonctionnalité d’activation basée sur des événements.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="runtimes-manifest-element"></a>[Élément manifeste Runtimes](../../manifest/runtimes.md)

Ajout de la prise en charge d’Outlook à `Runtimes` l’élément manifeste. Il fait référence aux fichiers HTML et JavaScript nécessaires pour la fonctionnalité d’activation basée sur des événements.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Intégration avec les messages actionnables

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

<br>

---

---

### <a name="mail-signature"></a>Signature électronique

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

Ajout d’une nouvelle fonction à l’objet qui ajoute ou remplace la signature dans le corps de l’élément `Body` en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction qui désactive la signature du client pour la boîte aux lettres d’envoi en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction qui vérifie si la signature du client est activée sur l’élément en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officemailboxenumscomposetype"></a>[Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

Ajout d’une nouvelle enum `ComposeType` disponible en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne, Configurer l’accès [en prévisualisation)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

<br>

---

---

### <a name="notification-messages-with-actions"></a>Messages de notification avec actions

Cette fonctionnalité permet à votre add-in d’inclure un message de notification avec une action personnalisée en plus de l’action d’ignorer **par** défaut. Dans Outlook sur le web moderne, cette fonctionnalité est disponible en mode composition uniquement.

#### <a name="officenotificationmessagedetailsactions"></a>[Office.NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

Ajout d’une nouvelle propriété qui vous permet d’ajouter une `InsightMessage` notification avec une action personnalisée.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

#### <a name="officenotificationmessageaction"></a>[Office.NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

Ajout d’un nouvel objet dans lequel vous définissez une action personnalisée pour votre `InsightMessage` notification.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

#### <a name="officemailboxenumsactiontype"></a>[Office.MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

Ajout d’une nouvelle enum `ActionType` .

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office.MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Ajout d’un nouveau type `InsightMessage` à `ItemNotificationMessageType` l’enum.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

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

#### <a name="officesessiondata"></a>[Office.SessionData](/javascript/api/outlook/office.sessiondata)

Ajout d’un nouvel objet qui représente les données de session d’un élément.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365)

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
